using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C2B_FBR_Connect.Forms
{
    public partial class MainForm : Form
    {
        #region Fields

        private DataGridView dgvInvoices;
        private Button btnFetchInvoices, btnUploadSelected, btnUploadAll, btnGeneratePDF, btnCompanySetup, btnRefresh;
        private Label lblCompanyName, lblFilter, lblSearch;
        private ComboBox cboStatusFilter;
        private TextBox txtSearchInvoice;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel, statusStats;
        private ToolStripProgressBar progressBar;
        private Panel topPanel, buttonPanel;

        private readonly DatabaseService _db;
        private readonly QuickBooksService _qb;
        private readonly FBRApiService _fbr;
        private readonly PDFService _pdf;
        private readonly CompanyManager _companyManager;
        private readonly InvoiceManager _invoiceManager;
        private readonly TransactionTypeService _transactionTypeService;
        private readonly SroDataService _sroDataService;
        private InvoiceTrackingService _trackingService;
        private bool _excludeUploadedInvoices = true; 

        private Company _currentCompany;
        private List<Invoice> _allInvoices;
        private string _currentSortColumn;
        private bool _sortAscending = true;
        private Dictionary<int, bool> _checkboxStates = new Dictionary<int, bool>();

        #endregion

        #region Initialization

        public MainForm()
        {
            InitializeComponent();

            _db = new DatabaseService();
            _qb = new QuickBooksService();
            _fbr = new FBRApiService();
            _pdf = new PDFService();
            _companyManager = new CompanyManager(_db);
            _invoiceManager = new InvoiceManager(_db, _qb, _fbr, _pdf);
            _transactionTypeService = new TransactionTypeService(_db);
            _sroDataService = new SroDataService(_fbr, _transactionTypeService);

            SetupCustomUI();
            ConnectToQuickBooks();

            this.FormClosing += MainForm_FormClosing;
            this.Resize += (s, e) => dgvInvoices?.Refresh();
        }

        private void SetupCustomUI()
        {
            ConfigureForm();
            CreateTopPanel();
            CreateButtonPanel();
            CreateDataGridView();
            CreateStatusStrip();

            this.Controls.Add(dgvInvoices);
            this.Controls.Add(buttonPanel);
            this.Controls.Add(topPanel);
            this.Controls.Add(statusStrip);
        }

        private void ConfigureForm()
        {
            this.Text = "C2B Smart App - Digital Invoicing System";
            this.Size = new Size(1400, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1000, 600);

            try
            {
                this.Icon = new Icon(System.IO.Path.Combine(Application.StartupPath, "assets", "favicon.ico"));
            }
            catch { }
        }

        private void CreateTopPanel()
        {
            topPanel = new Panel
            {
                Height = 45,
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(12, 10, 12, 10)
            };

            lblCompanyName = new Label
            {
                Text = "No QuickBooks Company Connected",
                AutoSize = false,
                Height = 25,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                TextAlign = ContentAlignment.MiddleLeft
            };

            topPanel.Controls.Add(lblCompanyName);
        }

        private void CreateButtonPanel()
        {
            buttonPanel = new Panel
            {
                Height = 95,
                Dock = DockStyle.Top,
                BackColor = Color.White,
                Padding = new Padding(12, 10, 12, 10)
            };

            CreateFilterControls();
            CreateActionButtons();
        }

        private void CreateFilterControls()
        {
            lblFilter = new Label
            {
                Text = "Status:",
                Location = new Point(0, 5),
                Size = new Size(55, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F)
            };

            cboStatusFilter = new ComboBox
            {
                Location = new Point(60, 5),
                Size = new Size(110, 23),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F)
            };
            cboStatusFilter.Items.AddRange(new object[] { "All", "Pending", "Uploaded", "Failed" });
            cboStatusFilter.SelectedIndex = 0;
            cboStatusFilter.SelectedIndexChanged += (s, e) => ApplyFilters();

            lblSearch = new Label
            {
                Text = "Search:",
                Location = new Point(185, 5),
                Size = new Size(55, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F)
            };

            txtSearchInvoice = new TextBox
            {
                Location = new Point(245, 5),
                Size = new Size(200, 23),
                PlaceholderText = "Invoice # or Customer Name",
                Font = new Font("Segoe UI", 9F)
            };
            txtSearchInvoice.TextChanged += (s, e) => ApplyFilters();

            buttonPanel.Controls.AddRange(new Control[] { lblFilter, cboStatusFilter, lblSearch, txtSearchInvoice });
        }

        private void CreateActionButtons()
        {
            int btnY = 40;
            int btnSpacing = 8;
            int currentX = 0;

            btnFetchInvoices = CreateButton("Fetch Invoices", ref currentX, btnY, 120, Color.FromArgb(0, 122, 204), btnSpacing);
            btnFetchInvoices.Click += BtnFetchInvoices_Click;

            btnRefresh = CreateButton("Refresh", ref currentX, btnY, 90, Color.FromArgb(46, 125, 50), btnSpacing);
            btnRefresh.Click += (s, e) => LoadInvoices();

            btnUploadSelected = CreateButton("Upload Selected", ref currentX, btnY, 130, Color.FromArgb(255, 152, 0), btnSpacing);
            btnUploadSelected.Click += BtnUploadSelected_Click;

            btnGeneratePDF = CreateButton("Generate PDF", ref currentX, btnY, 120, Color.FromArgb(156, 39, 176), btnSpacing);
            btnGeneratePDF.Click += BtnGeneratePDF_Click;

            btnCompanySetup = CreateButton("Company Setup", ref currentX, btnY, 130, Color.FromArgb(96, 125, 139), 0);
            btnCompanySetup.Click += (s, e) => ShowCompanySetup();
        }

        private Button CreateButton(string text, ref int x, int y, int width, Color backColor, int spacing)
        {
            var button = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, 35),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 0 }
            };

            buttonPanel.Controls.Add(button);
            x += width + spacing;
            return button;
        }

        private void CreateDataGridView()
        {
            dgvInvoices = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToResizeRows = false,
                RowTemplate = { Height = 28 },
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 32,
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(250, 250, 250) },
                Font = new Font("Segoe UI", 9F),
                GridColor = Color.FromArgb(230, 230, 230),
                EnableHeadersVisualStyles = false
            };

            dgvInvoices.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(52, 73, 94),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Padding = new Padding(8, 6, 8, 6),
                Alignment = DataGridViewContentAlignment.MiddleLeft
            };

            dgvInvoices.CellContentClick += DgvInvoices_CellContentClick;
            dgvInvoices.CellDoubleClick += DgvInvoices_CellDoubleClick;
            dgvInvoices.ColumnHeaderMouseClick += DgvInvoices_ColumnHeaderMouseClick;
            dgvInvoices.DataBindingComplete += DgvInvoices_DataBindingComplete;
            dgvInvoices.CurrentCellDirtyStateChanged += DgvInvoices_CurrentCellDirtyStateChanged;
            dgvInvoices.DoubleBuffered(true);
        }

        private void CreateStatusStrip()
        {
            statusStrip = new StatusStrip { BackColor = Color.FromArgb(240, 240, 240) };

            statusLabel = new ToolStripStatusLabel("Ready")
            {
                Spring = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F)
            };

            statusStats = new ToolStripStatusLabel("Total: 0 | Pending: 0 | Uploaded: 0 | Failed: 0")
            {
                BorderSides = ToolStripStatusLabelBorderSides.Left,
                Padding = new Padding(10, 0, 10, 0),
                Font = new Font("Segoe UI", 9F)
            };

            progressBar = new ToolStripProgressBar
            {
                Visible = false,
                Size = new Size(150, 16)
            };

            statusStrip.Items.AddRange(new ToolStripItem[] { statusLabel, progressBar, statusStats });
        }

        #endregion

        #region Connection Management

        private bool ConnectToQuickBooks()
        {
            try
            {
                if (!_qb.Connect())
                {
                    return ShowRetryDialog("Unable to connect to QuickBooks.\nWould you like to retry?",
                        "Connection Failed") && ConnectToQuickBooks();
                }

                lblCompanyName.Text = $"Connected: {_qb.CurrentCompanyName}";
                _currentCompany = _companyManager.GetCompany(_qb.CurrentCompanyName);

                if (_currentCompany == null)
                {
                    ShowInfo("Company not configured. Please setup FBR token and seller information.", "Setup Required");
                    ShowCompanySetup();
                }
                else
                {
                    _qb.Connect(_currentCompany);
                    _qb.SetSroDataService(_sroDataService);
                    _invoiceManager.SetCompany(_currentCompany);
                    _ = LoadTransactionTypesAsync();
                    LoadInvoices();
                }

                return true;
            }
            catch (Exception ex)
            {
                return ShowRetryDialog($"QuickBooks connection failed:\n{ex.Message}\n\nWould you like to retry?",
                    "Connection Error") && ConnectToQuickBooks();
            }
        }

        private async Task LoadTransactionTypesAsync()
        {
            try
            {
                statusLabel.Text = "Loading transaction types from FBR...";
                bool success = await _transactionTypeService.FetchAndStoreTransactionTypesAsync(_currentCompany?.FBRToken);
                statusLabel.Text = success ? "Transaction types loaded" : "Ready";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading transaction types: {ex.Message}");
                statusLabel.Text = "Ready";
            }
        }

        #endregion

        #region Data Management

        private void LoadInvoices()
        {
            if (_currentCompany == null) return;

            try
            {
                dgvInvoices.SuspendLayout();

                // Save checkbox states before clearing
                SaveCheckboxStates();

                dgvInvoices.DataSource = null;

                _allInvoices = _db.GetInvoices(_currentCompany.CompanyName);
                ApplyFilters();
                UpdateStatistics();

                dgvInvoices.Refresh();
                statusLabel.Text = $"Loaded {_allInvoices?.Count ?? 0} invoices";
            }
            catch (Exception ex)
            {
                ShowError("Error loading invoices", ex.Message);
                statusLabel.Text = "Error loading invoices";
            }
            finally
            {
                dgvInvoices.ResumeLayout(true);
            }
        }

        private void ApplyFilters()
        {
            if (_allInvoices == null) return;

            // Save checkbox states before filtering
            SaveCheckboxStates();

            var filtered = _allInvoices.AsEnumerable();

            string statusFilter = cboStatusFilter.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(statusFilter) && statusFilter != "All")
                filtered = filtered.Where(i => i.Status == statusFilter);

            string searchText = txtSearchInvoice.Text.Trim();
            if (!string.IsNullOrEmpty(searchText))
            {
                filtered = filtered.Where(i =>
                    i.InvoiceNumber?.Contains(searchText, StringComparison.OrdinalIgnoreCase) == true ||
                    i.CustomerName?.Contains(searchText, StringComparison.OrdinalIgnoreCase) == true ||
                    i.CustomerNTN?.Contains(searchText, StringComparison.OrdinalIgnoreCase) == true
                );
            }

            var filteredList = filtered.ToList();

            dgvInvoices.DataSource = null;
            dgvInvoices.DataSource = filteredList;

            FormatDataGridView(filteredList);

            statusLabel.Text = $"Showing {filteredList.Count} of {_allInvoices.Count} invoices";
        }

        private void UpdateStatistics()
        {
            if (_allInvoices == null) return;

            var stats = _allInvoices.GroupBy(i => i.Status)
                .ToDictionary(g => g.Key, g => g.Count());

            statusStats.Text = $"Total: {_allInvoices.Count} | " +
                              $"Pending: {stats.GetValueOrDefault("Pending", 0)} | " +
                              $"Uploaded: {stats.GetValueOrDefault("Uploaded", 0)} | " +
                              $"Failed: {stats.GetValueOrDefault("Failed", 0)}";
        }

        #endregion

        #region DataGridView Formatting

        private void FormatDataGridView(List<Invoice> invoices)
        {
            if (invoices == null) return;

            // Only add checkbox column if it doesn't exist
            if (!dgvInvoices.Columns.Contains("Select"))
            {
                AddCheckboxColumn();
            }

            HideUnnecessaryColumns();
            ConfigureVisibleColumns();

            if (invoices.Count > 0)
            {
                ApplyRowStyling();
                RestoreCheckboxStates();
            }
        }

        private void AddCheckboxColumn()
        {
            var checkboxColumn = new DataGridViewCheckBoxColumn
            {
                Name = "Select",
                HeaderText = "Select",
                Width = 50,
                FillWeight = 5,
                ReadOnly = false,
                Frozen = false,
                DisplayIndex = 0
            };

            dgvInvoices.Columns.Insert(0, checkboxColumn);
        }

        private void HideUnnecessaryColumns()
        {
            var columnsToHide = new[] {
                "Id", "QuickBooksInvoiceId", "CompanyName", "FBR_QRCode", "CreatedDate", "ModifiedDate",
                "CustomerAddress", "CustomerPhone", "CustomerEmail", "TotalAmount", "TaxAmount",
                "DiscountAmount", "InvoiceDate", "PaymentMode", "Items", "CustomerType"
            };

            foreach (var col in columnsToHide)
            {
                if (dgvInvoices.Columns[col] != null)
                    dgvInvoices.Columns[col].Visible = false;
            }
        }


        private void ConfigureVisibleColumns()
        {
            var columnConfig = new Dictionary<string, (string Header, int Weight, string Format, DataGridViewContentAlignment? Alignment)>
            {
                ["InvoiceNumber"] = ("Invoice #", 12, null, null),
                ["CustomerName"] = ("Customer", 22, null, null),
                ["CustomerNTN"] = ("NTN/CNIC", 15, null, null),
                ["Amount"] = ("Amount", 10, "N2", DataGridViewContentAlignment.MiddleRight),
                ["Status"] = ("Status", 8, null, DataGridViewContentAlignment.MiddleCenter),
                ["FBR_IRN"] = ("FBR IRN", 18, null, null),
                ["UploadDate"] = ("Upload Date", 15, "dd-MMM-yyyy HH:mm", null),
                ["ErrorMessage"] = ("Error", 25, null, null)
            };

            foreach (var config in columnConfig)
            {
                var column = dgvInvoices.Columns[config.Key];
                if (column != null)
                {
                    column.HeaderText = config.Value.Header;
                    column.FillWeight = config.Value.Weight;
                    column.ReadOnly = true;

                    if (!string.IsNullOrEmpty(config.Value.Format))
                        column.DefaultCellStyle.Format = config.Value.Format;

                    if (config.Value.Alignment.HasValue)
                        column.DefaultCellStyle.Alignment = config.Value.Alignment.Value;
                }
            }
        }

        private void ApplyRowStyling()
        {
            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                row.DefaultCellStyle.Padding = new Padding(8, 4, 8, 4);
                row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(100, 181, 246);
                row.DefaultCellStyle.SelectionForeColor = Color.White;
            }
        }

        private void SaveCheckboxStates()
        {
            _checkboxStates.Clear();

            if (!dgvInvoices.Columns.Contains("Select")) return;

            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                if (row.DataBoundItem is Invoice invoice)
                {
                    var isChecked = row.Cells["Select"].Value as bool? ?? false;
                    _checkboxStates[invoice.Id] = isChecked;
                }
            }
        }

        private void RestoreCheckboxStates()
        {
            if (!dgvInvoices.Columns.Contains("Select")) return;

            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                if (row.DataBoundItem is Invoice invoice)
                {
                    if (_checkboxStates.ContainsKey(invoice.Id))
                    {
                        row.Cells["Select"].Value = _checkboxStates[invoice.Id];
                    }
                    else
                    {
                        row.Cells["Select"].Value = false;
                    }
                }
            }
        }

        #endregion

        #region Sorting

        private void SortDataGridView(string columnName)
        {
            var dataSource = dgvInvoices.DataSource as List<Invoice>;
            if (dataSource == null) return;

            // Save checkbox states before sorting
            SaveCheckboxStates();

            _sortAscending = _currentSortColumn == columnName ? !_sortAscending : true;
            _currentSortColumn = columnName;

            var sorted = SortInvoicesByColumn(dataSource, columnName, _sortAscending);

            dgvInvoices.DataSource = null;
            dgvInvoices.DataSource = sorted;

            FormatDataGridView(sorted);
            UpdateSortIndicator(columnName);
        }

        private List<Invoice> SortInvoicesByColumn(List<Invoice> invoices, string columnName, bool ascending)
        {
            return columnName switch
            {
                "InvoiceNumber" => ascending ? invoices.OrderBy(i => i.InvoiceNumber).ToList() : invoices.OrderByDescending(i => i.InvoiceNumber).ToList(),
                "CustomerName" => ascending ? invoices.OrderBy(i => i.CustomerName).ToList() : invoices.OrderByDescending(i => i.CustomerName).ToList(),
                "CustomerNTN" => ascending ? invoices.OrderBy(i => i.CustomerNTN).ToList() : invoices.OrderByDescending(i => i.CustomerNTN).ToList(),
                "Amount" => ascending ? invoices.OrderBy(i => i.Amount).ToList() : invoices.OrderByDescending(i => i.Amount).ToList(),
                "Status" => ascending ? invoices.OrderBy(i => i.Status).ToList() : invoices.OrderByDescending(i => i.Status).ToList(),
                "FBR_IRN" => ascending ? invoices.OrderBy(i => i.FBR_IRN).ToList() : invoices.OrderByDescending(i => i.FBR_IRN).ToList(),
                "UploadDate" => ascending ? invoices.OrderBy(i => i.UploadDate).ToList() : invoices.OrderByDescending(i => i.UploadDate).ToList(),
                "ErrorMessage" => ascending ? invoices.OrderBy(i => i.ErrorMessage).ToList() : invoices.OrderByDescending(i => i.ErrorMessage).ToList(),
                _ => invoices
            };
        }

        private void UpdateSortIndicator(string columnName)
        {
            foreach (DataGridViewColumn column in dgvInvoices.Columns)
            {
                if (column.Name != "Select" && column.HeaderText != null)
                    column.HeaderText = column.HeaderText.Replace(" ▲", "").Replace(" ▼", "");
            }

            var targetColumn = dgvInvoices.Columns[columnName];
            if (targetColumn != null)
                targetColumn.HeaderText += _sortAscending ? " ▲" : " ▼";
        }

        #endregion

        #region Event Handlers

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!ShowConfirmDialog("Are you sure you want to exit the application?\n\nQuickBooks connection will be closed.",
                "Confirm Exit"))
            {
                e.Cancel = true;
                return;
            }

            CleanupServices();
        }

        private void DgvInvoices_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            if (dgvInvoices.Columns[e.ColumnIndex].Name == "Select")
            {
                dgvInvoices.CommitEdit(DataGridViewDataErrorContexts.Commit);

                // Save the checkbox state immediately after change
                if (dgvInvoices.Rows[e.RowIndex].DataBoundItem is Invoice invoice)
                {
                    var isChecked = dgvInvoices.Rows[e.RowIndex].Cells["Select"].Value as bool? ?? false;
                    _checkboxStates[invoice.Id] = isChecked;
                }
            }
        }

        private void DgvInvoices_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvInvoices.IsCurrentCellDirty && dgvInvoices.CurrentCell.OwningColumn.Name == "Select")
            {
                dgvInvoices.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void DgvInvoices_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dgvInvoices.Rows[e.RowIndex].DataBoundItem is Invoice invoice)
                _ = ShowInvoiceDetailsAsync(invoice);
        }

        private void DgvInvoices_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgvInvoices.Columns[e.ColumnIndex].Name == "Select")
                ToggleSelectAll();
            else
                SortDataGridView(dgvInvoices.Columns[e.ColumnIndex].Name);
        }

        private void DgvInvoices_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (!dgvInvoices.Columns.Contains("Select")) return;

            // Don't override saved states, just set defaults for new rows
            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                if (row.Cells["Select"].Value == null)
                {
                    if (row.DataBoundItem is Invoice invoice && _checkboxStates.ContainsKey(invoice.Id))
                    {
                        row.Cells["Select"].Value = _checkboxStates[invoice.Id];
                    }
                    else
                    {
                        row.Cells["Select"].Value = false;
                    }
                }
            }
        }

        #endregion

        #region Button Actions

        private async void BtnFetchInvoices_Click(object sender, EventArgs e)
        {
            if (!ValidateCompanyConfiguration()) return;

            try
            {
                // Clear checkbox states for new fetch
                _checkboxStates.Clear();

                SetBusyState(true, "Fetching invoices from QuickBooks...");
                await Task.Run(() => _invoiceManager.FetchFromQuickBooks());

                if (InvokeRequired)
                {
                    Invoke(new Action(() =>
                    {
                        LoadInvoices();
                        ShowSuccess("Invoices fetched successfully!");
                    }));
                }
                else
                {
                    LoadInvoices();
                    ShowSuccess("Invoices fetched successfully!");
                }
            }
            catch (Exception ex)
            {
                ShowError("Error fetching invoices", ex.Message);
            }
            finally
            {
                SetBusyState(false, "Ready");
            }
        }

        private async void BtnUploadSelected_Click(object sender, EventArgs e)
        {
            if (!ValidateFBRConfiguration()) return;

            var checkedInvoices = GetCheckedInvoices();
            if (checkedInvoices.Count == 0)
            {
                ShowWarning("Please select invoices to upload using the checkboxes.", "No Selection");
                return;
            }

            await UploadInvoicesAsync(checkedInvoices);
        }

        private async void BtnUploadAll_Click(object sender, EventArgs e)
        {
            if (!ValidateFBRConfiguration()) return;

            var pendingInvoices = _allInvoices?.Where(i => i.Status == "Pending").ToList();
            if (pendingInvoices == null || pendingInvoices.Count == 0)
            {
                ShowInfo("No pending invoices to upload.", "Nothing to Upload");
                return;
            }

            if (!ShowConfirmDialog($"Upload {pendingInvoices.Count} pending invoices to FBR?", "Confirm Upload"))
                return;

            await UploadInvoicesAsync(pendingInvoices);
        }

        private void BtnGeneratePDF_Click(object sender, EventArgs e)
        {
            if (dgvInvoices.SelectedRows.Count == 0)
            {
                ShowWarning("Please select an invoice to generate PDF.", "No Selection");
                return;
            }

            var invoice = dgvInvoices.SelectedRows[0].DataBoundItem as Invoice;
            if (invoice == null) return;

            if (invoice.Status != "Uploaded")
            {
                ShowWarning("Only uploaded invoices can generate FBR compliant PDFs.", "Not Uploaded");
                return;
            }

            GeneratePDF(invoice);
        }

        #endregion

        #region Invoice Operations

        private async Task UploadInvoicesAsync(List<Invoice> invoices)
        {
            progressBar.Visible = true;
            progressBar.Maximum = invoices.Count;
            progressBar.Value = 0;
            progressBar.Style = ProgressBarStyle.Blocks;

            int uploaded = 0;
            int failed = 0;
            var failedInvoices = new List<(string InvoiceNumber, string Error)>();

            foreach (var invoice in invoices)
            {
                if (invoice.Status == "Uploaded")
                {
                    progressBar.Value++;
                    continue;
                }

                statusLabel.Text = $"Uploading {invoice.InvoiceNumber}...";

                try
                {
                    var response = await _invoiceManager.UploadToFBR(invoice, _currentCompany.FBRToken);

                    if (response.Success)
                    {
                        uploaded++;
                    }
                    else
                    {
                        failed++;
                        var errorMsg = response.ErrorMessage ?? "Unknown error - no error message returned";
                        failedInvoices.Add((invoice.InvoiceNumber, errorMsg));
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    failedInvoices.Add((invoice.InvoiceNumber, $"Exception: {ex.Message}"));
                }

                progressBar.Value++;
            }

            LoadInvoices();
            progressBar.Visible = false;
            statusLabel.Text = $"Upload complete: {uploaded} success, {failed} failed";

            ShowUploadResults(uploaded, failed, failedInvoices);
        }

        private async Task ShowInvoiceDetailsAsync(Invoice invoice)
        {
            try
            {
                if (!ValidateCompanyConfiguration() || !_qb.Connect(_currentCompany))
                {
                    ShowError("Connection Error", "Failed to connect to QuickBooks.");
                    return;
                }

                statusLabel.Text = $"Loading details for {invoice.InvoiceNumber}...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
                Application.DoEvents();

                var details = await _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);
                if (details == null)
                {
                    ShowError("Error", "Could not retrieve invoice details from QuickBooks.");
                    return;
                }

                details.SellerNTN = _currentCompany.SellerNTN;
                details.SellerBusinessName = _currentCompany.CompanyName;
                details.SellerProvince = _currentCompany.SellerProvince;
                details.SellerAddress = _currentCompany.SellerAddress;

                var fbrPayload = _fbr.BuildFBRPayload(details);
                var apiPayload = _fbr.ConvertToApiPayload(fbrPayload);
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(apiPayload, Newtonsoft.Json.Formatting.Indented);

                var validationErrors = ValidateInvoiceDetails(fbrPayload);
                ShowInvoiceDetailsDialog(invoice.InvoiceNumber, json, validationErrors);
            }
            catch (Exception ex)
            {
                ShowErrorDetailsDialog(ex);
            }
            finally
            {
                progressBar.Visible = false;
                statusLabel.Text = "Ready";
            }
        }

        private void GeneratePDF(Invoice invoice)
        {
            using var saveDialog = new SaveFileDialog
            {
                Filter = "PDF Files|*.pdf",
                FileName = $"Invoice_{invoice.InvoiceNumber}_{DateTime.Now:yyyyMMdd}.pdf"
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _invoiceManager.GeneratePDF(invoice, saveDialog.FileName);
                    ShowSuccess("PDF generated successfully!");

                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    ShowError("Error generating PDF", ex.Message);
                }
            }
        }

        #endregion

        #region Company Setup

        private void ShowCompanySetup()
        {
            _currentCompany = _companyManager.GetCompany(_qb.CurrentCompanyName) ?? new Company
            {
                CompanyName = _qb.CurrentCompanyName
            };

            try
            {
                var qbInfo = _qb.GetCompanyInfo();
                if (qbInfo != null)
                {
                    _currentCompany.SellerAddress ??= qbInfo.Address;
                    _currentCompany.SellerPhone ??= qbInfo.Phone;
                    _currentCompany.SellerEmail ??= qbInfo.Email;
                }
            }
            catch (Exception ex)
            {
                ShowWarning($"Could not fetch company info from QuickBooks: {ex.Message}\n\nYou can still configure the FBR token and NTN.",
                    "Warning");
            }

            using var setupForm = new CompanySetupForm(_qb.CurrentCompanyName, _currentCompany);
            if (setupForm.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _currentCompany = setupForm.Company;
                    _companyManager.SaveCompany(_currentCompany);
                    _invoiceManager.SetCompany(_currentCompany);
                    LoadInvoices();
                    _ = LoadTransactionTypesAsync();
                    ShowSuccess("Company settings saved successfully!");
                }
                catch (Exception ex)
                {
                    ShowError("Error saving company settings", ex.Message);
                }
            }
        }

        #endregion

        #region Validation

        private bool ValidateCompanyConfiguration()
        {
            if (_currentCompany == null)
            {
                ShowWarning("Please configure company settings first.", "Configuration Required");
                return false;
            }
            return true;
        }

        private bool ValidateFBRConfiguration()
        {
            if (_currentCompany == null || string.IsNullOrEmpty(_currentCompany.FBRToken))
            {
                ShowWarning("Please configure company FBR token first.", "Configuration Required");
                return false;
            }

            if (string.IsNullOrEmpty(_currentCompany.SellerNTN))
            {
                ShowWarning("Please configure Seller NTN in company settings.", "Configuration Required");
                return false;
            }

            if (string.IsNullOrEmpty(_currentCompany.SellerProvince))
            {
                ShowWarning("Please configure Seller Province in company settings.", "Configuration Required");
                return false;
            }

            return true;
        }

        private List<string> ValidateInvoiceDetails(FBRInvoicePayload details)
        {
            var errors = new List<string>();

            if (string.IsNullOrEmpty(details.SellerNTN))
                errors.Add("⚠️ Seller NTN is missing");
            if (string.IsNullOrEmpty(details.SellerProvince))
                errors.Add("⚠️ Seller Province is missing");
            if (string.IsNullOrEmpty(details.CustomerNTN))
                errors.Add("⚠️ Customer NTN is missing");
            if (string.IsNullOrEmpty(details.BuyerProvince))
                errors.Add("⚠️ Customer Province is missing");

            var validSaleTypes = new HashSet<string>(
                _transactionTypeService.GetTransactionTypes().Select(t => t.TransactionDesc),
                StringComparer.OrdinalIgnoreCase
            );

            foreach (var item in details.Items)
            {
                if (string.IsNullOrEmpty(item.HSCode))
                    errors.Add($"⚠️ HS Code missing for item: {item.ItemName}");
                if (item.RetailPrice == 0)
                    errors.Add($"⚠️ Retail Price missing for item: {item.ItemName}");
                if (string.IsNullOrEmpty(item.SaleType))
                    errors.Add($"⚠️ Sale Type missing for item: {item.ItemName}");
                else if (!validSaleTypes.Contains(item.SaleType))
                    errors.Add($"⚠️ Invalid Sale Type '{item.SaleType}' for item: {item.ItemName}");
            }

            return errors;
        }

        #endregion

        #region UI Dialogs

        private void ShowInvoiceDetailsDialog(string invoiceNumber, string json, List<string> validationErrors)
        {
            using var detailForm = new Form
            {
                Text = $"Invoice JSON - {invoiceNumber}",
                Size = new Size(1000, 750),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimizeBox = false,
                MaximizeBox = true
            };

            if (validationErrors.Count > 0)
            {
                var warningPanel = CreateWarningPanel(validationErrors);
                detailForm.Controls.Add(warningPanel);
            }

            var txtJson = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F),
                BackColor = Color.White,
                Text = json,
                WordWrap = false
            };

            var buttonPanel = CreateInvoiceDetailButtons(json, validationErrors, detailForm);

            detailForm.Controls.Add(txtJson);
            detailForm.Controls.Add(buttonPanel);
            detailForm.ShowDialog();
        }

        private Panel CreateWarningPanel(List<string> errors)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = Math.Min(150, errors.Count * 25 + 50),
                BackColor = Color.FromArgb(255, 243, 205),
                BorderStyle = BorderStyle.FixedSingle
            };

            panel.Controls.Add(new Label
            {
                Text = $"⚠️ VALIDATION WARNINGS ({errors.Count})",
                Location = new Point(10, 10),
                Size = new Size(950, 20),
                Font = new Font("Arial", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(102, 60, 0)
            });

            panel.Controls.Add(new TextBox
            {
                Text = string.Join(Environment.NewLine, errors),
                Location = new Point(10, 35),
                Size = new Size(960, panel.Height - 45),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(255, 243, 205),
                ForeColor = Color.FromArgb(102, 60, 0),
                BorderStyle = BorderStyle.None,
                Font = new Font("Arial", 9F)
            });

            return panel;
        }

        private Panel CreateInvoiceDetailButtons(string json, List<string> errors, Form parentForm)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                BackColor = Color.FromArgb(240, 240, 240)
            };

            var btnCopy = new Button
            {
                Text = "Copy JSON",
                Size = new Size(120, 35),
                Location = new Point(10, 8),
                BackColor = Color.FromArgb(33, 150, 243),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCopy.Click += (s, e) =>
            {
                Clipboard.SetText(json);
                ShowSuccess("JSON copied to clipboard!");
            };

            var btnValidate = new Button
            {
                Text = "Validate",
                Size = new Size(100, 35),
                Location = new Point(140, 8),
                BackColor = errors.Count == 0 ? Color.FromArgb(46, 125, 50) : Color.FromArgb(255, 152, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnValidate.Click += (s, e) =>
            {
                if (errors.Count == 0)
                    ShowSuccess("All required fields are present!");
                else
                    ShowWarning($"Found {errors.Count} validation warning(s):\n\n{string.Join("\n", errors)}",
                        "Validation Warnings");
            };

            var btnClose = new Button
            {
                Text = "Close",
                Size = new Size(100, 35),
                Location = new Point(250, 8),
                BackColor = Color.FromArgb(96, 125, 139),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnClose.Click += (s, e) => parentForm.Close();

            panel.Controls.AddRange(new Control[] { btnCopy, btnValidate, btnClose });
            return panel;
        }

        private void ShowUploadResults(int uploaded, int failed, List<(string InvoiceNumber, string Error)> failedInvoices)
        {
            if (failedInvoices.Count == 0)
            {
                ShowSuccess($"All {uploaded} invoice(s) uploaded successfully!");
                return;
            }

            using var resultForm = new Form
            {
                Text = "Upload Results",
                Width = 900,
                Height = 650,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimumSize = new Size(700, 500)
            };

            var lblSummary = new Label
            {
                Text = $"✅ Uploaded: {uploaded}    ❌ Failed: {failed}",
                Location = new Point(20, 20),
                Size = new Size(850, 30),
                Font = new Font("Arial", 12F, FontStyle.Bold),
                ForeColor = failed > 0 ? Color.DarkRed : Color.DarkGreen,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var txtFailed = new TextBox
            {
                Text = string.Join(Environment.NewLine + new string('=', 80) + Environment.NewLine,
                    failedInvoices.Select(f => $"Invoice: {f.InvoiceNumber}\nError:\n{f.Error}\n")),
                Location = new Point(20, 85),
                Size = new Size(840, 450),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9F),
                BackColor = Color.FromArgb(255, 245, 245),
                ForeColor = Color.DarkRed,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            var btnCopy = new Button
            {
                Text = "Copy Errors",
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(33, 150, 243),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnCopy.Location = new Point(20, resultForm.ClientSize.Height - 55);
            btnCopy.Click += (s, e) =>
            {
                Clipboard.SetText(txtFailed.Text);
                ShowSuccess("Error details copied to clipboard!");
            };

            var btnClose = new Button
            {
                Text = "Close",
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(96, 125, 139),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnClose.Location = new Point(resultForm.ClientSize.Width - 140, resultForm.ClientSize.Height - 55);
            btnClose.Click += (s, e) => resultForm.Close();

            resultForm.Controls.AddRange(new Control[] { lblSummary, txtFailed, btnCopy, btnClose });
            resultForm.ShowDialog();
        }

        private void ShowErrorDetailsDialog(Exception ex)
        {
            var errorDetails = new StringBuilder();
            errorDetails.AppendLine("❌ ERROR DETAILS:");
            errorDetails.AppendLine($"Message: {ex.Message}");
            errorDetails.AppendLine();
            errorDetails.AppendLine("Stack Trace:");
            errorDetails.AppendLine(ex.StackTrace);

            if (ex.InnerException != null)
            {
                errorDetails.AppendLine();
                errorDetails.AppendLine("Inner Exception:");
                errorDetails.AppendLine(ex.InnerException.Message);
            }

            using var errorForm = new Form
            {
                Text = "Error Details",
                Size = new Size(800, 600),
                StartPosition = FormStartPosition.CenterParent
            };

            var txtError = new TextBox
            {
                Text = errorDetails.ToString(),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = Color.FromArgb(255, 235, 238),
                ForeColor = Color.DarkRed
            };

            var btnClose = new Button
            {
                Text = "Close",
                Dock = DockStyle.Bottom,
                Height = 40,
                BackColor = Color.FromArgb(211, 47, 47),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnClose.Click += (s, e) => errorForm.Close();

            errorForm.Controls.AddRange(new Control[] { txtError, btnClose });
            errorForm.ShowDialog();
        }

        #endregion

        #region Helper Methods

        private List<Invoice> GetCheckedInvoices()
        {
            var checkedInvoices = new List<Invoice>();

            if (!dgvInvoices.Columns.Contains("Select")) return checkedInvoices;

            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                if (row.Cells["Select"].Value as bool? == true && row.DataBoundItem is Invoice invoice)
                    checkedInvoices.Add(invoice);
            }

            return checkedInvoices;
        }

        private void ToggleSelectAll()
        {
            bool allSelected = dgvInvoices.Rows.Cast<DataGridViewRow>()
                .All(row => row.Cells["Select"].Value as bool? == true);

            dgvInvoices.EndEdit();
            foreach (DataGridViewRow row in dgvInvoices.Rows)
                row.Cells["Select"].Value = !allSelected;
            dgvInvoices.RefreshEdit();
            dgvInvoices.Refresh();
        }

        private void SetBusyState(bool busy, string statusText)
        {
            statusLabel.Text = statusText;
            progressBar.Visible = busy;
            progressBar.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
            btnFetchInvoices.Enabled = !busy;
        }

        private void CleanupServices()
        {
            statusLabel.Text = "Closing QuickBooks connection...";
            Application.DoEvents();

            try
            {
                _qb?.Dispose();
                _fbr?.Dispose();
                statusLabel.Text = "Application closing...";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error closing services: {ex.Message}");
            }
        }

        #endregion

        #region Message Box Helpers

        private void ShowError(string title, string message) =>
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);

        private void ShowWarning(string message, string title) =>
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);

        private void ShowInfo(string message, string title) =>
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);

        private void ShowSuccess(string message) =>
            MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

        private bool ShowConfirmDialog(string message, string title) =>
            MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

        private bool ShowRetryDialog(string message, string title) =>
            MessageBox.Show(message, title, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning) == DialogResult.Retry;

        #endregion

        #region Cleanup

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            CleanupServices();
        }

        #endregion
    }

    public static class ControlExtensions
    {
        public static void DoubleBuffered(this Control control, bool enable)
        {
            var propertyInfo = control.GetType().GetProperty("DoubleBuffered",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            propertyInfo?.SetValue(control, enable, null);
        }
    }
}