using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using C2B_FBR_Connect.Utilities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C2B_FBR_Connect.Forms
{
    public partial class MainForm : Form
    {
        #region Fields

        private DataGridView dgvInvoices;
        private Button btnFetchInvoices, btnUploadSelected, btnGeneratePDF, btnCompanySetup, btnRefresh;
        private Label lblCompanyName, lblFilter, lblSearch;
        private ComboBox cboStatusFilter;
        private TextBox txtSearchInvoice;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel, statusStats;
        private ToolStripProgressBar progressBar;
        private Panel topPanel, buttonPanel;
        private PictureBox picCompanyLogo;  // ✅ Added logo PictureBox

        private readonly DatabaseService _db;
        private readonly QuickBooksService _qb;
        private readonly FBRApiService _fbr;
        private readonly PDFService _pdf;
        private readonly CompanyManager _companyManager;
        private readonly InvoiceManager _invoiceManager;
        private readonly TransactionTypeService _transactionTypeService;
        private readonly SroDataService _sroDataService;

        private Company _currentCompany;
        private List<Invoice> _allInvoices;
        private HashSet<int> _selectedInvoiceIds = new HashSet<int>();

        private System.Windows.Forms.Timer _searchDebounceTimer;
        private CancellationTokenSource _filterCancellationTokenSource;
        private const int DEBOUNCE_DELAY_MS = 300;

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

            InitializeSearchDebouncer();
            SetupCustomUI();
            ConnectToQuickBooks();

            this.FormClosing += MainForm_FormClosing;
        }

        private void InitializeSearchDebouncer()
        {
            _searchDebounceTimer = new System.Windows.Forms.Timer { Interval = DEBOUNCE_DELAY_MS };
            _searchDebounceTimer.Tick += (s, e) =>
            {
                _searchDebounceTimer.Stop();
                _ = ApplyFiltersAsync();
            };
        }

        private void SetupCustomUI()
        {
            SuspendLayout();

            ConfigureForm();
            CreateTopPanel();
            CreateButtonPanel();
            CreateDataGridView();
            CreateStatusStrip();

            Controls.AddRange(new Control[] { dgvInvoices, buttonPanel, topPanel, statusStrip });

            ResumeLayout(false);
            PerformLayout();
        }

        private void ConfigureForm()
        {
            Text = "C2B Smart App - Digital Invoicing System";
            Size = new Size(1400, 750);
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1000, 600);

            try
            {
                Icon = new Icon(System.IO.Path.Combine(Application.StartupPath, "assets", "favicon.ico"));
            }
            catch { }
        }

        private void CreateTopPanel()
        {
            topPanel = new Panel
            {
                Height = 70,  // ✅ Increased height for logo
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(12, 10, 12, 10)
            };

            // ✅ Company Logo PictureBox
            picCompanyLogo = new PictureBox
            {
                Location = new Point(12, 10),
                Size = new Size(150, 50),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.None,
                BackColor = Color.Transparent
            };

            // ✅ Company Name Label (adjusted position)
            lblCompanyName = new Label
            {
                Text = "No QuickBooks Company Connected",
                Location = new Point(175, 10),  // Positioned after logo
                AutoSize = false,
                Height = 50,
                Width = topPanel.Width - 200,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                TextAlign = ContentAlignment.MiddleLeft,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            topPanel.Controls.Add(picCompanyLogo);
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
            cboStatusFilter.SelectedIndexChanged += (s, e) => _ = ApplyFiltersAsync();

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
            txtSearchInvoice.TextChanged += OnSearchTextChanged;

            buttonPanel.Controls.AddRange(new Control[] { lblFilter, cboStatusFilter, lblSearch, txtSearchInvoice });
        }

        private void CreateActionButtons()
        {
            int btnY = 40;
            int btnSpacing = 8;
            int currentX = 0;

            btnFetchInvoices = CreateButton("Fetch Invoices", ref currentX, btnY, 120, Color.FromArgb(0, 122, 204), btnSpacing);
            btnFetchInvoices.Click += async (s, e) => await BtnFetchInvoices_ClickAsync();

            btnRefresh = CreateButton("Refresh", ref currentX, btnY, 90, Color.FromArgb(46, 125, 50), btnSpacing);
            btnRefresh.Click += (s, e) => _ = LoadInvoicesAsync();

            btnUploadSelected = CreateButton("Upload Selected", ref currentX, btnY, 130, Color.FromArgb(255, 152, 0), btnSpacing);
            btnUploadSelected.Click += async (s, e) => await BtnUploadSelected_ClickAsync();

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

                lblCompanyName.Text = $"{_qb.CurrentCompanyName}";
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

                    // ✅ Load company logo
                    LoadCompanyLogo();

                    _ = LoadTransactionTypesAsync();
                    _ = LoadInvoicesAsync();
                }

                return true;
            }
            catch (Exception ex)
            {
                return ShowRetryDialog($"QuickBooks connection failed:\n{ex.Message}\n\nWould you like to retry?",
                    "Connection Error") && ConnectToQuickBooks();
            }
        }

        // ✅ New method to load company logo
        private void LoadCompanyLogo()
        {
            try
            {
                // Dispose previous image if exists
                if (picCompanyLogo.Image != null)
                {
                    picCompanyLogo.Image.Dispose();
                    picCompanyLogo.Image = null;
                }

                if (_currentCompany?.LogoImage != null && _currentCompany.LogoImage.Length > 0)
                {
                    picCompanyLogo.Image = ImageHelper.ByteArrayToImage(_currentCompany.LogoImage);
                    picCompanyLogo.Visible = true;

                    // Adjust company name label position
                    lblCompanyName.Location = new Point(175, 10);
                    lblCompanyName.Width = topPanel.Width - 200;
                }
                else
                {
                    // No logo - hide picture box and expand label
                    picCompanyLogo.Visible = false;
                    lblCompanyName.Location = new Point(12, 10);
                    lblCompanyName.Width = topPanel.Width - 24;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading company logo: {ex.Message}");
                picCompanyLogo.Visible = false;
                lblCompanyName.Location = new Point(12, 10);
                lblCompanyName.Width = topPanel.Width - 24;
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

        private async Task LoadInvoicesAsync()
        {
            if (_currentCompany == null) return;

            try
            {
                statusLabel.Text = "Loading invoices...";

                // Load data asynchronously
                _allInvoices = await Task.Run(() => _db.GetInvoices(_currentCompany.CompanyName));

                await ApplyFiltersAsync();
                UpdateStatistics();

                statusLabel.Text = $"Loaded {_allInvoices?.Count ?? 0} invoices";
            }
            catch (Exception ex)
            {
                ShowError("Error loading invoices", ex.Message);
                statusLabel.Text = "Error loading invoices";
            }
        }

        private void OnSearchTextChanged(object sender, EventArgs e)
        {
            _searchDebounceTimer.Stop();
            _searchDebounceTimer.Start();
        }

        private async Task ApplyFiltersAsync()
        {
            if (_allInvoices == null) return;

            // Cancel any ongoing filter operation
            _filterCancellationTokenSource?.Cancel();
            _filterCancellationTokenSource = new CancellationTokenSource();
            var token = _filterCancellationTokenSource.Token;

            try
            {
                var statusFilter = cboStatusFilter.SelectedItem?.ToString();
                var searchText = txtSearchInvoice.Text.Trim();

                // Perform filtering on background thread
                var filtered = await Task.Run(() =>
                {
                    var result = _allInvoices.AsEnumerable();

                    if (!string.IsNullOrEmpty(statusFilter) && statusFilter != "All")
                        result = result.Where(i => i.Status == statusFilter);

                    if (!string.IsNullOrEmpty(searchText))
                    {
                        result = result.Where(i =>
                            (i.InvoiceNumber?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (i.CustomerName?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (i.CustomerNTN?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                        );
                    }

                    return result.ToList();
                }, token);

                if (token.IsCancellationRequested) return;

                UpdateDataGridView(filtered);

                statusLabel.Text = $"Showing {filtered.Count} of {_allInvoices.Count} invoices";
            }
            catch (OperationCanceledException)
            {
                // Filter was cancelled, ignore
            }
        }

        private void UpdateDataGridView(List<Invoice> invoices)
        {
            dgvInvoices.SuspendLayout();

            try
            {
                dgvInvoices.DataSource = null;
                dgvInvoices.DataSource = invoices;

                FormatDataGridView(invoices);
            }
            finally
            {
                dgvInvoices.ResumeLayout();
            }
        }

        private void UpdateStatistics()
        {
            if (_allInvoices == null) return;

            var stats = new Dictionary<string, int>();
            foreach (var invoice in _allInvoices)
            {
                var status = invoice.Status ?? "Unknown";
                if (stats.ContainsKey(status))
                    stats[status]++;
                else
                    stats[status] = 1;
            }

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

            // Add checkbox column only once
            if (!dgvInvoices.Columns.Contains("Select"))
            {
                AddCheckboxColumn();
            }

            HideUnnecessaryColumns();
            ConfigureVisibleColumns();

            if (invoices.Count > 0)
            {
                RestoreCheckboxStates();
            }
        }

        private void AddCheckboxColumn()
        {
            var checkboxColumn = new DataGridViewCheckBoxColumn
            {
                Name = "Select",
                HeaderText = "☐",
                Width = 40,
                FillWeight = 5,
                ReadOnly = false,
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
                ["InvoiceType"] = ("Type", 10, null, DataGridViewContentAlignment.MiddleCenter), // Add this
                ["InvoiceNumber"] = ("Invoice #", 12, null, null),
                ["CustomerName"] = ("Customer", 22, null, null),
                ["CustomerNTN"] = ("NTN/CNIC", 15, null, null),
                ["Amount"] = ("Amount", 10, "N2", DataGridViewContentAlignment.MiddleRight),
                ["Status"] = ("Status", 8, null, DataGridViewContentAlignment.MiddleCenter),
                ["FBR_IRN"] = ("FBR IRN", 18, null, null),
                ["UploadDate"] = ("Upload Date", 15, "dd-MMM-yyyy HH:mm", null),
                ["ErrorMessage"] = ("Error", 20, null, null) // Reduced from 25 to make room
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

        private void RestoreCheckboxStates()
        {
            if (!dgvInvoices.Columns.Contains("Select")) return;

            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                if (row.DataBoundItem is Invoice invoice)
                {
                    row.Cells["Select"].Value = _selectedInvoiceIds.Contains(invoice.Id);
                }
            }
        }

        #endregion

        #region Sorting

        private void SortDataGridView(string columnName)
        {
            var dataSource = dgvInvoices.DataSource as List<Invoice>;
            if (dataSource == null || dataSource.Count == 0) return;

            var sorted = columnName switch
            {
                "InvoiceNumber" => dataSource.OrderBy(i => i.InvoiceNumber).ToList(),
                "CustomerName" => dataSource.OrderBy(i => i.CustomerName).ToList(),
                "CustomerNTN" => dataSource.OrderBy(i => i.CustomerNTN).ToList(),
                "Amount" => dataSource.OrderByDescending(i => i.Amount).ToList(),
                "Status" => dataSource.OrderBy(i => i.Status).ToList(),
                "FBR_IRN" => dataSource.OrderBy(i => i.FBR_IRN).ToList(),
                "UploadDate" => dataSource.OrderByDescending(i => i.UploadDate).ToList(),
                "ErrorMessage" => dataSource.OrderBy(i => i.ErrorMessage).ToList(),
                _ => dataSource
            };

            UpdateDataGridView(sorted);
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

                if (dgvInvoices.Rows[e.RowIndex].DataBoundItem is Invoice invoice)
                {
                    var isChecked = dgvInvoices.Rows[e.RowIndex].Cells["Select"].Value as bool? ?? false;
                    if (isChecked)
                        _selectedInvoiceIds.Add(invoice.Id);
                    else
                        _selectedInvoiceIds.Remove(invoice.Id);

                    UpdateSelectedCount();
                }
            }
        }

        private void DgvInvoices_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvInvoices.IsCurrentCellDirty && dgvInvoices.CurrentCell?.OwningColumn.Name == "Select")
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

        #endregion

        #region Button Actions

        private async Task BtnFetchInvoices_ClickAsync()
        {
            if (!ValidateCompanyConfiguration()) return;

            using (var dateRangeDialog = new DateRangeDialog())
            {
                if (dateRangeDialog.ShowDialog() != DialogResult.OK)
                    return;

                try
                {
                    _selectedInvoiceIds.Clear();

                    var dateFrom = dateRangeDialog.DateFrom;
                    var dateTo = dateRangeDialog.DateTo;
                    var includeInvoices = dateRangeDialog.IncludeInvoices;
                    var includeCreditMemos = dateRangeDialog.IncludeCreditMemos;

                    string transactionTypes = "";
                    if (includeInvoices && includeCreditMemos)
                        transactionTypes = "invoices and credit memos";
                    else if (includeInvoices)
                        transactionTypes = "invoices";
                    else
                        transactionTypes = "credit memos";

                    SetBusyState(true, $"Fetching {transactionTypes} from {dateFrom:dd-MMM-yyyy} to {dateTo:dd-MMM-yyyy}...");

                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                    // ✅ Use updated InvoiceManager method with date range
                    var transactions = await _invoiceManager.FetchFromQuickBooks(
                        dateFrom,
                        dateTo,
                        includeInvoices,
                        includeCreditMemos
                    );

                    stopwatch.Stop();

                    await LoadInvoicesAsync();

                    ShowSuccess($"Fetch completed!\n\n" +
                               $"Date Range: {dateFrom:dd-MMM-yyyy} to {dateTo:dd-MMM-yyyy}\n" +
                               $"Fetched: {transactions.Count} {transactionTypes}\n" +
                               $"Time: {stopwatch.Elapsed.TotalSeconds:F1} seconds");
                }
                catch (Exception ex)
                {
                    ShowError("Error fetching transactions", ex.Message);
                }
                finally
                {
                    SetBusyState(false, "Ready");
                }
            }
        }

        private async Task BtnUploadSelected_ClickAsync()
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

        private void BtnGeneratePDF_Click(object sender, EventArgs e)
        {
            // ✅ First check for checked invoices, then fall back to row selection
            var checkedInvoices = GetCheckedInvoices().Where(i => i.Status == "Uploaded").ToList();

            if (checkedInvoices.Count == 0)
            {
                // Fall back to selected row if no checkboxes are checked
                if (dgvInvoices.SelectedRows.Count > 0)
                {
                    var selectedInvoice = dgvInvoices.SelectedRows[0].DataBoundItem as Invoice;
                    if (selectedInvoice != null && selectedInvoice.Status == "Uploaded")
                    {
                        checkedInvoices.Add(selectedInvoice);
                    }
                }
            }

            if (checkedInvoices.Count == 0)
            {
                // Check if there are checked invoices that aren't uploaded
                var pendingChecked = GetCheckedInvoices().Where(i => i.Status != "Uploaded").ToList();
                if (pendingChecked.Count > 0)
                {
                    ShowWarning($"Selected invoice(s) are not uploaded yet.\nOnly uploaded invoices can generate FBR compliant PDFs.", "Not Uploaded");
                }
                else
                {
                    ShowWarning("Please select uploaded invoice(s) using checkboxes or row selection to generate PDF.", "No Selection");
                }
                return;
            }

            if (checkedInvoices.Count == 1)
            {
                // Single PDF - use save dialog
                GeneratePDF(checkedInvoices[0]);
            }
            else
            {
                // Multiple PDFs - use folder browser
                GenerateMultiplePDFs(checkedInvoices);
            }
        }

        #endregion

        #region Invoice Operations

        private async Task UploadInvoicesAsync(List<Invoice> invoices)
        {
            if (invoices == null || invoices.Count == 0)
            {
                ShowWarning("No invoices to upload.", "Nothing to Upload");
                return;
            }

            try
            {
                // Filter out already uploaded invoices
                var invoicesToUpload = invoices.Where(i => i.Status != "Uploaded").ToList();
                var alreadyUploadedCount = invoices.Count - invoicesToUpload.Count;

                if (invoicesToUpload.Count == 0)
                {
                    ShowInfo("All selected invoices are already uploaded.", "Nothing to Upload");
                    // ✅ Clear selections for already uploaded invoices
                    ClearSelectionsForInvoices(invoices);
                    return;
                }

                // ✅ Show confirmation if some are already uploaded
                if (alreadyUploadedCount > 0)
                {
                    var msg = $"{alreadyUploadedCount} invoice(s) already uploaded and will be skipped.\n" +
                              $"Proceed to upload {invoicesToUpload.Count} pending invoice(s)?";
                    if (!ShowConfirmDialog(msg, "Confirm Upload"))
                        return;
                }

                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Blocks;
                progressBar.Minimum = 0;
                progressBar.Maximum = invoicesToUpload.Count;
                progressBar.Value = 0;

                int uploaded = 0;
                int failed = 0;
                var failedInvoices = new List<(string InvoiceNumber, string Error)>();
                var successfullyUploadedIds = new List<int>();

                foreach (var invoice in invoicesToUpload)
                {
                    statusLabel.Text = $"Uploading {invoice.InvoiceNumber}... ({progressBar.Value + 1}/{invoicesToUpload.Count})";

                    try
                    {
                        var response = await _invoiceManager.UploadToFBR(invoice, _currentCompany.FBRToken);

                        if (response.Success)
                        {
                            uploaded++;
                            successfullyUploadedIds.Add(invoice.Id);
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

                // ✅ Clear checkboxes for successfully uploaded invoices
                foreach (var id in successfullyUploadedIds)
                {
                    _selectedInvoiceIds.Remove(id);
                }

                // ✅ Also clear checkboxes for already uploaded invoices that were selected
                foreach (var invoice in invoices.Where(i => i.Status == "Uploaded"))
                {
                    _selectedInvoiceIds.Remove(invoice.Id);
                }

                await LoadInvoicesAsync();
                progressBar.Visible = false;

                UpdateSelectedCount();
                statusLabel.Text = $"Upload complete: {uploaded} success, {failed} failed";

                ShowUploadResults(uploaded, failed, failedInvoices);
            }
            catch (Exception ex)
            {
                progressBar.Visible = false;
                ShowError("Upload Error", $"An error occurred during upload: {ex.Message}");
            }
        }

        // ✅ Helper method to clear selections
        private void ClearSelectionsForInvoices(List<Invoice> invoices)
        {
            foreach (var invoice in invoices)
            {
                _selectedInvoiceIds.Remove(invoice.Id);
            }
            RefreshCheckboxStates();
            UpdateSelectedCount();
        }

        // ✅ Refresh checkbox visual states
        private void RefreshCheckboxStates()
        {
            if (!dgvInvoices.Columns.Contains("Select")) return;

            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                if (row.DataBoundItem is Invoice invoice)
                {
                    row.Cells["Select"].Value = _selectedInvoiceIds.Contains(invoice.Id);
                }
            }
            dgvInvoices.Refresh();
        }

        // ✅ Update button text with selected count
        private void UpdateSelectedCount()
        {
            var count = _selectedInvoiceIds.Count;
            btnUploadSelected.Text = count > 0 ? $"Upload Selected ({count})" : "Upload Selected";
            btnGeneratePDF.Text = count > 0 ? $"Generate PDF ({count})" : "Generate PDF";
        }

        private async Task ShowInvoiceDetailsAsync(Invoice invoice)
        {
            try
            {
                if (!ValidateCompanyConfiguration())
                    return;

                string transactionType = invoice.InvoiceType ?? "Invoice";
                statusLabel.Text = $"Loading details for {transactionType} {invoice.InvoiceNumber}...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;

                // ✅ Use InvoiceManager's method instead of calling QB directly
                var details = await _invoiceManager.GetFullInvoicePayload(invoice);

                if (details == null)
                {
                    ShowError("Error", $"Could not retrieve {transactionType.ToLower()} details from QuickBooks.");
                    return;
                }

                var fbrPayload = _fbr.BuildFBRPayload(details);
                var apiPayload = _fbr.ConvertToApiPayload(fbrPayload);
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(apiPayload, Newtonsoft.Json.Formatting.Indented);

                var validationErrors = ValidateInvoiceDetails(details); // Pass details, not fbrPayload
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

        // ✅ New method for generating multiple PDFs
        private void GenerateMultiplePDFs(List<Invoice> invoices)
        {
            using var folderDialog = new FolderBrowserDialog
            {
                Description = $"Select folder to save {invoices.Count} PDF files",
                ShowNewFolderButton = true
            };

            if (folderDialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                SetBusyState(true, "Generating PDFs...");
                progressBar.Style = ProgressBarStyle.Blocks;
                progressBar.Minimum = 0;
                progressBar.Maximum = invoices.Count;
                progressBar.Value = 0;

                int success = 0;
                int failed = 0;
                var errors = new List<string>();

                foreach (var invoice in invoices)
                {
                    try
                    {
                        statusLabel.Text = $"Generating PDF for {invoice.InvoiceNumber}... ({progressBar.Value + 1}/{invoices.Count})";

                        var fileName = System.IO.Path.Combine(
                            folderDialog.SelectedPath,
                            $"Invoice_{invoice.InvoiceNumber}_{DateTime.Now:yyyyMMdd}.pdf"
                        );

                        _invoiceManager.GeneratePDF(invoice, fileName);
                        success++;
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        errors.Add($"{invoice.InvoiceNumber}: {ex.Message}");
                    }

                    progressBar.Value++;
                }

                SetBusyState(false, "Ready");

                if (failed == 0)
                {
                    var result = MessageBox.Show(
                        $"Successfully generated {success} PDF(s)!\n\nWould you like to open the folder?",
                        "Success",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information
                    );

                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = folderDialog.SelectedPath,
                            UseShellExecute = true
                        });
                    }
                }
                else
                {
                    ShowWarning(
                        $"Generated: {success}\nFailed: {failed}\n\nErrors:\n{string.Join("\n", errors.Take(5))}" +
                        (errors.Count > 5 ? $"\n...and {errors.Count - 5} more" : ""),
                        "PDF Generation Complete"
                    );
                }
            }
            catch (Exception ex)
            {
                SetBusyState(false, "Ready");
                ShowError("Error generating PDFs", ex.Message);
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
                ShowWarning($"Could not fetch company info from QuickBooks: {ex.Message}", "Warning");
            }

            using var setupForm = new CompanySetupForm(_qb.CurrentCompanyName, _currentCompany);
            if (setupForm.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _currentCompany = setupForm.Company;
                    _companyManager.SaveCompany(_currentCompany);
                    _invoiceManager.SetCompany(_currentCompany);

                    // ✅ Reload company logo after setup
                    LoadCompanyLogo();

                    _ = LoadInvoicesAsync();
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

            if (string.IsNullOrEmpty(_currentCompany.SellerNTN) ||
                string.IsNullOrEmpty(_currentCompany.SellerProvince))
            {
                ShowWarning("Please configure Seller NTN and Province in company settings.", "Configuration Required");
                return false;
            }

            return true;
        }

        private List<string> ValidateInvoiceDetails(FBRInvoicePayload details)
        {
            var errors = new List<string>();

            if (string.IsNullOrEmpty(details.SellerNTN)) errors.Add("⚠️ Seller NTN is missing");
            if (string.IsNullOrEmpty(details.SellerProvince)) errors.Add("⚠️ Seller Province is missing");
            if (string.IsNullOrEmpty(details.CustomerNTN)) errors.Add("⚠️ Customer NTN is missing");
            if (string.IsNullOrEmpty(details.BuyerProvince)) errors.Add("⚠️ Customer Province is missing");

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
                else if (item.SaleType != "Goods at standard rate (default)" &&
                         !validSaleTypes.Contains(item.SaleType))
                {
                    errors.Add($"⚠️ Invalid Sale Type '{item.SaleType}' for item: {item.ItemName}");
                }
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
            var errorDetails = $"❌ ERROR DETAILS:\nMessage: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}";

            if (ex.InnerException != null)
            {
                errorDetails += $"\n\nInner Exception:\n{ex.InnerException.Message}";
            }

            using var errorForm = new Form
            {
                Text = "Error Details",
                Size = new Size(800, 600),
                StartPosition = FormStartPosition.CenterParent
            };

            var txtError = new TextBox
            {
                Text = errorDetails,
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
            var dataSource = dgvInvoices.DataSource as List<Invoice>;
            if (dataSource == null) return new List<Invoice>();

            return dataSource.Where(i => _selectedInvoiceIds.Contains(i.Id)).ToList();
        }

        private void ToggleSelectAll()
        {
            bool allSelected = dgvInvoices.Rows.Cast<DataGridViewRow>()
                .All(row => row.Cells["Select"].Value as bool? == true);

            _selectedInvoiceIds.Clear();

            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                row.Cells["Select"].Value = !allSelected;

                if (!allSelected && row.DataBoundItem is Invoice invoice)
                    _selectedInvoiceIds.Add(invoice.Id);
            }

            dgvInvoices.Refresh();
            UpdateSelectedCount();
        }

        private void SetBusyState(bool busy, string statusText)
        {
            statusLabel.Text = statusText;
            progressBar.Visible = busy;
            progressBar.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;

            btnFetchInvoices.Enabled = !busy;
            btnUploadSelected.Enabled = !busy;
            btnGeneratePDF.Enabled = !busy;
            btnRefresh.Enabled = !busy;
        }

        private void CleanupServices()
        {
            _searchDebounceTimer?.Dispose();
            _filterCancellationTokenSource?.Cancel();
            _filterCancellationTokenSource?.Dispose();

            // ✅ Dispose logo image
            if (picCompanyLogo?.Image != null)
            {
                picCompanyLogo.Image.Dispose();
            }

            try
            {
                _qb?.Dispose();
                _fbr?.Dispose();
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