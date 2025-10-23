using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C2B_FBR_Connect.Forms
{
    public partial class MainForm : Form
    {
        private DataGridView dgvInvoices = null!;
        private Button btnFetchInvoices = null!;
        private Button btnUploadSelected = null!;
        private Button btnUploadAll = null!;
        private Button btnGeneratePDF = null!;
        private Button btnCompanySetup = null!;
        private Button btnRefresh = null!;
        private Label lblCompanyName = null!;
        private Label lblFilter = null!;
        private ComboBox cboStatusFilter = null!;
        private TextBox txtSearchInvoice = null!;
        private Label lblSearch = null!;
        private StatusStrip statusStrip = null!;
        private ToolStripStatusLabel statusLabel = null!;
        private ToolStripProgressBar progressBar = null!;
        private ToolStripStatusLabel statusStats = null!;
        private Panel topPanel = null!;
        private Panel buttonPanel = null!;

        private DatabaseService _db = null!;
        private QuickBooksService _qb = null!;
        private FBRApiService _fbr = null!;
        private PDFService _pdf = null!;
        private CompanyManager _companyManager = null!;
        private InvoiceManager _invoiceManager = null!;
        private Company? _currentCompany;
        private List<Invoice>? _allInvoices;

        public MainForm()
        {
            InitializeComponent();
            SetupCustomUI();
            InitializeServices();

            this.FormClosing += MainForm_FormClosing;
            this.Resize += MainForm_Resize;
        }

        private void MainForm_Resize(object? sender, EventArgs e)
        {
            dgvInvoices?.Refresh();
        }

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            var result = MessageBox.Show(
                "Are you sure you want to exit the application?\n\nQuickBooks connection will be closed.",
                "Confirm Exit",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }

            statusLabel.Text = "Closing QuickBooks connection...";
            Application.DoEvents();

            try
            {
                if (_qb != null)
                {
                    _qb.Dispose();
                    System.Diagnostics.Debug.WriteLine("✅ QuickBooks connection closed successfully during application exit");
                }

                _fbr?.Dispose();

                statusLabel.Text = "Application closing...";
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Error closing QuickBooks connection: {ex.Message}");

                MessageBox.Show(
                    $"Warning: Error while closing QuickBooks connection:\n{ex.Message}\n\nThe application will still close.",
                    "Connection Close Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void SetupCustomUI()
        {
            this.Text = "C2B Smart App - Digital Invoicing System";
            this.Size = new Size(1400, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1000, 600);

            this.Icon = new System.Drawing.Icon(System.IO.Path.Combine(
    Application.StartupPath, "assets", "favicon.ico"));


            // Top Panel for company info
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

            // Button Panel
            buttonPanel = new Panel
            {
                Height = 95,
                Dock = DockStyle.Top,
                BackColor = Color.White,
                Padding = new Padding(12, 10, 12, 10)
            };

            // Filter Section
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
            cboStatusFilter.SelectedIndexChanged += CboStatusFilter_SelectedIndexChanged;

            // Search Section
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
            txtSearchInvoice.TextChanged += TxtSearchInvoice_TextChanged;

            // Action Buttons
            int btnY = 40;
            int btnSpacing = 8;
            int currentX = 0;

            btnFetchInvoices = CreateButton("Fetch Invoices", currentX, btnY, 120, Color.FromArgb(0, 122, 204));
            btnFetchInvoices.Click += BtnFetchInvoices_Click;
            currentX += 120 + btnSpacing;

            btnRefresh = CreateButton("Refresh", currentX, btnY, 90, Color.FromArgb(46, 125, 50));
            btnRefresh.Click += (s, e) => LoadInvoices();
            currentX += 90 + btnSpacing;

            btnUploadSelected = CreateButton("Upload Selected", currentX, btnY, 130, Color.FromArgb(255, 152, 0));
            btnUploadSelected.Click += BtnUploadSelected_Click;
            currentX += 130 + btnSpacing;

            btnUploadAll = CreateButton("Upload All", currentX, btnY, 100, Color.FromArgb(255, 87, 34));
            btnUploadAll.Click += BtnUploadAll_Click;
            currentX += 100 + btnSpacing;

            btnGeneratePDF = CreateButton("Generate PDF", currentX, btnY, 120, Color.FromArgb(156, 39, 176));
            btnGeneratePDF.Click += BtnGeneratePDF_Click;
            currentX += 120 + btnSpacing;

            btnCompanySetup = CreateButton("Company Setup", currentX, btnY, 130, Color.FromArgb(96, 125, 139));
            btnCompanySetup.Click += BtnCompanySetup_Click;

            buttonPanel.Controls.AddRange(new Control[] {
                lblFilter, cboStatusFilter, lblSearch, txtSearchInvoice,
                btnFetchInvoices, btnRefresh, btnUploadSelected, btnUploadAll, btnGeneratePDF, btnCompanySetup
            });

            // DataGridView
            dgvInvoices = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
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
                GridColor = Color.FromArgb(230, 230, 230)
            };

            dgvInvoices.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(52, 73, 94),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Padding = new Padding(8, 6, 8, 6),
                Alignment = DataGridViewContentAlignment.MiddleLeft
            };
            dgvInvoices.EnableHeadersVisualStyles = false;
            dgvInvoices.CellDoubleClick += DgvInvoices_CellDoubleClick;
            dgvInvoices.DoubleBuffered(true);

            // Status Strip
            statusStrip = new StatusStrip
            {
                BackColor = Color.FromArgb(240, 240, 240)
            };
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

            statusStrip.Items.Add(statusLabel);
            statusStrip.Items.Add(progressBar);
            statusStrip.Items.Add(statusStats);

            // Add controls in correct order
            this.Controls.Add(dgvInvoices);
            this.Controls.Add(buttonPanel);
            this.Controls.Add(topPanel);
            this.Controls.Add(statusStrip);
        }

        private Button CreateButton(string text, int x, int y, int width, Color backColor)
        {
            return new Button
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
        }

        private void InitializeServices()
        {
            _db = new DatabaseService();
            _qb = new QuickBooksService();
            _fbr = new FBRApiService();
            _pdf = new PDFService();
            _companyManager = new CompanyManager(_db);
            _invoiceManager = new InvoiceManager(_db, _qb, _fbr, _pdf);

            ConnectToQuickBooks();
        }

        private bool ConnectToQuickBooks()
        {
            try
            {
                _currentCompany = null;

                if (!_qb.Connect())
                {
                    var retryResult = MessageBox.Show(
                        "Unable to connect to QuickBooks.\nWould you like to retry?",
                        "Connection Failed",
                        MessageBoxButtons.RetryCancel,
                        MessageBoxIcon.Warning);

                    return retryResult == DialogResult.Retry ? ConnectToQuickBooks() : false;
                }

                lblCompanyName.Text = $"Connected: {_qb.CurrentCompanyName}";
                _currentCompany = _companyManager.GetCompany(_qb.CurrentCompanyName);

                if (_currentCompany == null)
                {
                    MessageBox.Show(
                        "Company not configured. Please setup FBR token and seller information.",
                        "Setup Required",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    ShowCompanySetup();
                }
                else
                {
                    _qb.Connect(_currentCompany);

                    _invoiceManager.SetCompany(_currentCompany);
                    LoadInvoices();
                }

                return true;
            }
            catch (Exception ex)
            {
                var retryResult = MessageBox.Show(
                    $"QuickBooks connection failed:\n{ex.Message}\n\nWould you like to retry?",
                    "Connection Error",
                    MessageBoxButtons.RetryCancel,
                    MessageBoxIcon.Error);

                return retryResult == DialogResult.Retry ? ConnectToQuickBooks() : false;
            }
        }

        private void LoadInvoices()
        {
            if (_currentCompany == null) return;

            try
            {
                dgvInvoices.SuspendLayout();
                _allInvoices = _db.GetInvoices(_currentCompany.CompanyName);
                ApplyFilters();
                UpdateStatistics();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading invoices: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Error loading invoices";
            }
            finally
            {
                dgvInvoices.ResumeLayout();
            }
        }

        private void ApplyFilters()
        {
            if (_allInvoices == null) return;

            var filtered = _allInvoices.AsEnumerable();

            string? statusFilter = cboStatusFilter.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(statusFilter) && statusFilter != "All")
            {
                filtered = filtered.Where(i => i.Status == statusFilter);
            }

            string searchText = txtSearchInvoice.Text.Trim();
            if (!string.IsNullOrEmpty(searchText))
            {
                filtered = filtered.Where(i =>
                    (i.InvoiceNumber?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (i.CustomerName?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (i.CustomerNTN?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false)
                );
            }

            var filteredList = filtered.ToList();
            dgvInvoices.DataSource = null;
            dgvInvoices.DataSource = filteredList;

            FormatDataGridView(filteredList);
            statusLabel.Text = $"Showing {filteredList.Count} of {_allInvoices.Count} invoices";
        }

        private void FormatDataGridView(List<Invoice> invoices)
        {
            if (invoices == null || invoices.Count == 0) return;

            HideColumn("Id");
            HideColumn("QuickBooksInvoiceId");
            HideColumn("CompanyName");
            HideColumn("FBR_QRCode");
            HideColumn("CreatedDate");
            HideColumn("ModifiedDate");
            HideColumn("CustomerAddress");
            HideColumn("CustomerPhone");
            HideColumn("CustomerEmail");
            HideColumn("TotalAmount");
            HideColumn("TaxAmount");
            HideColumn("DiscountAmount");
            HideColumn("InvoiceDate");
            HideColumn("PaymentMode");
            HideColumn("Items");

            SetColumnHeader("InvoiceNumber", "Invoice #", 15);
            SetColumnHeader("CustomerName", "Customer", 25);
            SetColumnHeader("CustomerNTN", "NTN/CNIC", 18);
            SetColumnHeader("Amount", "Amount", 12, "N2");
            SetColumnHeader("Status", "Status", 10);
            SetColumnHeader("FBR_IRN", "FBR IRN", 20);
            SetColumnHeader("UploadDate", "Upload Date", 18, "dd-MMM-yyyy HH:mm");
            SetColumnHeader("ErrorMessage", "Error", 30);

            // Remove all custom row coloring, keep default DataGridView colors
            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                row.DefaultCellStyle.BackColor = dgvInvoices.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.ForeColor = dgvInvoices.DefaultCellStyle.ForeColor;
                row.DefaultCellStyle.Padding = new Padding(8, 4, 8, 4);
                row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(100, 181, 246);
                row.DefaultCellStyle.SelectionForeColor = Color.White;
            }

            if (dgvInvoices.Columns["Status"] != null)
            {
                dgvInvoices.Columns["Status"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            if (dgvInvoices.Columns["Amount"] != null)
            {
                dgvInvoices.Columns["Amount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        private void HideColumn(string columnName)
        {
            if (dgvInvoices.Columns[columnName] != null)
                dgvInvoices.Columns[columnName].Visible = false;
        }

        private void SetColumnHeader(string columnName, string headerText, int fillWeight, string? format = null)
        {
            if (dgvInvoices.Columns[columnName] != null)
            {
                dgvInvoices.Columns[columnName].HeaderText = headerText;
                dgvInvoices.Columns[columnName].FillWeight = fillWeight;

                if (!string.IsNullOrEmpty(format))
                    dgvInvoices.Columns[columnName].DefaultCellStyle.Format = format;
            }
        }

        private void UpdateStatistics()
        {
            if (_allInvoices == null) return;

            int total = _allInvoices.Count;
            int pending = _allInvoices.Count(i => i.Status == "Pending");
            int uploaded = _allInvoices.Count(i => i.Status == "Uploaded");
            int failed = _allInvoices.Count(i => i.Status == "Failed");

            statusStats.Text = $"Total: {total} | Pending: {pending} | Uploaded: {uploaded} | Failed: {failed}";
        }

        private void CboStatusFilter_SelectedIndexChanged(object? sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void TxtSearchInvoice_TextChanged(object? sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void DgvInvoices_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var invoice = dgvInvoices.Rows[e.RowIndex].DataBoundItem as Invoice;
            if (invoice == null) return;

            ShowInvoiceDetails(invoice);
        }

        private async void BtnFetchInvoices_Click(object? sender, EventArgs e)
        {
            if (_currentCompany == null)
            {
                MessageBox.Show("Please configure company settings first.",
                    "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                statusLabel.Text = "Fetching invoices from QuickBooks...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
                btnFetchInvoices.Enabled = false;

                await Task.Run(() =>
                {
                    _invoiceManager.FetchFromQuickBooks();
                });

                LoadInvoices();
                MessageBox.Show("Invoices fetched successfully!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error fetching invoices: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                btnFetchInvoices.Enabled = true;
                statusLabel.Text = "Ready";
            }
        }

        private async void BtnUploadSelected_Click(object? sender, EventArgs e)
        {
            if (_currentCompany == null || string.IsNullOrEmpty(_currentCompany.FBRToken))
            {
                MessageBox.Show("Please configure company FBR token first.",
                    "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(_currentCompany.SellerNTN))
            {
                MessageBox.Show("Please configure Seller NTN in company settings.",
                    "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(_currentCompany.SellerProvince))
            {
                MessageBox.Show("Please configure Seller Province in company settings.",
                    "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedRows = dgvInvoices.SelectedRows;
            if (selectedRows.Count == 0)
            {
                MessageBox.Show("Please select invoices to upload.",
                    "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            progressBar.Visible = true;
            progressBar.Maximum = selectedRows.Count;
            progressBar.Value = 0;
            progressBar.Style = ProgressBarStyle.Blocks;

            int uploaded = 0;
            int failed = 0;
            var failedInvoices = new List<(string InvoiceNumber, string Error)>();

            foreach (DataGridViewRow row in selectedRows)
            {
                var selectedInvoice = row.DataBoundItem as Invoice;
                if (selectedInvoice != null && selectedInvoice.Status != "Uploaded")
                {
                    var invoice = _db.GetInvoices(_currentCompany.CompanyName)
                                     .FirstOrDefault(i => i.Id == selectedInvoice.Id);

                    if (invoice == null)
                    {
                        failed++;
                        failedInvoices.Add((selectedInvoice.InvoiceNumber, "Invoice not found in database"));
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

                            var updatedInvoice = _db.GetInvoices(_currentCompany.CompanyName)
                                .FirstOrDefault(i => i.Id == invoice.Id);

                            string errorMsg = !string.IsNullOrEmpty(updatedInvoice?.ErrorMessage)
                                ? updatedInvoice.ErrorMessage
                                : response.ErrorMessage ?? "Unknown error - no error message returned";

                            failedInvoices.Add((invoice.InvoiceNumber, errorMsg));

                            System.Diagnostics.Debug.WriteLine($"Failed to upload {invoice.InvoiceNumber}: {errorMsg}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        string errorMsg = $"Exception: {ex.Message}";
                        failedInvoices.Add((invoice.InvoiceNumber, errorMsg));

                        System.Diagnostics.Debug.WriteLine($"Exception uploading {invoice.InvoiceNumber}: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine(ex.StackTrace);
                    }

                    progressBar.Value++;
                }
            }

            LoadInvoices();
            progressBar.Visible = false;
            statusLabel.Text = $"Upload complete: {uploaded} success, {failed} failed";

            if (failedInvoices.Count > 0)
            {
                var resultForm = new Form
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

                var lblFailedTitle = new Label
                {
                    Text = "Failed Invoices Details:",
                    Location = new Point(20, 60),
                    Size = new Size(200, 20),
                    Font = new Font("Arial", 10F, FontStyle.Bold),
                    ForeColor = Color.DarkRed,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left
                };

                var txtFailed = new TextBox
                {
                    Text = string.Join(Environment.NewLine + new string('=', 80) + Environment.NewLine,
                        failedInvoices.Select(f =>
                            $"Invoice: {f.InvoiceNumber}\n" +
                            $"Error:\n{f.Error}\n")),
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

                var btnCopyErrors = new Button
                {
                    Text = "Copy Errors",
                    Size = new Size(120, 35),
                    BackColor = Color.FromArgb(33, 150, 243),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left
                };
                btnCopyErrors.Location = new Point(20, resultForm.ClientSize.Height - 55);
                btnCopyErrors.Click += (s, e) =>
                {
                    Clipboard.SetText(txtFailed.Text);
                    MessageBox.Show("Error details copied to clipboard!", "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                resultForm.Controls.Add(lblSummary);
                resultForm.Controls.Add(lblFailedTitle);
                resultForm.Controls.Add(txtFailed);
                resultForm.Controls.Add(btnCopyErrors);
                resultForm.Controls.Add(btnClose);

                resultForm.ShowDialog();
            }
            else
            {
                MessageBox.Show($"✅ All {uploaded} invoice(s) uploaded successfully!",
                    "Upload Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async void ShowInvoiceDetails(Invoice invoice)
        {
            try
            {
                if (_currentCompany == null)
                {
                    MessageBox.Show("Company settings not loaded. Please configure company setup first.",
                        "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!_qb.Connect(_currentCompany))
                {
                    MessageBox.Show("Failed to connect to QuickBooks.",
                        "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var details = await _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);

                if (details == null)
                {
                    MessageBox.Show("Could not retrieve invoice details from QuickBooks.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var validationErrors = new List<string>();

                if (string.IsNullOrEmpty(details.SellerNTN))
                    validationErrors.Add("⚠️ Seller NTN is missing");

                if (string.IsNullOrEmpty(details.SellerProvince))
                    validationErrors.Add("⚠️ Seller Province is missing");

                if (string.IsNullOrEmpty(details.CustomerNTN))
                    validationErrors.Add("⚠️ Customer NTN is missing");

                if (string.IsNullOrEmpty(details.BuyerProvince))
                    validationErrors.Add("⚠️ Customer Province is missing");

                foreach (var item in details.Items)
                {
                    if (string.IsNullOrEmpty(item.HSCode))
                        validationErrors.Add($"⚠️ HS Code missing for item: {item.ItemName}");

                    if (item.RetailPrice == 0)
                        validationErrors.Add($"⚠️ Retail Price missing for item: {item.ItemName}");
                }

                var fbrPayload = _fbr.BuildFBRPayload(details);
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(fbrPayload, Newtonsoft.Json.Formatting.Indented);

                var detailForm = new Form
                {
                    Text = $"Invoice JSON - {invoice.InvoiceNumber}",
                    Size = new Size(1000, 750),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.Sizable,
                    MinimizeBox = false,
                    MaximizeBox = true
                };

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

                var buttonPanel = new Panel
                {
                    Dock = DockStyle.Bottom,
                    Height = 50,
                    BackColor = Color.FromArgb(240, 240, 240)
                };

                var btnCopyJson = new Button
                {
                    Text = "Copy JSON",
                    Size = new Size(120, 35),
                    Location = new Point(10, 8),
                    BackColor = Color.FromArgb(33, 150, 243),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnCopyJson.Click += (s, e) =>
                {
                    Clipboard.SetText(json);
                    MessageBox.Show("JSON copied to clipboard!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                var btnValidate = new Button
                {
                    Text = "Validate",
                    Size = new Size(100, 35),
                    Location = new Point(140, 8),
                    BackColor = validationErrors.Count == 0 ? Color.FromArgb(46, 125, 50) : Color.FromArgb(255, 152, 0),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnValidate.Click += (s, e) =>
                {
                    if (validationErrors.Count == 0)
                    {
                        MessageBox.Show("✅ All required fields are present!", "Validation Passed",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"⚠️ Found {validationErrors.Count} validation warning(s):\n\n" +
                            string.Join("\n", validationErrors),
                            "Validation Warnings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
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
                btnClose.Click += (s, e) => detailForm.Close();

                buttonPanel.Controls.Add(btnCopyJson);
                buttonPanel.Controls.Add(btnValidate);
                buttonPanel.Controls.Add(btnClose);

                detailForm.Controls.Add(txtJson);
                detailForm.Controls.Add(buttonPanel);

                if (validationErrors.Count > 0)
                {
                    var warningPanel = new Panel
                    {
                        Dock = DockStyle.Top,
                        Height = Math.Min(150, validationErrors.Count * 25 + 50),
                        BackColor = Color.FromArgb(255, 243, 205),
                        BorderStyle = BorderStyle.FixedSingle
                    };

                    var lblWarningTitle = new Label
                    {
                        Text = $"⚠️ VALIDATION WARNINGS ({validationErrors.Count})",
                        Location = new Point(10, 10),
                        Size = new Size(950, 20),
                        Font = new Font("Arial", 10F, FontStyle.Bold),
                        ForeColor = Color.FromArgb(102, 60, 0)
                    };

                    var txtWarnings = new TextBox
                    {
                        Text = string.Join(Environment.NewLine, validationErrors),
                        Location = new Point(10, 35),
                        Size = new Size(960, warningPanel.Height - 45),
                        Multiline = true,
                        ReadOnly = true,
                        ScrollBars = ScrollBars.Vertical,
                        BackColor = Color.FromArgb(255, 243, 205),
                        ForeColor = Color.FromArgb(102, 60, 0),
                        BorderStyle = BorderStyle.None,
                        Font = new Font("Arial", 9F)
                    };

                    warningPanel.Controls.Add(lblWarningTitle);
                    warningPanel.Controls.Add(txtWarnings);
                    detailForm.Controls.Add(warningPanel);
                }

                detailForm.ShowDialog();
            }
            catch (Exception ex)
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

                var errorForm = new Form
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

                var btnCloseError = new Button
                {
                    Text = "Close",
                    Dock = DockStyle.Bottom,
                    Height = 40,
                    BackColor = Color.FromArgb(211, 47, 47),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnCloseError.Click += (s, e) => errorForm.Close();

                errorForm.Controls.Add(txtError);
                errorForm.Controls.Add(btnCloseError);
                errorForm.ShowDialog();
            }
        }

        private async void BtnUploadAll_Click(object? sender, EventArgs e)
        {
            if (_currentCompany == null || string.IsNullOrEmpty(_currentCompany.FBRToken))
            {
                MessageBox.Show("Please configure company FBR token first.",
                    "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(_currentCompany.SellerNTN))
            {
                MessageBox.Show("Please configure Seller NTN in company settings.",
                    "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var pendingInvoices = _allInvoices?.Where(i => i.Status == "Pending").ToList();

            if (pendingInvoices == null || pendingInvoices.Count == 0)
            {
                MessageBox.Show("No pending invoices to upload.",
                    "Nothing to Upload", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show($"Upload {pendingInvoices.Count} pending invoices to FBR?",
                "Confirm Upload", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes) return;

            progressBar.Visible = true;
            progressBar.Maximum = pendingInvoices.Count;
            progressBar.Value = 0;
            progressBar.Style = ProgressBarStyle.Blocks;

            int uploaded = 0;
            int failed = 0;

            foreach (var invoice in pendingInvoices)
            {
                statusLabel.Text = $"Uploading {invoice.InvoiceNumber}...";

                try
                {
                    var response = await _invoiceManager.UploadToFBR(invoice, _currentCompany.FBRToken);

                    if (response.Success)
                        uploaded++;
                    else
                        failed++;
                }
                catch (Exception)
                {
                    failed++;
                }

                progressBar.Value++;
            }

            LoadInvoices();
            progressBar.Visible = false;
            statusLabel.Text = $"Upload complete: {uploaded} success, {failed} failed";

            MessageBox.Show($"Upload complete!\nSuccess: {uploaded}\nFailed: {failed}",
                "Upload Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnGeneratePDF_Click(object? sender, EventArgs e)
        {
            if (dgvInvoices.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select an invoice to generate PDF.",
                    "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var invoice = dgvInvoices.SelectedRows[0].DataBoundItem as Invoice;
            if (invoice == null) return;

            if (invoice.Status != "Uploaded")
            {
                MessageBox.Show("Only uploaded invoices can generate FBR compliant PDFs.",
                    "Not Uploaded", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "PDF Files|*.pdf";
                saveDialog.FileName = $"Invoice_{invoice.InvoiceNumber}_{DateTime.Now:yyyyMMdd}.pdf";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _invoiceManager.GeneratePDF(invoice, saveDialog.FileName);
                        MessageBox.Show("PDF generated successfully!",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = saveDialog.FileName,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error generating PDF: {ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnCompanySetup_Click(object? sender, EventArgs e)
        {
            ShowCompanySetup();
        }

        private void ShowCompanySetup()
        {
            _currentCompany = _companyManager.GetCompany(_qb.CurrentCompanyName);

            if (_currentCompany == null)
            {
                _currentCompany = new Company
                {
                    CompanyName = _qb.CurrentCompanyName
                };
            }

            try
            {
                var qbCompanyInfo = _qb.GetCompanyInfo();
                if (qbCompanyInfo != null)
                {
                    if (string.IsNullOrEmpty(_currentCompany.SellerAddress))
                    {
                        _currentCompany.SellerAddress = qbCompanyInfo.Address;
                    }

                    if (string.IsNullOrEmpty(_currentCompany.SellerPhone))
                    {
                        _currentCompany.SellerPhone = qbCompanyInfo.Phone;
                    }

                    if (string.IsNullOrEmpty(_currentCompany.SellerEmail))
                    {
                        _currentCompany.SellerEmail = qbCompanyInfo.Email;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Warning: Could not fetch company info from QuickBooks: {ex.Message}\n\nYou can still configure the FBR token and NTN.",
                    "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            using (var setupForm = new CompanySetupForm(_qb.CurrentCompanyName, _currentCompany))
            {
                if (setupForm.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _currentCompany = setupForm.Company;
                        _companyManager.SaveCompany(_currentCompany);
                        _invoiceManager.SetCompany(_currentCompany);
                        LoadInvoices();
                        MessageBox.Show("Company settings saved successfully!", "Success",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error saving company settings: {ex.Message}", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ShowFBRError(string invoiceNumber, string errorMessage, string? responseData = null)
        {
            var errorForm = new Form
            {
                Text = $"FBR Upload Error - Invoice {invoiceNumber}",
                Size = new Size(900, 600),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable
            };

            var lblTitle = new Label
            {
                Text = $"❌ Failed to upload Invoice #{invoiceNumber} to FBR",
                Location = new Point(20, 20),
                Size = new Size(850, 30),
                Font = new Font("Arial", 12F, FontStyle.Bold),
                ForeColor = Color.DarkRed
            };

            var lblErrorLabel = new Label
            {
                Text = "Error Message:",
                Location = new Point(20, 60),
                Size = new Size(120, 20),
                Font = new Font("Arial", 9F, FontStyle.Bold)
            };

            var txtError = new TextBox
            {
                Text = errorMessage,
                Location = new Point(20, 85),
                Size = new Size(850, 80),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Arial", 9F),
                BackColor = Color.FromArgb(255, 235, 238),
                ForeColor = Color.DarkRed
            };

            errorForm.Controls.Add(lblTitle);
            errorForm.Controls.Add(lblErrorLabel);
            errorForm.Controls.Add(txtError);

            if (!string.IsNullOrEmpty(responseData))
            {
                var lblResponseLabel = new Label
                {
                    Text = "FBR Response Data:",
                    Location = new Point(20, 180),
                    Size = new Size(150, 20),
                    Font = new Font("Arial", 9F, FontStyle.Bold)
                };

                var txtResponse = new TextBox
                {
                    Text = responseData,
                    Location = new Point(20, 205),
                    Size = new Size(850, 300),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    Font = new Font("Consolas", 9F),
                    BackColor = Color.White,
                    WordWrap = false
                };

                errorForm.Controls.Add(lblResponseLabel);
                errorForm.Controls.Add(txtResponse);
            }

            var btnClose = new Button
            {
                Text = "Close",
                Location = new Point(770, 520),
                Size = new Size(100, 35),
                BackColor = Color.FromArgb(211, 47, 47),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnClose.Click += (s, e) => errorForm.Close();

            errorForm.Controls.Add(btnClose);
            errorForm.ShowDialog();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            try
            {
                _qb?.Dispose();
                _fbr?.Dispose();
                System.Diagnostics.Debug.WriteLine("✅ Final cleanup completed");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Error in final cleanup: {ex.Message}");
            }
        }
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