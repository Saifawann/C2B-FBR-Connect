using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C2B_FBR_Connect.Forms
{
    public partial class MainForm : Form
    {
        private DataGridView dgvInvoices;
        private Button btnFetchInvoices;
        private Button btnUploadSelected;
        private Button btnUploadAll;
        private Button btnGeneratePDF;
        private Button btnCompanySetup;
        private Button btnRefresh;
        private Label lblCompanyName;
        private Label lblFilter;
        private ComboBox cboStatusFilter;
        private TextBox txtSearchInvoice;
        private Label lblSearch;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;
        private ToolStripProgressBar progressBar;
        private ToolStripStatusLabel statusStats;

        private DatabaseService _db;
        private QuickBooksService _qb;
        private FBRApiService _fbr;
        private PDFService _pdf;
        private CompanyManager _companyManager;
        private InvoiceManager _invoiceManager;
        private Company _currentCompany;
        private List<Invoice> _allInvoices;

        public MainForm()
        {
            InitializeComponent();
            SetupCustomUI();
            InitializeServices();
        }

        private void SetupCustomUI()
        {
            this.Text = "C2B FBR Connect - Digital Invoicing System";
            this.Size = new Size(1400, 750);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Company Label
            lblCompanyName = new Label
            {
                Text = "No QuickBooks Company Connected",
                Location = new Point(12, 12),
                Size = new Size(500, 25),
                Font = new Font("Arial", 11F, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };

            // Filter Section
            lblFilter = new Label
            {
                Text = "Status Filter:",
                Location = new Point(12, 45),
                Size = new Size(80, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };

            cboStatusFilter = new ComboBox
            {
                Location = new Point(95, 45),
                Size = new Size(120, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStatusFilter.Items.AddRange(new object[] { "All", "Pending", "Uploaded", "Failed" });
            cboStatusFilter.SelectedIndex = 0;
            cboStatusFilter.SelectedIndexChanged += CboStatusFilter_SelectedIndexChanged;

            // Search Section
            lblSearch = new Label
            {
                Text = "Search:",
                Location = new Point(230, 45),
                Size = new Size(55, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };

            txtSearchInvoice = new TextBox
            {
                Location = new Point(290, 45),
                Size = new Size(200, 23),
                PlaceholderText = "Invoice # or Customer Name"
            };
            txtSearchInvoice.TextChanged += TxtSearchInvoice_TextChanged;

            // Action Buttons
            btnFetchInvoices = new Button
            {
                Text = "Fetch Invoices",
                Location = new Point(12, 80),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnFetchInvoices.Click += BtnFetchInvoices_Click;

            btnRefresh = new Button
            {
                Text = "Refresh",
                Location = new Point(142, 80),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(46, 125, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnRefresh.Click += (s, e) => LoadInvoices();

            btnUploadSelected = new Button
            {
                Text = "Upload Selected",
                Location = new Point(242, 80),
                Size = new Size(130, 35),
                BackColor = Color.FromArgb(255, 152, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnUploadSelected.Click += BtnUploadSelected_Click;

            btnUploadAll = new Button
            {
                Text = "Upload All Pending",
                Location = new Point(382, 80),
                Size = new Size(140, 35),
                BackColor = Color.FromArgb(255, 87, 34),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnUploadAll.Click += BtnUploadAll_Click;

            btnGeneratePDF = new Button
            {
                Text = "Generate PDF",
                Location = new Point(532, 80),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(156, 39, 176),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnGeneratePDF.Click += BtnGeneratePDF_Click;

            btnCompanySetup = new Button
            {
                Text = "Company Setup",
                Location = new Point(662, 80),
                Size = new Size(130, 35),
                BackColor = Color.FromArgb(96, 125, 139),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCompanySetup.Click += BtnCompanySetup_Click;

            // DataGridView
            dgvInvoices = new DataGridView
            {
                Location = new Point(12, 125),
                Size = new Size(1360, 540),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.Fixed3D,
                RowHeadersVisible = false,
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(240, 240, 240) }
            };
            dgvInvoices.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(52, 73, 94),
                ForeColor = Color.White,
                Font = new Font("Arial", 9F, FontStyle.Bold),
                Padding = new Padding(5)
            };
            dgvInvoices.EnableHeadersVisualStyles = false;
            dgvInvoices.CellDoubleClick += DgvInvoices_CellDoubleClick;

            // Status Strip
            statusStrip = new StatusStrip();
            statusLabel = new ToolStripStatusLabel("Ready")
            {
                Spring = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            statusStats = new ToolStripStatusLabel("Total: 0 | Pending: 0 | Uploaded: 0 | Failed: 0")
            {
                BorderSides = ToolStripStatusLabelBorderSides.Left,
                Padding = new Padding(10, 0, 0, 0)
            };
            progressBar = new ToolStripProgressBar { Visible = false, Size = new Size(150, 16) };

            statusStrip.Items.Add(statusLabel);
            statusStrip.Items.Add(progressBar);
            statusStrip.Items.Add(statusStats);

            // Add controls
            this.Controls.Add(lblCompanyName);
            this.Controls.Add(lblFilter);
            this.Controls.Add(cboStatusFilter);
            this.Controls.Add(lblSearch);
            this.Controls.Add(txtSearchInvoice);
            this.Controls.Add(btnFetchInvoices);
            this.Controls.Add(btnRefresh);
            this.Controls.Add(btnUploadSelected);
            this.Controls.Add(btnUploadAll);
            this.Controls.Add(btnGeneratePDF);
            this.Controls.Add(btnCompanySetup);
            this.Controls.Add(dgvInvoices);
            this.Controls.Add(statusStrip);
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

        private void ConnectToQuickBooks()
        {
            try
            {
                _currentCompany = null;

                if (_qb.Connect())
                {
                    lblCompanyName.Text = $"Connected: {_qb.CurrentCompanyName}";
                    _currentCompany = _companyManager.GetCompany(_qb.CurrentCompanyName);

                    if (_currentCompany == null)
                    {
                        MessageBox.Show("Company not configured. Please setup FBR token and seller information.",
                            "Setup Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ShowCompanySetup();
                    }
                    else
                    {
                        // Set company context in invoice manager
                        _invoiceManager.SetCompany(_currentCompany);
                        LoadInvoices();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"QuickBooks connection failed: {ex.Message}",
                    "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadInvoices()
        {
            if (_currentCompany == null) return;

            try
            {
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
        }

        private void ApplyFilters()
        {
            if (_allInvoices == null) return;

            var filtered = _allInvoices.AsEnumerable();

            // Status filter
            string statusFilter = cboStatusFilter.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(statusFilter) && statusFilter != "All")
            {
                filtered = filtered.Where(i => i.Status == statusFilter);
            }

            // Search filter
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

            // Hide unnecessary columns
            HideColumn("Id");
            HideColumn("QuickBooksInvoiceId");
            HideColumn("CompanyName");
            //HideColumn("ErrorMessage");
            HideColumn("FBR_QRCode");
            HideColumn("CreatedDate");

            // Set column headers and formatting
            SetColumnHeader("InvoiceNumber", "Invoice #", 120);
            SetColumnHeader("CustomerName", "Customer", 200);
            SetColumnHeader("CustomerNTN", "NTN/CNIC", 150);
            SetColumnHeader("Amount", "Amount", 120, "N2");
            SetColumnHeader("Status", "Status", 100);
            SetColumnHeader("FBR_IRN", "FBR IRN", 180);
            SetColumnHeader("UploadDate", "Upload Date", 150, "dd-MMM-yyyy HH:mm");

            // Color code rows by status
            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                var status = row.Cells["Status"]?.Value?.ToString();
                switch (status)
                {
                    case "Uploaded":
                        row.DefaultCellStyle.BackColor = Color.FromArgb(200, 250, 205);
                        break;
                    case "Failed":
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 205, 210);
                        break;
                    case "Pending":
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 225);
                        break;
                }
            }
        }

        private void HideColumn(string columnName)
        {
            if (dgvInvoices.Columns[columnName] != null)
                dgvInvoices.Columns[columnName].Visible = false;
        }

        private void SetColumnHeader(string columnName, string headerText, int width = -1, string format = null)
        {
            if (dgvInvoices.Columns[columnName] != null)
            {
                dgvInvoices.Columns[columnName].HeaderText = headerText;
                if (width > 0)
                {
                    dgvInvoices.Columns[columnName].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    dgvInvoices.Columns[columnName].Width = width;
                }
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

        private void CboStatusFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void TxtSearchInvoice_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void DgvInvoices_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var invoice = dgvInvoices.Rows[e.RowIndex].DataBoundItem as Invoice;
            if (invoice == null) return;

            ShowInvoiceDetails(invoice);
        }

        private void ShowInvoiceDetails(Invoice invoice)
        {
            try
            {
                // Ensure QuickBooks service has current company settings
                if (_currentCompany == null)
                {
                    MessageBox.Show("Company settings not loaded. Please configure company setup first.",
                        "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Make sure QB service has the company context
                if (!_qb.Connect(_currentCompany))
                {
                    MessageBox.Show("Failed to connect to QuickBooks.",
                        "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Fetch full invoice details from QuickBooks
                var details = _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);

                if (details == null)
                {
                    MessageBox.Show("Could not retrieve invoice details from QuickBooks.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Build JSON-like object (no defaults, just values from QuickBooks)
                var invoiceObject = new
                {
                    invoiceType = details.InvoiceType,
                    invoiceDate = details.InvoiceDate.ToString("yyyy-MM-dd"),
                    sellerNTNCNIC = details.SellerNTN,
                    sellerBusinessName = details.SellerBusinessName,
                    sellerProvince = details.SellerProvince,
                    sellerAddress = details.SellerAddress,
                    buyerNTNCNIC = details.CustomerNTN,
                    buyerBusinessName = details.CustomerName,
                    buyerProvince = details.BuyerProvince,
                    buyerAddress = details.BuyerAddress,
                    buyerRegistrationType = details.BuyerRegistrationType,
                    invoiceRefNo = "",
                    scenarioId = details.BuyerRegistrationType == "Registered" ? "SN001" : "SN002",
                    items = details.Items.Select(item => new
                    {
                        hsCode = item.HSCode,
                        productDescription = item.ItemName,
                        rate = item.TaxRate > 0 ? $"{item.TaxRate:N0}%" : null,
                        uoM = item.UnitOfMeasure,
                        quantity = item.Quantity,
                        totalValues = item.TotalValue,
                        valueSalesExcludingST = item.TotalPrice,
                        fixedNotifiedValueOrRetailPrice = item.RetailPrice,
                        salesTaxApplicable = item.SalesTaxAmount,
                        extraTax = item.ExtraTax,
                        furtherTax = item.FurtherTax, 
                        discount = 0.00,
                        fedPayable= 0.00,
                        salesTaxWithheldAtSource = 0.00,
                        saleType = "Goods at standard rate (default)"
                    }).ToList()
                };

                // Convert to formatted JSON string
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(invoiceObject, Newtonsoft.Json.Formatting.Indented);

                // Show JSON in form
                var detailForm = new Form
                {
                    Text = $"Invoice JSON - {invoice.InvoiceNumber}",
                    Size = new Size(900, 700),
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

                var btnClose = new Button
                {
                    Text = "Close",
                    Size = new Size(100, 35),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                    Location = new Point(detailForm.ClientSize.Width - 110, detailForm.ClientSize.Height - 45),
                    BackColor = Color.FromArgb(96, 125, 139),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnClose.Click += (s, e) => detailForm.Close();

                detailForm.Controls.Add(txtJson);
                detailForm.Controls.Add(btnClose);
                detailForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error displaying invoice details: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private async void BtnFetchInvoices_Click(object sender, EventArgs e)
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

        private async void BtnUploadSelected_Click(object sender, EventArgs e)
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

            foreach (DataGridViewRow row in selectedRows)
            {
                var invoice = row.DataBoundItem as Invoice;
                if (invoice != null && invoice.Status != "Uploaded")
                {
                    statusLabel.Text = $"Uploading {invoice.InvoiceNumber}...";

                    try
                    {
                        bool success = await _invoiceManager.UploadToFBR(invoice, _currentCompany.FBRToken);
                        if (success) uploaded++;
                        else failed++;
                    }
                    catch
                    {
                        failed++;
                    }

                    progressBar.Value++;
                }
            }

            LoadInvoices();
            progressBar.Visible = false;
            statusLabel.Text = $"Upload complete: {uploaded} success, {failed} failed";

            MessageBox.Show($"Upload complete!\nSuccess: {uploaded}\nFailed: {failed}",
                "Upload Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private async void BtnUploadAll_Click(object sender, EventArgs e)
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
                    bool success = await _invoiceManager.UploadToFBR(invoice, _currentCompany.FBRToken);
                    if (success) uploaded++;
                    else failed++;
                }
                catch
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

        private void BtnGeneratePDF_Click(object sender, EventArgs e)
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

        private void BtnCompanySetup_Click(object sender, EventArgs e)
        {
            ShowCompanySetup();
        }

        private void ShowCompanySetup()
        {
            // Always reload company from database to get latest saved values
            _currentCompany = _companyManager.GetCompany(_qb.CurrentCompanyName);

            // If company doesn't exist in database, create new one
            if (_currentCompany == null)
            {
                _currentCompany = new Company
                {
                    CompanyName = _qb.CurrentCompanyName
                };
            }

            // Only update address from QuickBooks (NOT province)
            // Province should only come from user input, never from QuickBooks
            try
            {
                var qbCompanyInfo = _qb.GetCompanyInfo();
                if (qbCompanyInfo != null)
                {
                    // Only set address if it's empty or hasn't been set before
                    if (string.IsNullOrEmpty(_currentCompany.SellerAddress))
                    {
                        _currentCompany.SellerAddress = qbCompanyInfo.Address;
                    }

                    // ✅ NEVER set SellerProvince from QuickBooks
                    // User must enter it manually in the form
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

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            _qb?.Dispose();
            _fbr?.Dispose();
        }
    }
}