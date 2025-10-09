using C2B_FBR_Connect.Helpers;
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
            InvoiceGridHelper.FormatDataGridView(dgvInvoices, invoices);
            // Hide and format columns as before
            InvoiceGridHelper.HideColumn(dgvInvoices, "Id");
            InvoiceGridHelper.HideColumn(dgvInvoices, "QuickBooksInvoiceId");
            InvoiceGridHelper.HideColumn(dgvInvoices, "CompanyName");
            //InvoiceGridHelper.HideColumn(dgvInvoices, "ErrorMessage");
            InvoiceGridHelper.HideColumn(dgvInvoices, "FBR_QRCode");
            InvoiceGridHelper.HideColumn(dgvInvoices, "CreatedDate");

            InvoiceGridHelper.SetColumnHeader(dgvInvoices, "InvoiceNumber", "Invoice #", 120);
            InvoiceGridHelper.SetColumnHeader(dgvInvoices, "CustomerName", "Customer", 200);
            InvoiceGridHelper.SetColumnHeader(dgvInvoices, "CustomerNTN", "NTN/CNIC", 150);
            InvoiceGridHelper.SetColumnHeader(dgvInvoices, "Amount", "Amount", 120, "N2");
            InvoiceGridHelper.SetColumnHeader(dgvInvoices, "Status", "Status", 100);
            InvoiceGridHelper.SetColumnHeader(dgvInvoices, "FBR_IRN", "FBR IRN", 180);
            InvoiceGridHelper.SetColumnHeader(dgvInvoices, "UploadDate", "Upload Date", 150, "dd-MMM-yyyy HH:mm");
        }

        private void HideColumn(string columnName)
        {
            InvoiceGridHelper.HideColumn(dgvInvoices, columnName);
        }

        private void SetColumnHeader(string columnName, string headerText, int width = -1, string format = null)
        {
            InvoiceGridHelper.SetColumnHeader(dgvInvoices, columnName, headerText, width, format);
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
                //var invoice = row.DataBoundItem as Invoice;
                //if (invoice != null && invoice.Status != "Uploaded")
                //{
                //    statusLabel.Text = $"Uploading {invoice.InvoiceNumber}...";

                //    try
                //    {
                //        bool success = await _invoiceManager.UploadToFBR(invoice, _currentCompany.FBRToken);
                //        if (success)
                //        {
                //            uploaded++;
                //        }
                //        else
                //        {
                //            failed++;
                //            // Get error from database
                //            var updatedInvoice = _db.GetInvoices(_currentCompany.CompanyName)
                //                .FirstOrDefault(i => i.Id == invoice.Id);

                //            failedInvoices.Add((invoice.InvoiceNumber, updatedInvoice?.ErrorMessage ?? "Unknown error"));
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        failed++;
                //        failedInvoices.Add((invoice.InvoiceNumber, ex.Message));
                //    }

                //    progressBar.Value++;
                //}
                var selectedInvoice = row.DataBoundItem as Invoice;
                if (selectedInvoice != null && selectedInvoice.Status != "Uploaded")
                {
                    // ✅ Always use the invoice as stored in DB
                    var invoice = _db.GetInvoices(_currentCompany.CompanyName)
                                     .FirstOrDefault(i => i.Id == selectedInvoice.Id);

                    if (invoice == null)
                        continue;

                    statusLabel.Text = $"Uploading {invoice.InvoiceNumber}...";

                    try
                    {
                        bool success = await _invoiceManager.UploadToFBR(invoice, _currentCompany.FBRToken);
                        if (success)
                        {
                            uploaded++;
                        }
                        else
                        {
                            failed++;
                            // Get error from database
                            var updatedInvoice = _db.GetInvoices(_currentCompany.CompanyName)
                                .FirstOrDefault(i => i.Id == invoice.Id);

                            failedInvoices.Add((invoice.InvoiceNumber, updatedInvoice?.ErrorMessage ?? "Unknown error"));
                        }
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        failedInvoices.Add((invoice.InvoiceNumber, ex.Message));
                    }

                    progressBar.Value++;
                }
            }

            LoadInvoices();
            progressBar.Visible = false;
            statusLabel.Text = $"Upload complete: {uploaded} success, {failed} failed";

            // Show detailed results
            if (failedInvoices.Count > 0)
            {
                var resultForm = new Form
                {
                    Text = "Upload Results",
                    Size = new Size(700, 500),
                    StartPosition = FormStartPosition.CenterParent
                };

                var lblSummary = new Label
                {
                    Text = $"✅ Uploaded: {uploaded}    ❌ Failed: {failed}",
                    Location = new Point(20, 20),
                    Size = new Size(650, 30),
                    Font = new Font("Arial", 11F, FontStyle.Bold)
                };

                var lblFailedTitle = new Label
                {
                    Text = "Failed Invoices:",
                    Location = new Point(20, 60),
                    Size = new Size(150, 20),
                    Font = new Font("Arial", 9F, FontStyle.Bold)
                };

                var txtFailed = new TextBox
                {
                    Text = string.Join(Environment.NewLine + Environment.NewLine,
                        failedInvoices.Select(f => $"Invoice: {f.InvoiceNumber}\nError: {f.Error}")),
                    Location = new Point(20, 85),
                    Size = new Size(650, 350),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    Font = new Font("Consolas", 9F),
                    BackColor = Color.FromArgb(255, 245, 245)
                };

                var btnClose = new Button
                {
                    Text = "Close",
                    Location = new Point(570, 445),
                    Size = new Size(100, 35),
                    BackColor = Color.FromArgb(96, 125, 139),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnClose.Click += (s, e) => resultForm.Close();

                resultForm.Controls.Add(lblSummary);
                resultForm.Controls.Add(lblFailedTitle);
                resultForm.Controls.Add(txtFailed);
                resultForm.Controls.Add(btnClose);
                resultForm.ShowDialog();
            }
            else
            {
                MessageBox.Show($"✅ All {uploaded} invoice(s) uploaded successfully!",
                    "Upload Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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

                // Validate critical fields before displaying
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

                    if (string.IsNullOrEmpty(item.RetailPrice))
                        validationErrors.Add($"⚠️ Retail Price missing for item: {item.ItemName}");
                }

                // Build JSON object
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
                        rate = item.TaxRate > 0 ? $"{item.TaxRate:N0}%" : "0%",
                        uoM = item.UnitOfMeasure,
                        quantity = item.Quantity,
                        totalValues = item.TotalValue,
                        valueSalesExcludingST = item.TotalPrice,
                        fixedNotifiedValueOrRetailPrice = item.RetailPrice,
                        salesTaxApplicable = item.SalesTaxAmount,
                        extraTax = item.ExtraTax,
                        furtherTax = item.FurtherTax,
                        discount = 0.00,
                        fedPayable = 0.00,
                        salesTaxWithheldAtSource = 0.00,
                        saleType = "Goods at standard rate (default)"
                    }).ToList()
                };

                // Convert to formatted JSON string
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(invoiceObject, Newtonsoft.Json.Formatting.Indented);

                // Create main form
                var detailForm = new Form
                {
                    Text = $"Invoice JSON - {invoice.InvoiceNumber}",
                    Size = new Size(1000, 750),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.Sizable,
                    MinimizeBox = false,
                    MaximizeBox = true
                };

                // JSON display textbox
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

                // Button panel
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

                // Add JSON textbox first, then buttons
                detailForm.Controls.Add(txtJson);
                detailForm.Controls.Add(buttonPanel);

                // Add warning panel last so layout adjusts properly
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
                // Detailed error message
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

                // Show error in a scrollable form
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

        private void ShowFBRError(string invoiceNumber, string errorMessage, string responseData = null)
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
            _qb?.Dispose();
            _fbr?.Dispose();
        }
    }
}