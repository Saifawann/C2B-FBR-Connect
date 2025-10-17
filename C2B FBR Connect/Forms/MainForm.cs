using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C2B_FBR_Connect.Forms
{
    public partial class MainForm : Form
    {
        // UI Controls
        private Panel headerPanel;
        private Panel navigationPanel;
        private Panel contentPanel;
        private Panel filterPanel;
        private DataGridView dgvInvoices;
        private Button btnFetchInvoices;
        private Button btnUploadSelected;
        private Button btnUploadAll;
        private Button btnGeneratePDF;
        private Button btnCompanySetup;
        private Button btnRefresh;
        private Label lblCompanyName;
        private Label lblCompanyStatus;
        private Label lblFilter;
        private ComboBox cboStatusFilter;
        private TextBox txtSearchInvoice;
        private Label lblSearch;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;
        private ToolStripProgressBar progressBar;
        private ToolStripStatusLabel statusStats;
        private PictureBox logoBox;
        private Label lblAppTitle;
        private Panel statsPanel;
        private Label lblTotalInvoices;
        private Label lblPendingCount;
        private Label lblUploadedCount;
        private Label lblFailedCount;

        // Services
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
            SetupEnhancedUI();
            InitializeServices();
        }

        private void SetupEnhancedUI()
        {
            this.Text = "C2B Smart App - Digital Invoicing System";
            this.Size = new Size(1450, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 247, 250);
            this.Font = new Font("Segoe UI", 9F);

            // Header Panel
            headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = Color.White
            };
            headerPanel.Paint += HeaderPanel_Paint;

            // Logo (placeholder)
            logoBox = new PictureBox
            {
                Location = new Point(20, 15),
                Size = new Size(50, 50),
                BackColor = Color.FromArgb(0, 122, 204),
                SizeMode = PictureBoxSizeMode.CenterImage
            };
            logoBox.Paint += LogoBox_Paint;

            // App Title
            lblAppTitle = new Label
            {
                Text = "C2B Smart Digital Invoicing",
                Location = new Point(85, 15),
                Size = new Size(400, 30),
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41)
            };

            // Company Info Section
            lblCompanyName = new Label
            {
                Text = "No QuickBooks Company Connected",
                Location = new Point(85, 45),
                Size = new Size(400, 20),
                Font = new Font("Segoe UI", 10F),
                ForeColor = Color.FromArgb(108, 117, 125)
            };

            lblCompanyStatus = new Label
            {
                Text = "● Disconnected",
                Location = new Point(500, 30),
                Size = new Size(150, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69),
                TextAlign = ContentAlignment.MiddleRight
            };

            headerPanel.Controls.Add(logoBox);
            headerPanel.Controls.Add(lblAppTitle);
            headerPanel.Controls.Add(lblCompanyName);
            headerPanel.Controls.Add(lblCompanyStatus);

            // Statistics Panel
            statsPanel = new Panel
            {
                Location = new Point(900, 10),
                Size = new Size(520, 60),
                BackColor = Color.Transparent
            };

            CreateStatCard(statsPanel, 0, "Total", "0", Color.FromArgb(73, 80, 87), ref lblTotalInvoices);
            CreateStatCard(statsPanel, 130, "Pending", "0", Color.FromArgb(255, 193, 7), ref lblPendingCount);
            CreateStatCard(statsPanel, 260, "Uploaded", "0", Color.FromArgb(40, 167, 69), ref lblUploadedCount);
            CreateStatCard(statsPanel, 390, "Failed", "0", Color.FromArgb(220, 53, 69), ref lblFailedCount);

            headerPanel.Controls.Add(statsPanel);

            // Navigation Panel
            navigationPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 130,
                BackColor = Color.White,
                Padding = new Padding(20, 15, 20, 15)
            };

            // Filter Panel
            filterPanel = new Panel
            {
                Location = new Point(20, 10),
                Size = new Size(600, 40),
                BackColor = Color.FromArgb(248, 249, 250)
            };
            filterPanel.Paint += Panel_Paint;

            lblFilter = new Label
            {
                Text = "Status:",
                Location = new Point(15, 10),
                Size = new Size(50, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F)
            };

            cboStatusFilter = new ComboBox
            {
                Location = new Point(70, 8),
                Size = new Size(140, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F),
                FlatStyle = FlatStyle.Flat
            };
            cboStatusFilter.Items.AddRange(new object[] { "All Invoices", "Pending", "Uploaded", "Failed" });
            cboStatusFilter.SelectedIndex = 0;
            cboStatusFilter.SelectedIndexChanged += CboStatusFilter_SelectedIndexChanged;

            lblSearch = new Label
            {
                Text = "Search:",
                Location = new Point(230, 10),
                Size = new Size(50, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F)
            };

            txtSearchInvoice = new TextBox
            {
                Location = new Point(285, 8),
                Size = new Size(300, 25),
                Font = new Font("Segoe UI", 9F),
                BorderStyle = BorderStyle.FixedSingle
            };
            txtSearchInvoice.TextChanged += TxtSearchInvoice_TextChanged;

            // Add placeholder text
            SetPlaceholder(txtSearchInvoice, "Search by Invoice # or Customer Name...");

            filterPanel.Controls.Add(lblFilter);
            filterPanel.Controls.Add(cboStatusFilter);
            filterPanel.Controls.Add(lblSearch);
            filterPanel.Controls.Add(txtSearchInvoice);

            // Action Buttons with modern styling
            int buttonY = 60;
            int buttonHeight = 40;
            int buttonSpacing = 10;

            btnFetchInvoices = CreateModernButton(
                "📥 Fetch Invoices",
                new Point(20, buttonY),
                new Size(145, buttonHeight),
                Color.FromArgb(13, 110, 253),
                Color.White
            );
            btnFetchInvoices.Click += BtnFetchInvoices_Click;

            btnRefresh = CreateModernButton(
                "🔄 Refresh",
                new Point(175, buttonY),
                new Size(110, buttonHeight),
                Color.FromArgb(25, 135, 84),
                Color.White
            );
            btnRefresh.Click += (s, e) => LoadInvoices();

            btnUploadSelected = CreateModernButton(
                "📤 Upload Selected",
                new Point(295, buttonY),
                new Size(155, buttonHeight),
                Color.FromArgb(255, 193, 7),
                Color.FromArgb(33, 37, 41)
            );
            btnUploadSelected.Click += BtnUploadSelected_Click;

            btnUploadAll = CreateModernButton(
                "⬆️ Upload All Pending",
                new Point(460, buttonY),
                new Size(170, buttonHeight),
                Color.FromArgb(253, 126, 20),
                Color.White
            );
            btnUploadAll.Click += BtnUploadAll_Click;

            btnGeneratePDF = CreateModernButton(
                "📄 Generate PDF",
                new Point(640, buttonY),
                new Size(145, buttonHeight),
                Color.FromArgb(111, 66, 193),
                Color.White
            );
            btnGeneratePDF.Click += BtnGeneratePDF_Click;

            btnCompanySetup = CreateModernButton(
                "⚙️ Company Setup",
                new Point(795, buttonY),
                new Size(155, buttonHeight),
                Color.FromArgb(108, 117, 125),
                Color.White
            );
            btnCompanySetup.Click += BtnCompanySetup_Click;

            navigationPanel.Controls.Add(filterPanel);
            navigationPanel.Controls.Add(btnFetchInvoices);
            navigationPanel.Controls.Add(btnRefresh);
            navigationPanel.Controls.Add(btnUploadSelected);
            navigationPanel.Controls.Add(btnUploadAll);
            navigationPanel.Controls.Add(btnGeneratePDF);
            navigationPanel.Controls.Add(btnCompanySetup);

            // Content Panel with shadow
            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 0, 20, 20),
                BackColor = Color.FromArgb(245, 247, 250)
            };

            // DataGridView with enhanced styling
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
                GridColor = Color.FromArgb(233, 236, 239),
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    SelectionBackColor = Color.FromArgb(232, 240, 254),
                    SelectionForeColor = Color.Black,
                    Padding = new Padding(8, 4, 8, 4),
                    Font = new Font("Segoe UI", 9F)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(248, 249, 250),
                    SelectionBackColor = Color.FromArgb(232, 240, 254),
                    SelectionForeColor = Color.Black,
                    Padding = new Padding(8, 4, 8, 4)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(248, 249, 250),
                    ForeColor = Color.FromArgb(73, 80, 87),
                    Font = new Font("Segoe UI Semibold", 9.5F),
                    Padding = new Padding(8, 8, 8, 8),
                    Alignment = DataGridViewContentAlignment.MiddleLeft
                },
                RowTemplate = { Height = 45 },
                EnableHeadersVisualStyles = false
            };
            dgvInvoices.CellDoubleClick += DgvInvoices_CellDoubleClick;
            dgvInvoices.Paint += DgvInvoices_Paint;

            // Wrap DataGridView in a panel for shadow effect
            var gridPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0),
                BackColor = Color.White
            };
            gridPanel.Paint += GridPanel_Paint;
            gridPanel.Controls.Add(dgvInvoices);

            contentPanel.Controls.Add(gridPanel);

            // Enhanced Status Strip
            statusStrip = new StatusStrip
            {
                BackColor = Color.White,
                RenderMode = ToolStripRenderMode.Professional
            };

            statusLabel = new ToolStripStatusLabel("Ready")
            {
                Spring = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F)
            };

            progressBar = new ToolStripProgressBar
            {
                Visible = false,
                Size = new Size(200, 16),
                Style = ProgressBarStyle.Blocks
            };

            statusStats = new ToolStripStatusLabel("")
            {
                BorderSides = ToolStripStatusLabelBorderSides.Left,
                Padding = new Padding(10, 0, 0, 0),
                Font = new Font("Segoe UI", 9F)
            };

            statusStrip.Items.Add(statusLabel);
            statusStrip.Items.Add(progressBar);
            statusStrip.Items.Add(statusStats);

            // Add all panels to form
            this.Controls.Add(contentPanel);
            this.Controls.Add(navigationPanel);
            this.Controls.Add(headerPanel);
            this.Controls.Add(statusStrip);
        }

        private Button CreateModernButton(string text, Point location, Size size, Color backColor, Color foreColor)
        {
            var button = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                BackColor = backColor,
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9.5F),
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };

            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = ControlPaint.Dark(backColor, 0.1f);
            button.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(backColor, 0.2f);

            // Add rounded corners
            button.Paint += Button_Paint;

            return button;
        }

        private void CreateStatCard(Panel parent, int x, string label, string value, Color color, ref Label valueLabel)
        {
            var card = new Panel
            {
                Location = new Point(x, 0),
                Size = new Size(120, 60),
                BackColor = Color.White
            };
            card.Paint += StatCard_Paint;

            var lblTitle = new Label
            {
                Text = label,
                Location = new Point(10, 8),
                Size = new Size(100, 18),
                Font = new Font("Segoe UI", 8F),
                ForeColor = Color.FromArgb(108, 117, 125),
                TextAlign = ContentAlignment.MiddleLeft
            };

            valueLabel = new Label
            {
                Text = value,
                Location = new Point(10, 26),
                Size = new Size(100, 25),
                Font = new Font("Segoe UI Semibold", 14F),
                ForeColor = color,
                TextAlign = ContentAlignment.MiddleLeft
            };

            card.Controls.Add(lblTitle);
            card.Controls.Add(valueLabel);
            parent.Controls.Add(card);
        }

        private void SetPlaceholder(TextBox textBox, string placeholder)
        {
            textBox.Text = placeholder;
            textBox.ForeColor = Color.Gray;

            textBox.GotFocus += (s, e) =>
            {
                if (textBox.Text == placeholder)
                {
                    textBox.Text = "";
                    textBox.ForeColor = Color.Black;
                }
            };

            textBox.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(textBox.Text))
                {
                    textBox.Text = placeholder;
                    textBox.ForeColor = Color.Gray;
                }
            };
        }

        // Paint event handlers for custom drawing
        private void HeaderPanel_Paint(object sender, PaintEventArgs e)
        {
            // Draw bottom border
            using (var pen = new Pen(Color.FromArgb(233, 236, 239), 1))
            {
                e.Graphics.DrawLine(pen, 0, headerPanel.Height - 1, headerPanel.Width, headerPanel.Height - 1);
            }
        }

        private void LogoBox_Paint(object sender, PaintEventArgs e)
        {
            // Draw placeholder logo
            var rect = new Rectangle(10, 10, 30, 30);
            using (var brush = new SolidBrush(Color.White))
            using (var font = new Font("Segoe UI", 14F, FontStyle.Bold))
            {
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                e.Graphics.DrawString("C2B", font, brush, rect, format);
            }
        }

        private void Panel_Paint(object sender, PaintEventArgs e)
        {
            var panel = sender as Panel;
            using (var path = GetRoundedRectPath(panel.ClientRectangle, 8))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var brush = new SolidBrush(panel.BackColor))
                {
                    e.Graphics.FillPath(brush, path);
                }
            }
        }

        private void Button_Paint(object sender, PaintEventArgs e)
        {
            var button = sender as Button;
            using (var path = GetRoundedRectPath(button.ClientRectangle, 6))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                button.Region = new Region(path);
            }
        }

        private void StatCard_Paint(object sender, PaintEventArgs e)
        {
            var panel = sender as Panel;
            using (var path = GetRoundedRectPath(panel.ClientRectangle, 8))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                // Fill background
                using (var brush = new SolidBrush(Color.White))
                {
                    e.Graphics.FillPath(brush, path);
                }

                // Draw border
                using (var pen = new Pen(Color.FromArgb(233, 236, 239), 1))
                {
                    e.Graphics.DrawPath(pen, path);
                }
            }
        }

        private void GridPanel_Paint(object sender, PaintEventArgs e)
        {
            // Draw subtle shadow
            var panel = sender as Panel;
            var rect = new Rectangle(2, 2, panel.Width - 4, panel.Height - 4);
            using (var path = GetRoundedRectPath(rect, 8))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                // Shadow
                using (var shadowBrush = new SolidBrush(Color.FromArgb(20, 0, 0, 0)))
                {
                    var shadowRect = new Rectangle(rect.X + 2, rect.Y + 2, rect.Width, rect.Height);
                    using (var shadowPath = GetRoundedRectPath(shadowRect, 8))
                    {
                        e.Graphics.FillPath(shadowBrush, shadowPath);
                    }
                }

                // Background
                using (var brush = new SolidBrush(Color.White))
                {
                    e.Graphics.FillPath(brush, path);
                }
            }
        }

        private void DgvInvoices_Paint(object sender, PaintEventArgs e)
        {
            // Custom painting if needed
        }

        private GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int diameter = radius * 2;

            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();

            return path;
        }

        private void UpdateStatistics()
        {
            if (_allInvoices == null) return;

            int total = _allInvoices.Count;
            int pending = _allInvoices.Count(i => i.Status == "Pending");
            int uploaded = _allInvoices.Count(i => i.Status == "Uploaded");
            int failed = _allInvoices.Count(i => i.Status == "Failed");

            // Update stat cards
            lblTotalInvoices.Text = total.ToString();
            lblPendingCount.Text = pending.ToString();
            lblUploadedCount.Text = uploaded.ToString();
            lblFailedCount.Text = failed.ToString();

            // Update status bar
            statusStats.Text = $"Displaying {dgvInvoices.RowCount} invoice(s)";
        }

        private void FormatDataGridView(List<Invoice> invoices)
        {
            if (invoices == null || invoices.Count == 0) return;

            dgvInvoices.ColumnHeadersHeight = 40; // Adjust height as needed
            dgvInvoices.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;


            // Hide unnecessary columns
            HideColumn("Id");
            HideColumn("QuickBooksInvoiceId");
            HideColumn("CompanyName");
            HideColumn("FBR_QRCode");
            HideColumn("CreatedDate");

            // Set column headers and formatting
            SetColumnHeader("InvoiceNumber", "Invoice #", 120);
            SetColumnHeader("CustomerName", "Customer", 220);
            SetColumnHeader("CustomerNTN", "NTN/CNIC", 150);
            SetColumnHeader("Amount", "Amount", 120, "N2");
            SetColumnHeader("Status", "Status", 100);
            SetColumnHeader("FBR_IRN", "FBR IRN", 200);
            SetColumnHeader("UploadDate", "Upload Date", 160, "dd-MMM-yyyy HH:mm");

            // Add status indicators
            foreach (DataGridViewRow row in dgvInvoices.Rows)
            {
                var status = row.Cells["Status"]?.Value?.ToString();
                var statusCell = row.Cells["Status"];

                if (statusCell != null)
                {
                    switch (status)
                    {
                        case "Uploaded":
                            row.DefaultCellStyle.BackColor = Color.FromArgb(240, 253, 244);
                            statusCell.Style.ForeColor = Color.FromArgb(22, 163, 74);
                            statusCell.Style.Font = new Font("Segoe UI Semibold", 9F);
                            break;
                        case "Failed":
                            row.DefaultCellStyle.BackColor = Color.FromArgb(254, 242, 242);
                            statusCell.Style.ForeColor = Color.FromArgb(220, 53, 69);
                            statusCell.Style.Font = new Font("Segoe UI Semibold", 9F);
                            break;
                        case "Pending":
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 251, 235);
                            statusCell.Style.ForeColor = Color.FromArgb(217, 119, 6);
                            statusCell.Style.Font = new Font("Segoe UI Semibold", 9F);
                            break;
                    }
                }
            }
        }

        // Rest of the methods remain the same but with updated styling...
        // (InitializeServices, ConnectToQuickBooks, LoadInvoices, etc.)

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
                    lblCompanyStatus.Text = "● Disconnected";
                    lblCompanyStatus.ForeColor = Color.FromArgb(220, 53, 69);

                    var retryResult = MessageBox.Show(
                        "Unable to connect to QuickBooks.\nWould you like to retry?",
                        "Connection Failed",
                        MessageBoxButtons.RetryCancel,
                        MessageBoxIcon.Warning);

                    return retryResult == DialogResult.Retry ? ConnectToQuickBooks() : false;
                }

                lblCompanyName.Text = $"Connected: {_qb.CurrentCompanyName}";
                lblCompanyStatus.Text = "● Connected";
                lblCompanyStatus.ForeColor = Color.FromArgb(40, 167, 69);

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
                    _invoiceManager.SetCompany(_currentCompany);
                    LoadInvoices();
                }

                return true;
            }
            catch (Exception ex)
            {
                lblCompanyStatus.Text = "● Error";
                lblCompanyStatus.ForeColor = Color.FromArgb(220, 53, 69);

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
            if (!string.IsNullOrEmpty(statusFilter) && statusFilter != "All Invoices")
            {
                filtered = filtered.Where(i => i.Status == statusFilter);
            }

            // Search filter
            string searchText = txtSearchInvoice.Text.Trim();
            if (!string.IsNullOrEmpty(searchText) &&
                searchText != "Search by Invoice # or Customer Name...")
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

        // Event handlers remain the same...
        private void CboStatusFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void TxtSearchInvoice_TextChanged(object sender, EventArgs e)
        {
            if (txtSearchInvoice.Text != "Search by Invoice # or Customer Name...")
            {
                ApplyFilters();
            }
        }

        private void DgvInvoices_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var invoice = dgvInvoices.Rows[e.RowIndex].DataBoundItem as Invoice;
            if (invoice == null) return;

            ShowInvoiceDetails(invoice);
        }

        // Remaining event handlers and methods stay the same...
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
                        }
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        failedInvoices.Add((invoice.InvoiceNumber, $"Exception: {ex.Message}"));
                    }

                    progressBar.Value++;
                }
            }

            LoadInvoices();
            progressBar.Visible = false;
            statusLabel.Text = $"Upload complete: {uploaded} success, {failed} failed";

            if (failedInvoices.Count > 0)
            {
                ShowUploadResults(uploaded, failed, failedInvoices);
            }
            else
            {
                MessageBox.Show($"✅ All {uploaded} invoice(s) uploaded successfully!",
                    "Upload Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowUploadResults(int uploaded, int failed, List<(string InvoiceNumber, string Error)> failedInvoices)
        {
            var resultForm = new Form
            {
                Text = "Upload Results",
                Width = 900,
                Height = 650,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimumSize = new Size(700, 500),
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9F)
            };

            var lblSummary = new Label
            {
                Text = $"✅ Uploaded: {uploaded}    ❌ Failed: {failed}",
                Location = new Point(20, 20),
                Size = new Size(850, 30),
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = failed > 0 ? Color.FromArgb(220, 53, 69) : Color.FromArgb(40, 167, 69),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var lblFailedTitle = new Label
            {
                Text = "Failed Invoices Details:",
                Location = new Point(20, 60),
                Size = new Size(200, 20),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69),
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
                BackColor = Color.FromArgb(254, 242, 242),
                ForeColor = Color.FromArgb(153, 27, 27),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            var btnCopyErrors = CreateModernButton(
                "📋 Copy Errors",
                new Point(20, resultForm.ClientSize.Height - 55),
                new Size(120, 35),
                Color.FromArgb(33, 150, 243),
                Color.White
            );
            btnCopyErrors.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnCopyErrors.Click += (s, e) =>
            {
                Clipboard.SetText(txtFailed.Text);
                MessageBox.Show("Error details copied to clipboard!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            var btnClose = CreateModernButton(
                "Close",
                new Point(resultForm.ClientSize.Width - 140, resultForm.ClientSize.Height - 55),
                new Size(120, 35),
                Color.FromArgb(108, 117, 125),
                Color.White
            );
            btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnClose.Click += (s, e) => resultForm.Close();

            resultForm.Controls.Add(lblSummary);
            resultForm.Controls.Add(lblFailedTitle);
            resultForm.Controls.Add(txtFailed);
            resultForm.Controls.Add(btnCopyErrors);
            resultForm.Controls.Add(btnClose);

            resultForm.ShowDialog();
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
            var failedInvoices = new List<(string InvoiceNumber, string Error)>();

            foreach (var invoice in pendingInvoices)
            {
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
                        failedInvoices.Add((invoice.InvoiceNumber, response.ErrorMessage ?? "Unknown error"));
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

            if (failedInvoices.Count > 0)
            {
                ShowUploadResults(uploaded, failed, failedInvoices);
            }
            else
            {
                MessageBox.Show($"✅ All {uploaded} invoice(s) uploaded successfully!",
                    "Upload Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
                        _currentCompany.SellerAddress = qbCompanyInfo.Address;
                    if (string.IsNullOrEmpty(_currentCompany.SellerPhone))
                        _currentCompany.SellerPhone = qbCompanyInfo.Phone;
                    if (string.IsNullOrEmpty(_currentCompany.SellerEmail))
                        _currentCompany.SellerEmail = qbCompanyInfo.Email;
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

                var validationErrors = ValidateFBRInvoicePayload(details);
                var fbrPayload = _fbr.BuildFBRPayload(details);
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(fbrPayload, Newtonsoft.Json.Formatting.Indented);

                ShowInvoiceJsonForm(invoice, json, validationErrors);
            }
            catch (Exception ex)
            {
                ShowDetailedError(ex);
            }
        }

        private List<string> ValidateFBRInvoicePayload(FBRInvoicePayload details)
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

            foreach (var item in details.Items)
            {
                if (string.IsNullOrEmpty(item.HSCode))
                    errors.Add($"⚠️ HS Code missing for item: {item.ItemName}");
                if (item.RetailPrice == 0)
                    errors.Add($"⚠️ Retail Price missing for item: {item.ItemName}");
            }

            return errors;
        }

        private void ShowInvoiceJsonForm(Invoice invoice, string json, List<string> validationErrors)
        {
            var detailForm = new Form
            {
                Text = $"Invoice JSON - {invoice.InvoiceNumber}",
                Size = new Size(1000, 750),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimizeBox = false,
                MaximizeBox = true,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9F)
            };

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
                    Font = new Font("Segoe UI", 10F, FontStyle.Bold),
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
                    Font = new Font("Segoe UI", 9F)
                };

                warningPanel.Controls.Add(lblWarningTitle);
                warningPanel.Controls.Add(txtWarnings);
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

            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                BackColor = Color.FromArgb(248, 249, 250)
            };

            var btnCopyJson = CreateModernButton(
                "📋 Copy JSON",
                new Point(10, 8),
                new Size(120, 35),
                Color.FromArgb(33, 150, 243),
                Color.White
            );
            btnCopyJson.Click += (s, e) =>
            {
                Clipboard.SetText(json);
                MessageBox.Show("JSON copied to clipboard!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            var btnValidate = CreateModernButton(
                "✓ Validate",
                new Point(140, 8),
                new Size(100, 35),
                validationErrors.Count == 0 ? Color.FromArgb(40, 167, 69) : Color.FromArgb(255, 193, 7),
                validationErrors.Count == 0 ? Color.White : Color.FromArgb(33, 37, 41)
            );
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

            var btnClose = CreateModernButton(
                "Close",
                new Point(250, 8),
                new Size(100, 35),
                Color.FromArgb(108, 117, 125),
                Color.White
            );
            btnClose.Click += (s, e) => detailForm.Close();

            buttonPanel.Controls.Add(btnCopyJson);
            buttonPanel.Controls.Add(btnValidate);
            buttonPanel.Controls.Add(btnClose);

            detailForm.Controls.Add(txtJson);
            detailForm.Controls.Add(buttonPanel);

            detailForm.ShowDialog();
        }

        private void ShowDetailedError(Exception ex)
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
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9F)
            };

            var txtError = new TextBox
            {
                Text = errorDetails.ToString(),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = Color.FromArgb(254, 242, 242),
                ForeColor = Color.FromArgb(153, 27, 27),
                BorderStyle = BorderStyle.None
            };

            var btnCloseError = CreateModernButton(
                "Close",
                new Point(0, 0),
                new Size(100, 40),
                Color.FromArgb(220, 53, 69),
                Color.White
            );
            btnCloseError.Dock = DockStyle.Bottom;
            btnCloseError.Click += (s, e) => errorForm.Close();

            errorForm.Controls.Add(txtError);
            errorForm.Controls.Add(btnCloseError);
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