using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using C2B_FBR_Connect.Utilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C2B_FBR_Connect.Forms
{
    public partial class CompanySetupForm : Form
    {
        #region Controls

        // Header
        private Panel pnlHeader;
        private Label lblTitle;
        private Label lblSubtitle;

        // FBR Configuration
        private GroupBox grpFBRConfig;
        private Label lblCompanyName, lblFBRToken, lblEnvironment;
        private TextBox txtCompanyName, txtFBRToken;
        private ComboBox cboEnvironment;
        private Panel pnlEnvironmentIndicator;

        // Company Logo
        private GroupBox grpLogo;
        private PictureBox picLogo;
        private Button btnSelectLogo, btnClearLogo;
        private Label lblLogoInfo;

        // Seller Information
        private GroupBox grpSellerInfo;
        private Label lblSellerNTN, lblStrNo, lblSellerProvince, lblSellerAddress, lblSellerPhone, lblSellerEmail;
        private TextBox txtSellerNTN, txtStrNo, txtSellerAddress, txtSellerPhone, txtSellerEmail;
        private ComboBox cboSellerProvince;
        private Button btnRefreshFromQB;
        private Label lblLastRefresh;

        // Footer
        private Panel pnlFooter;
        private Button btnSave, btnCancel;
        private Label lblRequiredNote;

        #endregion

        #region Fields

        private Dictionary<string, int> _provinceCodeMap = new Dictionary<string, int>();
        private FBRApiService _fbrApi;
        private QuickBooksService _qbService;
        public Company Company { get; private set; }

        // Colors
        private readonly Color PrimaryColor = Color.FromArgb(45, 62, 80);
        private readonly Color SecondaryColor = Color.FromArgb(52, 152, 219);
        private readonly Color SuccessColor = Color.FromArgb(46, 204, 113);
        private readonly Color WarningColor = Color.FromArgb(241, 196, 15);
        private readonly Color DangerColor = Color.FromArgb(231, 76, 60);
        private readonly Color LightBgColor = Color.FromArgb(248, 249, 250);
        private readonly Color BorderColor = Color.FromArgb(222, 226, 230);
        private readonly Color TextMutedColor = Color.FromArgb(108, 117, 125);

        #endregion

        #region Constructor

        public CompanySetupForm(string companyName, Company existingCompany = null, QuickBooksService qbService = null)
        {
            InitializeComponent();
            SetupCustomUI();

            _fbrApi = new FBRApiService();
            _qbService = qbService;

            txtCompanyName.Text = companyName;

            if (existingCompany != null)
            {
                txtFBRToken.Text = existingCompany.FBRToken ?? "";

                // Set NTN with proper color (not placeholder)
                if (!string.IsNullOrWhiteSpace(existingCompany.SellerNTN))
                {
                    txtSellerNTN.Text = existingCompany.SellerNTN;
                    txtSellerNTN.ForeColor = Color.Black;
                }

                // Set STR with proper color (not placeholder)
                if (!string.IsNullOrWhiteSpace(existingCompany.StrNo))
                {
                    txtStrNo.Text = existingCompany.StrNo;
                    txtStrNo.ForeColor = Color.Black;
                }

                txtSellerAddress.Text = existingCompany.SellerAddress ?? "";
                txtSellerPhone.Text = existingCompany.SellerPhone ?? "";
                txtSellerEmail.Text = existingCompany.SellerEmail ?? "";

                if (!string.IsNullOrEmpty(existingCompany.Environment))
                {
                    int envIndex = cboEnvironment.FindStringExact(existingCompany.Environment);
                    if (envIndex >= 0)
                        cboEnvironment.SelectedIndex = envIndex;
                }

                Company = existingCompany;
                LoadExistingLogo();
            }
            else
            {
                Company = new Company { CompanyName = companyName };
            }

            LoadProvincesAsync(existingCompany?.SellerProvince);
            UpdateEnvironmentIndicator();
            UpdateRefreshButtonState();
        }

        #endregion

        #region UI Setup

        private void SetupCustomUI()
        {
            // Form Configuration
            this.Text = "Company Setup";
            this.Size = new Size(520, 785);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = LightBgColor;
            this.Font = new Font("Segoe UI", 9F);

            CreateHeader();
            CreateFBRConfigGroup();
            CreateLogoGroup();
            CreateSellerInfoGroup();
            CreateFooter();

            // Add controls to form
            this.Controls.Add(pnlFooter);
            this.Controls.Add(grpSellerInfo);
            this.Controls.Add(grpLogo);
            this.Controls.Add(grpFBRConfig);
            this.Controls.Add(pnlHeader);
        }

        private void CreateHeader()
        {
            pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 70,
                BackColor = PrimaryColor,
                Padding = new Padding(20, 15, 20, 15)
            };

            lblTitle = new Label
            {
                Text = "⚙ Company Configuration",
                Font = new Font("Segoe UI Semibold", 14F),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 12)
            };

            lblSubtitle = new Label
            {
                Text = "Configure FBR Digital Invoicing settings",
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(200, 200, 200),
                AutoSize = true,
                Location = new Point(20, 40)
            };

            pnlHeader.Controls.Add(lblTitle);
            pnlHeader.Controls.Add(lblSubtitle);
        }

        private void CreateFBRConfigGroup()
        {
            grpFBRConfig = new GroupBox
            {
                Text = "  FBR Configuration  ",
                Location = new Point(15, 85),
                Size = new Size(475, 145),
                Font = new Font("Segoe UI Semibold", 9F),
                ForeColor = PrimaryColor,
                BackColor = Color.White,
                Padding = new Padding(15, 10, 15, 10)
            };

            // Company Name
            lblCompanyName = CreateLabel("Company Name:", 15, 28);
            txtCompanyName = CreateTextBox(130, 25, 325, true);
            txtCompanyName.BackColor = Color.FromArgb(245, 245, 245);

            // FBR Token
            lblFBRToken = CreateLabel("FBR Token: *", 15, 60);
            txtFBRToken = CreateTextBox(130, 57, 325);
            txtFBRToken.UseSystemPasswordChar = false;

            // Environment
            lblEnvironment = CreateLabel("Environment: *", 15, 92);

            cboEnvironment = new ComboBox
            {
                Location = new Point(130, 89),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.5F),
                FlatStyle = FlatStyle.Flat
            };
            cboEnvironment.Items.AddRange(new object[] { "Sandbox", "Production" });
            cboEnvironment.SelectedIndex = 0;
            cboEnvironment.SelectedIndexChanged += (s, e) => UpdateEnvironmentIndicator();

            // Environment Indicator
            pnlEnvironmentIndicator = new Panel
            {
                Location = new Point(340, 89),
                Size = new Size(115, 25),
                BackColor = WarningColor
            };

            var lblEnvIndicator = new Label
            {
                Name = "lblEnvIndicator",
                Text = "⚠ SANDBOX",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleCenter
            };
            pnlEnvironmentIndicator.Controls.Add(lblEnvIndicator);

            grpFBRConfig.Controls.AddRange(new Control[] {
                lblCompanyName, txtCompanyName,
                lblFBRToken, txtFBRToken,
                lblEnvironment, cboEnvironment, pnlEnvironmentIndicator
            });
        }

        private void CreateLogoGroup()
        {
            grpLogo = new GroupBox
            {
                Text = "  Company Logo  ",
                Location = new Point(15, 240),
                Size = new Size(475, 110),
                Font = new Font("Segoe UI Semibold", 9F),
                ForeColor = PrimaryColor,
                BackColor = Color.White
            };

            picLogo = new PictureBox
            {
                Location = new Point(15, 25),
                Size = new Size(120, 70),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = LightBgColor
            };

            btnSelectLogo = CreateButton("📁 Select Logo", 150, 30, 110, 32, SecondaryColor);
            btnSelectLogo.Click += BtnSelectLogo_Click;

            btnClearLogo = CreateButton("🗑 Clear", 270, 30, 85, 32, DangerColor);
            btnClearLogo.Enabled = false;
            btnClearLogo.Click += BtnClearLogo_Click;

            lblLogoInfo = new Label
            {
                Text = "No logo selected\nRecommended: 500x500px or smaller",
                Location = new Point(150, 68),
                Size = new Size(310, 35),
                Font = new Font("Segoe UI", 8F),
                ForeColor = TextMutedColor
            };

            grpLogo.Controls.AddRange(new Control[] { picLogo, btnSelectLogo, btnClearLogo, lblLogoInfo });
        }

        private void CreateSellerInfoGroup()
        {
            grpSellerInfo = new GroupBox
            {
                Text = "  Seller Information  ",
                Location = new Point(15, 360),
                Size = new Size(475, 310),
                Font = new Font("Segoe UI Semibold", 9F),
                ForeColor = PrimaryColor,
                BackColor = Color.White
            };

            int labelX = 15;
            int inputX = 130;
            int inputWidth = 325;
            int rowHeight = 35;
            int startY = 28;

            // Seller NTN
            lblSellerNTN = CreateLabel("NTN / CNIC: *", labelX, startY);
            txtSellerNTN = CreateTextBox(inputX, startY - 3, inputWidth);
            AddPlaceholder(txtSellerNTN, "Enter your NTN or CNIC number");

            // STR No
            lblStrNo = CreateLabel("STR No:", labelX, startY + rowHeight);
            txtStrNo = CreateTextBox(inputX, startY + rowHeight - 3, inputWidth);
            AddPlaceholder(txtStrNo, "Enter STR number");

            // Seller Province
            lblSellerProvince = CreateLabel("Province: *", labelX, startY + rowHeight * 2);
            cboSellerProvince = new ComboBox
            {
                Location = new Point(inputX, startY + rowHeight * 2 - 3),
                Size = new Size(inputWidth, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.5F),
                FlatStyle = FlatStyle.Flat
            };

            // Seller Address
            lblSellerAddress = CreateLabel("Address:", labelX, startY + rowHeight * 3);
            txtSellerAddress = new TextBox
            {
                Location = new Point(inputX, startY + rowHeight * 3 - 3),
                Size = new Size(inputWidth, 50),
                Font = new Font("Segoe UI", 9.5F),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BorderStyle = BorderStyle.FixedSingle,
                ForeColor = Color.Black
            };

            // Seller Phone
            lblSellerPhone = CreateLabel("Phone:", labelX, startY + rowHeight * 3 + 55);
            txtSellerPhone = CreateTextBox(inputX, startY + rowHeight * 3 + 52, inputWidth);

            // Seller Email
            lblSellerEmail = CreateLabel("Email:", labelX, startY + rowHeight * 3 + 90);
            txtSellerEmail = CreateTextBox(inputX, startY + rowHeight * 3 + 87, inputWidth);

            // Refresh from QuickBooks button
            btnRefreshFromQB = CreateButton("🔄 Refresh from QuickBooks", labelX, 253, 200, 32, SecondaryColor);
            btnRefreshFromQB.Click += BtnRefreshFromQB_Click;

            // Last refresh label
            lblLastRefresh = new Label
            {
                Text = "",
                Location = new Point(220, 260),
                Size = new Size(240, 20),
                Font = new Font("Segoe UI", 8F, FontStyle.Italic),
                ForeColor = TextMutedColor,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Help text
            var lblHelp = new Label
            {
                Text = "ℹ Click 'Refresh' to load Address, Phone, and Email from QuickBooks",
                Location = new Point(labelX, 288),
                Size = new Size(440, 20),
                Font = new Font("Segoe UI", 8F, FontStyle.Italic),
                ForeColor = SecondaryColor
            };

            grpSellerInfo.Controls.AddRange(new Control[] {
                lblSellerNTN, txtSellerNTN,
                lblStrNo, txtStrNo,
                lblSellerProvince, cboSellerProvince,
                lblSellerAddress, txtSellerAddress,
                lblSellerPhone, txtSellerPhone,
                lblSellerEmail, txtSellerEmail,
                btnRefreshFromQB, lblLastRefresh,
                lblHelp
            });
        }

        private void CreateFooter()
        {
            pnlFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 65,
                BackColor = Color.White,
                Padding = new Padding(15, 12, 15, 12)
            };

            // Top border
            pnlFooter.Paint += (s, e) =>
            {
                using (var pen = new Pen(BorderColor, 1))
                {
                    e.Graphics.DrawLine(pen, 0, 0, pnlFooter.Width, 0);
                }
            };

            lblRequiredNote = new Label
            {
                Text = "* Required fields",
                Location = new Point(15, 22),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = TextMutedColor
            };

            btnCancel = CreateButton("Cancel", 290, 12, 95, 40, Color.FromArgb(108, 117, 125));
            btnCancel.DialogResult = DialogResult.Cancel;

            btnSave = CreateButton("💾 Save", 395, 12, 95, 40, SuccessColor);
            btnSave.Click += BtnSave_Click;

            pnlFooter.Controls.AddRange(new Control[] { lblRequiredNote, btnCancel, btnSave });
        }

        #endregion

        #region Helper Methods

        private Label CreateLabel(string text, int x, int y)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(73, 80, 87)
            };
        }

        private TextBox CreateTextBox(int x, int y, int width, bool readOnly = false)
        {
            var textBox = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(width, 25),
                Font = new Font("Segoe UI", 9.5F),
                BorderStyle = BorderStyle.FixedSingle,
                ReadOnly = readOnly,
                ForeColor = Color.Black
            };

            if (readOnly)
            {
                textBox.BackColor = Color.FromArgb(245, 245, 245);
            }

            return textBox;
        }

        private Button CreateButton(string text, int x, int y, int width, int height, Color backColor)
        {
            var button = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, height),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F),
                Cursor = Cursors.Hand
            };

            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = ControlPaint.Dark(backColor, 0.1f);
            button.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(backColor, 0.2f);

            return button;
        }

        private void AddPlaceholder(TextBox textBox, string placeholder)
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

        private void UpdateEnvironmentIndicator()
        {
            var env = cboEnvironment.SelectedItem?.ToString() ?? "Sandbox";
            var lblIndicator = pnlEnvironmentIndicator.Controls.Find("lblEnvIndicator", false).FirstOrDefault() as Label;

            if (env == "Production")
            {
                pnlEnvironmentIndicator.BackColor = SuccessColor;
                if (lblIndicator != null) lblIndicator.Text = "✓ PRODUCTION";
            }
            else
            {
                pnlEnvironmentIndicator.BackColor = WarningColor;
                if (lblIndicator != null) lblIndicator.Text = "⚠ SANDBOX";
            }
        }

        private void UpdateRefreshButtonState()
        {
            btnRefreshFromQB.Enabled = _qbService != null;

            if (_qbService == null)
            {
                btnRefreshFromQB.BackColor = Color.FromArgb(180, 180, 180);
                lblLastRefresh.Text = "QuickBooks service not available";
            }
        }

        #endregion

        #region QuickBooks Refresh

        private void BtnRefreshFromQB_Click(object sender, EventArgs e)
        {
            if (_qbService == null)
            {
                MessageBox.Show(
                    "QuickBooks service is not available.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // Disable button and show loading state
            btnRefreshFromQB.Enabled = false;
            var originalText = btnRefreshFromQB.Text;
            btnRefreshFromQB.Text = "⏳ Refreshing...";
            Application.DoEvents();

            try
            {
                // Get fresh company info from QuickBooks (forces re-fetch)
                var companyInfo = _qbService.RefreshCompanyInfo();

                if (companyInfo != null)
                {
                    // Update form fields only
                    if (!string.IsNullOrWhiteSpace(companyInfo.Name))
                    {
                        txtCompanyName.Text = companyInfo.Name;
                    }

                    if (!string.IsNullOrWhiteSpace(companyInfo.Address))
                    {
                        txtSellerAddress.Text = companyInfo.Address;
                    }

                    if (!string.IsNullOrWhiteSpace(companyInfo.Phone))
                    {
                        txtSellerPhone.Text = companyInfo.Phone;
                    }

                    if (!string.IsNullOrWhiteSpace(companyInfo.Email))
                    {
                        txtSellerEmail.Text = companyInfo.Email;
                    }

                    lblLastRefresh.Text = $"✓ Refreshed: {DateTime.Now:HH:mm:ss}";
                    lblLastRefresh.ForeColor = SuccessColor;
                }
                else
                {
                    lblLastRefresh.Text = "✗ No data returned";
                    lblLastRefresh.ForeColor = DangerColor;

                    MessageBox.Show(
                        "No company data returned from QuickBooks.",
                        "Refresh Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                lblLastRefresh.Text = "✗ Refresh failed";
                lblLastRefresh.ForeColor = DangerColor;

                MessageBox.Show(
                    $"Failed to refresh data:\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                btnRefreshFromQB.Text = originalText;
                btnRefreshFromQB.Enabled = true;
            }
        }

        #endregion

        #region Logo Management

        private void BtnSelectLogo_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp;*.gif)|*.jpg;*.jpeg;*.png;*.bmp;*.gif";
                openFileDialog.Title = "Select Company Logo";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (!ImageHelper.IsValidImageFile(openFileDialog.FileName))
                        {
                            MessageBox.Show("The selected file is not a valid image format.",
                                "Invalid Image",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                            return;
                        }

                        var dimensions = ImageHelper.GetImageDimensions(openFileDialog.FileName);
                        var fileSize = ImageHelper.GetFileSizeFormatted(openFileDialog.FileName);

                        Company.LogoImage = ImageHelper.LoadAndResizeImage(openFileDialog.FileName, 500, 500);

                        DisplayLogo(Company.LogoImage);

                        lblLogoInfo.Text = $"✓ Logo loaded successfully\nOriginal: {dimensions.width}x{dimensions.height} ({fileSize})";
                        lblLogoInfo.ForeColor = SuccessColor;

                        btnClearLogo.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error loading image:\n{ex.Message}",
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                        ClearLogoDisplay();
                    }
                }
            }
        }

        private void BtnClearLogo_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Are you sure you want to remove the company logo?",
                "Confirm Clear Logo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                ClearLogoDisplay();
                Company.LogoImage = null;
            }
        }

        private void LoadExistingLogo()
        {
            if (Company?.LogoImage != null && Company.LogoImage.Length > 0)
            {
                try
                {
                    DisplayLogo(Company.LogoImage);

                    double sizeKB = Company.LogoImage.Length / 1024.0;
                    string sizeText = sizeKB < 1024
                        ? $"{sizeKB:0.##} KB"
                        : $"{sizeKB / 1024:0.##} MB";

                    lblLogoInfo.Text = $"✓ Current logo loaded\nSize: {sizeText}";
                    lblLogoInfo.ForeColor = SuccessColor;
                    btnClearLogo.Enabled = true;
                }
                catch (Exception ex)
                {
                    lblLogoInfo.Text = "✗ Error loading saved logo";
                    lblLogoInfo.ForeColor = DangerColor;
                    btnClearLogo.Enabled = false;

                    System.Diagnostics.Debug.WriteLine($"Error loading logo: {ex.Message}");
                }
            }
            else
            {
                ClearLogoDisplay();
            }
        }

        private void DisplayLogo(byte[] logoBytes)
        {
            if (picLogo.Image != null)
            {
                picLogo.Image.Dispose();
                picLogo.Image = null;
            }

            picLogo.Image = ImageHelper.ByteArrayToImage(logoBytes);
        }

        private void ClearLogoDisplay()
        {
            if (picLogo.Image != null)
            {
                picLogo.Image.Dispose();
                picLogo.Image = null;
            }

            lblLogoInfo.Text = "No logo selected\nRecommended: 500x500px or smaller";
            lblLogoInfo.ForeColor = TextMutedColor;
            btnClearLogo.Enabled = false;
        }

        #endregion

        #region Province Management

        private async void LoadProvincesAsync(string selectedProvince = null)
        {
            cboSellerProvince.Items.Clear();
            cboSellerProvince.Items.Add("Loading provinces...");
            cboSellerProvince.SelectedIndex = 0;
            cboSellerProvince.Enabled = false;

            try
            {
                string token = !string.IsNullOrWhiteSpace(txtFBRToken.Text) ? txtFBRToken.Text : null;

                var provinces = await _fbrApi.FetchProvincesAsync(token);

                cboSellerProvince.Items.Clear();
                _provinceCodeMap.Clear();

                if (provinces != null && provinces.Count > 0)
                {
                    foreach (var province in provinces)
                    {
                        cboSellerProvince.Items.Add(province.StateProvinceDesc);
                        _provinceCodeMap[province.StateProvinceDesc] = province.StateProvinceCode;
                    }

                    if (!string.IsNullOrWhiteSpace(selectedProvince))
                    {
                        int index = cboSellerProvince.FindStringExact(selectedProvince);
                        if (index >= 0)
                        {
                            cboSellerProvince.SelectedIndex = index;
                        }
                    }
                }
                else
                {
                    AddFallbackProvinces();
                }
            }
            catch (Exception)
            {
                AddFallbackProvinces();
            }
            finally
            {
                cboSellerProvince.Enabled = true;
            }
        }

        private void AddFallbackProvinces()
        {
            cboSellerProvince.Items.Clear();
            _provinceCodeMap.Clear();

            var fallbackProvinces = new Dictionary<string, int>
            {
                { "BALOCHISTAN", 2 },
                { "AZAD JAMMU AND KASHMIR", 4 },
                { "CAPITAL TERRITORY", 5 },
                { "KHYBER PAKHTUNKHWA", 6 },
                { "PUNJAB", 7 },
                { "SINDH", 8 },
                { "GILGIT BALTISTAN", 9 }
            };

            foreach (var province in fallbackProvinces)
            {
                cboSellerProvince.Items.Add(province.Key);
                _provinceCodeMap[province.Key] = province.Value;
            }
        }

        public int GetSelectedProvinceCode()
        {
            if (cboSellerProvince.SelectedItem != null)
            {
                string selectedProvince = cboSellerProvince.SelectedItem.ToString();
                if (_provinceCodeMap.ContainsKey(selectedProvince))
                {
                    return _provinceCodeMap[selectedProvince];
                }
            }
            return 0;
        }

        #endregion

        #region Save & Validation

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // Validate FBR Token
            if (string.IsNullOrWhiteSpace(txtFBRToken.Text))
            {
                ShowValidationError("Please enter FBR Token.", txtFBRToken);
                return;
            }

            // Validate Seller NTN
            string ntnText = txtSellerNTN.Text;
            if (ntnText == "Enter your NTN or CNIC number") ntnText = "";

            if (string.IsNullOrWhiteSpace(ntnText))
            {
                ShowValidationError("Please enter Seller NTN/CNIC. This is required for FBR invoicing.", txtSellerNTN);
                return;
            }

            // Get SRO No (handle placeholder)
            string StrNoText = txtStrNo.Text;
            if (StrNoText == "Enter STR number (if applicable)") StrNoText = "";

            // Validate Environment
            if (cboEnvironment.SelectedItem == null)
            {
                ShowValidationError("Please select an environment (Production or Sandbox).", cboEnvironment);
                return;
            }

            // Validate Seller Province
            if (cboSellerProvince.SelectedItem == null ||
                cboSellerProvince.SelectedItem.ToString() == "Loading provinces...")
            {
                ShowValidationError("Please select Seller Province. This is required for FBR invoicing.", cboSellerProvince);
                return;
            }

            // Confirm Production environment
            if (cboEnvironment.SelectedItem.ToString() == "Production")
            {
                var result = MessageBox.Show(
                    "You have selected PRODUCTION environment.\n\n" +
                    "All invoices will be submitted to FBR's live system.\n\n" +
                    "Are you sure you want to continue?",
                    "Confirm Production Environment",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result != DialogResult.Yes)
                {
                    cboEnvironment.Focus();
                    return;
                }
            }

            // Save information to Company object
            Company.CompanyName = txtCompanyName.Text.Trim();
            Company.Environment = cboEnvironment.SelectedItem.ToString();
            Company.FBRToken = txtFBRToken.Text.Trim();
            Company.SellerNTN = ntnText.Trim();
            Company.StrNo = !string.IsNullOrWhiteSpace(StrNoText) ? StrNoText.Trim() : null;
            Company.SellerProvince = cboSellerProvince.SelectedItem.ToString();
            Company.SellerPhone = txtSellerPhone.Text?.Trim();
            Company.SellerEmail = txtSellerEmail.Text?.Trim();
            Company.SellerAddress = txtSellerAddress.Text?.Trim();
            Company.ModifiedDate = DateTime.Now;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void ShowValidationError(string message, Control focusControl)
        {
            MessageBox.Show(message, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            focusControl?.Focus();
        }

        #endregion

        #region Form Cleanup

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (picLogo?.Image != null)
                {
                    picLogo.Image.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}