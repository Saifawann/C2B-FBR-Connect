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
        private TextBox txtCompanyName;
        private TextBox txtFBRToken;
        private TextBox txtSellerNTN;
        private TextBox txtSellerAddress;
        private ComboBox cboSellerProvince;
        private TextBox txtSellerPhone;
        private TextBox txtSellerEmail;
        private Button btnSave;
        private Button btnCancel;
        private Label lblCompanyName;
        private Label lblFBRToken;
        private Label lblSellerNTN;
        private Label lblSellerAddress;
        private Label lblSellerProvince;
        private Label lblSellerPhone;
        private Label lblSellerEmail;
        private Label lblInstructions;
        private GroupBox grpSellerInfo;

        // Logo controls
        private GroupBox grpLogo;
        private PictureBox picLogo;
        private Button btnSelectLogo;
        private Button btnClearLogo;
        private Label lblLogoInfo;

        private Dictionary<string, int> _provinceCodeMap = new Dictionary<string, int>();
        private FBRApiService _fbrApi;
        public Company Company { get; private set; }

        public CompanySetupForm(string companyName, Company existingCompany = null)
        {
            InitializeComponent();
            SetupCustomUI();

            _fbrApi = new FBRApiService();

            txtCompanyName.Text = companyName;
            txtCompanyName.ReadOnly = true;

            if (existingCompany != null)
            {
                txtFBRToken.Text = existingCompany.FBRToken ?? "";
                txtSellerNTN.Text = existingCompany.SellerNTN ?? "";
                txtSellerAddress.Text = existingCompany.SellerAddress ?? "Fetched from QuickBooks";
                txtSellerPhone.Text = existingCompany.SellerPhone ?? "";
                txtSellerEmail.Text = existingCompany.SellerEmail ?? "";

                Company = existingCompany;

                // Load existing logo if available
                LoadExistingLogo();
            }
            else
            {
                Company = new Company { CompanyName = companyName };
                txtSellerAddress.Text = "Will be fetched from QuickBooks";
            }

            // Load provinces from API
            LoadProvincesAsync(existingCompany?.SellerProvince);
        }

        private void SetupCustomUI()
        {
            this.Text = "Company FBR Setup";
            this.Size = new Size(500, 600);  
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            lblInstructions = new Label
            {
                Text = "Configure FBR Digital Invoicing settings for this QuickBooks company:",
                Location = new Point(12, 12),
                Size = new Size(460, 30),
                Font = new Font("Arial", 9F)
            };

            lblCompanyName = new Label
            {
                Text = "Company Name:",
                Location = new Point(12, 55),
                Size = new Size(120, 23)
            };

            txtCompanyName = new TextBox
            {
                Location = new Point(140, 52),
                Size = new Size(320, 23),
                BackColor = Color.FromArgb(240, 240, 240)
            };

            lblFBRToken = new Label
            {
                Text = "FBR Token:",
                Location = new Point(12, 90),
                Size = new Size(120, 23)
            };

            txtFBRToken = new TextBox
            {
                Location = new Point(140, 87),
                Size = new Size(320, 23),
                UseSystemPasswordChar = false
            };

            // ─── Company Logo Group ─────────────────────────────────────────────
            grpLogo = new GroupBox
            {
                Text = "Company Logo",
                Location = new Point(12, 125),
                Size = new Size(460, 120),
                Font = new Font("Arial", 9F, FontStyle.Bold)
            };

            picLogo = new PictureBox
            {
                Location = new Point(12, 25),
                Size = new Size(150, 80),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White
            };

            btnSelectLogo = new Button
            {
                Text = "Select Logo",
                Location = new Point(175, 30),
                Size = new Size(100, 32),
                BackColor = Color.FromArgb(33, 150, 243),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 9F),
                Cursor = Cursors.Hand
            };
            btnSelectLogo.FlatAppearance.BorderSize = 0;
            btnSelectLogo.Click += BtnSelectLogo_Click;

            btnClearLogo = new Button
            {
                Text = "Clear Logo",
                Location = new Point(285, 30),
                Size = new Size(100, 32),
                BackColor = Color.FromArgb(244, 67, 54),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 9F),
                Enabled = false,
                Cursor = Cursors.Hand
            };
            btnClearLogo.FlatAppearance.BorderSize = 0;
            btnClearLogo.Click += BtnClearLogo_Click;

            lblLogoInfo = new Label
            {
                Text = "No logo selected",
                Location = new Point(175, 70),
                Size = new Size(270, 35),
                Font = new Font("Arial", 8F, FontStyle.Italic),
                ForeColor = Color.Gray
            };

            grpLogo.Controls.Add(picLogo);
            grpLogo.Controls.Add(btnSelectLogo);
            grpLogo.Controls.Add(btnClearLogo);
            grpLogo.Controls.Add(lblLogoInfo);

            // ─── Seller Info Group ─────────────────────────────────────────────
            grpSellerInfo = new GroupBox
            {
                Text = "Seller Information",
                Location = new Point(12, 255),  // Adjusted position
                Size = new Size(460, 250),
                Font = new Font("Arial", 9F, FontStyle.Bold)
            };

            lblSellerNTN = new Label
            {
                Text = "Seller NTN/CNIC:*",
                Location = new Point(12, 30),
                Size = new Size(120, 23)
            };

            txtSellerNTN = new TextBox
            {
                Location = new Point(128, 27),
                Size = new Size(315, 23),
                Font = new Font("Arial", 9F),
                PlaceholderText = "Enter your NTN or CNIC number"
            };

            lblSellerAddress = new Label
            {
                Text = "Seller Address:",
                Location = new Point(12, 65),
                Size = new Size(120, 23)
            };

            txtSellerAddress = new TextBox
            {
                Location = new Point(128, 62),
                Size = new Size(315, 50),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Arial", 9F),
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.DarkGray
            };

            lblSellerProvince = new Label
            {
                Text = "Seller Province:*",
                Location = new Point(12, 120),
                Size = new Size(120, 23)
            };

            cboSellerProvince = new ComboBox
            {
                Location = new Point(128, 117),
                Size = new Size(315, 23),
                Font = new Font("Arial", 9F),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.White,
                ForeColor = Color.Black
            };

            lblSellerPhone = new Label
            {
                Text = "Seller Phone:",
                Location = new Point(12, 155),
                Size = new Size(120, 23)
            };

            txtSellerPhone = new TextBox
            {
                Location = new Point(128, 152),
                Size = new Size(315, 23),
                Font = new Font("Arial", 9F),
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.DarkGray
            };

            lblSellerEmail = new Label
            {
                Text = "Seller Email:",
                Location = new Point(12, 185),
                Size = new Size(120, 23)
            };

            txtSellerEmail = new TextBox
            {
                Location = new Point(128, 182),
                Size = new Size(315, 23),
                Font = new Font("Arial", 9F),
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.DarkGray
            };

            Label lblHelp = new Label
            {
                Text = "* Required fields. Address, Phone, and Email are auto-fetched from QuickBooks.",
                Location = new Point(12, 215),
                Size = new Size(430, 25),
                Font = new Font("Arial", 8F, FontStyle.Italic),
                ForeColor = Color.DarkBlue
            };

            grpSellerInfo.Controls.Add(lblSellerNTN);
            grpSellerInfo.Controls.Add(txtSellerNTN);
            grpSellerInfo.Controls.Add(lblSellerAddress);
            grpSellerInfo.Controls.Add(txtSellerAddress);
            grpSellerInfo.Controls.Add(lblSellerProvince);
            grpSellerInfo.Controls.Add(cboSellerProvince);
            grpSellerInfo.Controls.Add(lblSellerPhone);
            grpSellerInfo.Controls.Add(txtSellerPhone);
            grpSellerInfo.Controls.Add(lblSellerEmail);
            grpSellerInfo.Controls.Add(txtSellerEmail);
            grpSellerInfo.Controls.Add(lblHelp);

            // ─── Buttons ─────────────────────────────────────────────
            btnSave = new Button
            {
                Text = "Save",
                Location = new Point(285, 515),  // Adjusted position
                Size = new Size(85, 35),
                DialogResult = DialogResult.OK,
                BackColor = Color.FromArgb(46, 125, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(375, 515),  // Adjusted position
                Size = new Size(85, 35),
                DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(211, 47, 47),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;

            this.Controls.Add(lblInstructions);
            this.Controls.Add(lblCompanyName);
            this.Controls.Add(txtCompanyName);
            this.Controls.Add(lblFBRToken);
            this.Controls.Add(txtFBRToken);
            this.Controls.Add(grpLogo);
            this.Controls.Add(grpSellerInfo);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
        }

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
                        // Validate image file
                        if (!ImageHelper.IsValidImageFile(openFileDialog.FileName))
                        {
                            MessageBox.Show("The selected file is not a valid image format.",
                                "Invalid Image",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                            return;
                        }

                        // Get image dimensions and file size for info
                        var dimensions = ImageHelper.GetImageDimensions(openFileDialog.FileName);
                        var fileSize = ImageHelper.GetFileSizeFormatted(openFileDialog.FileName);

                        // Load, resize, and convert to bytes
                        Company.LogoImage = ImageHelper.LoadAndResizeImage(openFileDialog.FileName, 500, 500);

                        // Display in PictureBox
                        DisplayLogo(Company.LogoImage);

                        // Update info label
                        lblLogoInfo.Text = $"Logo loaded successfully\nOriginal: {dimensions.width}x{dimensions.height} ({fileSize})";
                        lblLogoInfo.ForeColor = Color.Green;

                        // Enable clear button
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

                MessageBox.Show("Logo removed successfully.",
                    "Logo Cleared",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void LoadExistingLogo()
        {
            if (Company?.LogoImage != null && Company.LogoImage.Length > 0)
            {
                try
                {
                    DisplayLogo(Company.LogoImage);

                    // Calculate approximate size
                    double sizeKB = Company.LogoImage.Length / 1024.0;
                    string sizeText = sizeKB < 1024
                        ? $"{sizeKB:0.##} KB"
                        : $"{sizeKB / 1024:0.##} MB";

                    lblLogoInfo.Text = $"Current logo loaded\nSize: {sizeText}";
                    lblLogoInfo.ForeColor = Color.Green;
                    btnClearLogo.Enabled = true;
                }
                catch (Exception ex)
                {
                    lblLogoInfo.Text = "Error loading saved logo";
                    lblLogoInfo.ForeColor = Color.Red;
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
            // Dispose previous image to free memory
            if (picLogo.Image != null)
            {
                picLogo.Image.Dispose();
                picLogo.Image = null;
            }

            // Load new image
            picLogo.Image = ImageHelper.ByteArrayToImage(logoBytes);
        }

        private void ClearLogoDisplay()
        {
            if (picLogo.Image != null)
            {
                picLogo.Image.Dispose();
                picLogo.Image = null;
            }

            lblLogoInfo.Text = "No logo selected";
            lblLogoInfo.ForeColor = Color.Gray;
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
                // Get token if available
                string token = !string.IsNullOrWhiteSpace(txtFBRToken.Text) ? txtFBRToken.Text : null;

                // Fetch provinces from FBR API using FBRApiService
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

                    // Select the existing province if available
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
                MessageBox.Show("Please enter FBR token.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtFBRToken.Focus();
                this.DialogResult = DialogResult.None;
                return;
            }

            // Validate Seller NTN
            if (string.IsNullOrWhiteSpace(txtSellerNTN.Text))
            {
                MessageBox.Show("Please enter Seller NTN/CNIC. This is required for FBR invoicing.",
                    "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSellerNTN.Focus();
                this.DialogResult = DialogResult.None;
                return;
            }

            // Validate Seller Province
            if (cboSellerProvince.SelectedItem == null)
            {
                MessageBox.Show("Please select Seller Province. This is required for FBR invoicing.",
                    "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboSellerProvince.Focus();
                this.DialogResult = DialogResult.None;
                return;
            }

            // Save information to Company object
            Company.CompanyName = txtCompanyName.Text.Trim();
            Company.FBRToken = txtFBRToken.Text.Trim();
            Company.SellerNTN = txtSellerNTN.Text.Trim();
            Company.SellerProvince = cboSellerProvince.SelectedItem.ToString();
            Company.SellerPhone = txtSellerPhone.Text?.Trim();
            Company.SellerEmail = txtSellerEmail.Text?.Trim();
            Company.SellerAddress = txtSellerAddress.Text.Trim();

            // Logo is already set in Company.LogoImage via BtnSelectLogo_Click
            // or cleared via BtnClearLogo_Click

            Company.ModifiedDate = DateTime.Now;

            // The Company object is now updated and will be returned to the caller
            // The caller MUST save it using DatabaseService.SaveCompany(Company)
        }

        #endregion

        #region Form Cleanup

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose image to free memory
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