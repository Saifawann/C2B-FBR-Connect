using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
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

                // Auto-populate from QuickBooks company info (read-only for address)
                txtSellerAddress.Text = existingCompany.SellerAddress ?? "Fetched from QuickBooks";

                // Province will be set after loading from API
                txtSellerPhone.Text = existingCompany.SellerPhone ?? "";
                txtSellerEmail.Text = existingCompany.SellerEmail ?? "";

                Company = existingCompany;
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
            this.Size = new Size(500, 460);
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

            // ─── Seller Info Group ─────────────────────────────────────────────
            grpSellerInfo = new GroupBox
            {
                Text = "Seller Information",
                Location = new Point(12, 125),
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
                Location = new Point(285, 385),
                Size = new Size(85, 35),
                DialogResult = DialogResult.OK,
                BackColor = Color.FromArgb(46, 125, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 9F, FontStyle.Bold)
            };
            btnSave.Click += BtnSave_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(375, 385),
                Size = new Size(85, 35),
                DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(211, 47, 47),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 9F, FontStyle.Bold)
            };

            this.Controls.Add(lblInstructions);
            this.Controls.Add(lblCompanyName);
            this.Controls.Add(txtCompanyName);
            this.Controls.Add(lblFBRToken);
            this.Controls.Add(txtFBRToken);
            this.Controls.Add(grpSellerInfo);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
        }

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


            // SellerAddress is auto-populated from QuickBooks
            Company.SellerAddress = txtSellerAddress.Text.Trim();

            Company.ModifiedDate = DateTime.Now;

            // The Company object is now updated and will be returned to the caller
            // The caller MUST save it using DatabaseService.SaveCompany(Company)
        }
    }
}