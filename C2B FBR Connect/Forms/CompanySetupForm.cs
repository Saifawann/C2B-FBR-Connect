using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace C2B_FBR_Connect.Forms
{
    public partial class CompanySetupForm : Form
    {
        private TextBox txtCompanyName;
        private TextBox txtFBRToken;
        private TextBox txtSellerNTN;
        private TextBox txtSellerAddress;
        private TextBox txtSellerProvince;
        private Button btnSave;
        private Button btnCancel;
        private Label lblCompanyName;
        private Label lblFBRToken;
        private Label lblSellerNTN;
        private Label lblSellerAddress;
        private Label lblSellerProvince;
        private Label lblInstructions;
        private GroupBox grpSellerInfo;

        public Company Company { get; private set; }

        public CompanySetupForm(string companyName, Company existingCompany = null)
        {
            InitializeComponent();
            SetupCustomUI();

            txtCompanyName.Text = companyName;
            txtCompanyName.ReadOnly = true;

            if (existingCompany != null)
            {
                txtFBRToken.Text = existingCompany.FBRToken ?? "";
                txtSellerNTN.Text = existingCompany.SellerNTN ?? "";

                // Auto-populate from QuickBooks company info (read-only for address)
                txtSellerAddress.Text = existingCompany.SellerAddress ?? "Fetched from QuickBooks";

                // Province is editable - populate from existing data (DO NOT override with QuickBooks data)
                txtSellerProvince.Text = existingCompany.SellerProvince ?? "";

                Company = existingCompany;
            }
            else
            {
                Company = new Company { CompanyName = companyName };
                txtSellerAddress.Text = "Will be fetched from QuickBooks";
                txtSellerProvince.Text = "";
            }
        }

        private void SetupCustomUI()
        {
            this.Text = "Company FBR Setup";
            this.Size = new Size(500, 420);
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

            // Seller Information Group Box
            grpSellerInfo = new GroupBox
            {
                Text = "Seller Information",
                Location = new Point(12, 125),
                Size = new Size(460, 200),
                Font = new Font("Arial", 9F, FontStyle.Bold)
            };

            lblSellerNTN = new Label
            {
                Text = "Seller NTN/CNIC:*",
                Location = new Point(12, 30),
                Size = new Size(120, 23),
                Font = new Font("Arial", 9F, FontStyle.Regular)
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
                Size = new Size(120, 23),
                Font = new Font("Arial", 9F, FontStyle.Regular)
            };

            txtSellerAddress = new TextBox
            {
                Location = new Point(128, 62),
                Size = new Size(315, 60),
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
                Location = new Point(12, 135),
                Size = new Size(120, 23),
                Font = new Font("Arial", 9F, FontStyle.Regular)
            };

            txtSellerProvince = new TextBox
            {
                Location = new Point(128, 132),
                Size = new Size(315, 23),
                Font = new Font("Arial", 9F),
                PlaceholderText = "Enter Province (e.g., Punjab, Sindh, KPK, Balochistan)",
                ReadOnly = false,  // Make it editable
                BackColor = Color.White,  // Use white background for editable field
                ForeColor = Color.Black  // Use normal text color
            };

            // Help label - updated text
            Label lblHelp = new Label
            {
                Text = "* Required fields. Address is auto-fetched from QuickBooks (read-only).",
                Location = new Point(12, 165),
                Size = new Size(430, 30),
                Font = new Font("Arial", 8F, FontStyle.Italic),
                ForeColor = Color.DarkBlue
            };

            // Add controls to group box
            grpSellerInfo.Controls.Add(lblSellerNTN);
            grpSellerInfo.Controls.Add(txtSellerNTN);
            grpSellerInfo.Controls.Add(lblSellerAddress);
            grpSellerInfo.Controls.Add(txtSellerAddress);
            grpSellerInfo.Controls.Add(lblSellerProvince);
            grpSellerInfo.Controls.Add(txtSellerProvince);
            grpSellerInfo.Controls.Add(lblHelp);

            btnSave = new Button
            {
                Text = "Save",
                Location = new Point(285, 345),
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
                Location = new Point(375, 345),
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
            if (string.IsNullOrWhiteSpace(txtSellerProvince.Text))
            {
                MessageBox.Show("Please enter Seller Province. This is required for FBR invoicing.",
                    "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSellerProvince.Focus();
                this.DialogResult = DialogResult.None;
                return;
            }

            // Save information to Company object
            Company.CompanyName = txtCompanyName.Text.Trim();
            Company.FBRToken = txtFBRToken.Text.Trim();
            Company.SellerNTN = txtSellerNTN.Text.Trim();
            Company.SellerProvince = txtSellerProvince.Text.Trim();

            // SellerAddress is auto-populated from QuickBooks
            Company.SellerAddress = txtSellerAddress.Text.Trim();

            Company.ModifiedDate = DateTime.Now;

            // The Company object is now updated and will be returned to the caller
            // The caller MUST save it using DatabaseService.SaveCompany(Company)
        }
    }
}