using System;
using System.Windows.Forms;
using System.Drawing;

namespace C2B_FBR_Connect.Forms
{
    public partial class ActivationForm : Form
    {
        private TextBox txtHardwareId;
        private TextBox txtLicenseKey;
        private Button btnActivate;
        private Button btnCopyHardwareId;
        private Button btnExit;
        private Label lblTitle;
        private Label lblHardwareInfo;
        private Label lblLicenseInfo;
        private Label lblStatus;
        private PictureBox picLock;

        public string LicenseKey { get; private set; }
        public bool IsActivated { get; private set; }

        private readonly LicenseSystem.LicenseValidator _validator;

        public ActivationForm()
        {
            InitializeComponent();

            // Load public key from file
            string publicKeyPath = System.IO.Path.Combine(Application.StartupPath, "public_key_for_client.xml");

            if (!System.IO.File.Exists(publicKeyPath))
            {
                MessageBox.Show("License validation system not configured properly.\nPlease contact support.",
                    "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }

            string publicKey = System.IO.File.ReadAllText(publicKeyPath);
            _validator = new LicenseSystem.LicenseValidator(publicKey);

            LoadHardwareId();
            CheckExistingLicense();
        }

        private void InitializeComponent()
        {
            this.Text = "Software Activation Required";
            this.Size = new Size(600, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            // Title
            lblTitle = new Label
            {
                Text = "🔒 C2B Smart App - License Activation",
                Location = new Point(20, 20),
                Size = new Size(540, 40),
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204),
                TextAlign = ContentAlignment.MiddleCenter
            };

            // Hardware ID Section
            var panelHardware = new GroupBox
            {
                Text = "Step 1: Your Hardware ID",
                Location = new Point(20, 80),
                Size = new Size(540, 120),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            lblHardwareInfo = new Label
            {
                Text = "Send this Hardware ID to us to receive your license key:",
                Location = new Point(15, 30),
                Size = new Size(510, 20),
                Font = new Font("Segoe UI", 9F)
            };

            txtHardwareId = new TextBox
            {
                Location = new Point(15, 55),
                Size = new Size(390, 25),
                ReadOnly = true,
                Font = new Font("Consolas", 10F),
                BackColor = Color.FromArgb(240, 240, 240)
            };

            btnCopyHardwareId = new Button
            {
                Text = "Copy",
                Location = new Point(415, 53),
                Size = new Size(100, 28),
                BackColor = Color.FromArgb(33, 150, 243),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnCopyHardwareId.FlatAppearance.BorderSize = 0;
            btnCopyHardwareId.Click += BtnCopyHardwareId_Click;

            panelHardware.Controls.AddRange(new Control[] { lblHardwareInfo, txtHardwareId, btnCopyHardwareId });

            // License Key Section
            var panelLicense = new GroupBox
            {
                Text = "Step 2: Enter License Key",
                Location = new Point(20, 220),
                Size = new Size(540, 120),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            lblLicenseInfo = new Label
            {
                Text = "Paste the license key you received from us:",
                Location = new Point(15, 30),
                Size = new Size(510, 20),
                Font = new Font("Segoe UI", 9F)
            };

            txtLicenseKey = new TextBox
            {
                Location = new Point(15, 55),
                Size = new Size(510, 40),
                Multiline = true,
                Font = new Font("Consolas", 9F),
                ScrollBars = ScrollBars.Vertical
            };

            panelLicense.Controls.AddRange(new Control[] { lblLicenseInfo, txtLicenseKey });

            // Status Label
            lblStatus = new Label
            {
                Location = new Point(20, 350),
                Size = new Size(540, 40),
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };

            // Buttons
            btnActivate = new Button
            {
                Text = "Activate License",
                Location = new Point(150, 410),
                Size = new Size(150, 40),
                BackColor = Color.FromArgb(46, 125, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnActivate.FlatAppearance.BorderSize = 0;
            btnActivate.Click += BtnActivate_Click;

            btnExit = new Button
            {
                Text = "Exit",
                Location = new Point(310, 410),
                Size = new Size(150, 40),
                BackColor = Color.FromArgb(211, 47, 47),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnExit.FlatAppearance.BorderSize = 0;
            btnExit.Click += (s, e) => Application.Exit();

            this.Controls.AddRange(new Control[] {
                lblTitle, panelHardware, panelLicense, lblStatus, btnActivate, btnExit
            });
        }

        private void LoadHardwareId()
        {
            try
            {
                string hardwareId = LicenseSystem.HardwareInfo.GetHardwareId();
                txtHardwareId.Text = LicenseSystem.HardwareInfo.FormatHardwareId(hardwareId);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting hardware ID: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void CheckExistingLicense()
        {
            string existingLicense = _validator.LoadLicense();

            if (!string.IsNullOrEmpty(existingLicense))
            {
                LicenseSystem.LicenseData licenseData;
                string errorMessage;

                if (_validator.ValidateLicense(existingLicense, out licenseData, out errorMessage))
                {
                    // Valid license found
                    IsActivated = true;
                    LicenseKey = existingLicense;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                    return;
                }
                else
                {
                    // Invalid or expired license
                    lblStatus.Text = $"Previous license invalid: {errorMessage}";
                    lblStatus.ForeColor = Color.Red;
                    _validator.DeleteLicense();
                }
            }
        }

        private void BtnCopyHardwareId_Click(object sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(txtHardwareId.Text);
                lblStatus.Text = "✅ Hardware ID copied to clipboard! Send this to us via email.";
                lblStatus.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying to clipboard: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnActivate_Click(object sender, EventArgs e)
        {
            string licenseKey = txtLicenseKey.Text.Trim();

            if (string.IsNullOrEmpty(licenseKey))
            {
                lblStatus.Text = "⚠️ Please enter a license key";
                lblStatus.ForeColor = Color.Red;
                return;
            }

            try
            {
                LicenseSystem.LicenseData licenseData;
                string errorMessage;

                btnActivate.Enabled = false;
                btnActivate.Text = "Validating...";
                Application.DoEvents();

                if (_validator.ValidateLicense(licenseKey, out licenseData, out errorMessage))
                {
                    // Save the license
                    _validator.SaveLicense(licenseKey);

                    IsActivated = true;
                    LicenseKey = licenseKey;

                    MessageBox.Show(
                        $"✅ License activated successfully!\n\n" +
                        $"Licensed to: {licenseData.CustomerEmail}\n" +
                        $"Valid until: {licenseData.ExpiryDate:yyyy-MM-dd}\n" +
                        $"Days remaining: {licenseData.DaysRemaining()}",
                        "Activation Successful",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    lblStatus.Text = $"❌ {errorMessage}";
                    lblStatus.ForeColor = Color.Red;

                    MessageBox.Show(
                        $"License activation failed:\n\n{errorMessage}\n\n" +
                        "Please contact support if you believe this is an error.",
                        "Activation Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"❌ Error: {ex.Message}";
                lblStatus.ForeColor = Color.Red;

                MessageBox.Show($"Error during activation: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnActivate.Enabled = true;
                btnActivate.Text = "Activate License";
            }
        }
    }
}