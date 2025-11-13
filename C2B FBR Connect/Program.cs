using C2B_FBR_Connect.Forms;
using System;
using System.Windows.Forms;

namespace C2B_FBR_Connect
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Check license before starting main application
            if (!CheckLicense())
            {
                // License check failed or user cancelled - exit application
                return;
            }

            // License is valid - start main application
            Application.Run(new MainForm());
        }

        private static bool CheckLicense()
        {
            try
            {
                // Load public key
                string publicKeyPath = System.IO.Path.Combine(Application.StartupPath, "public_key_for_client.xml");

                if (!System.IO.File.Exists(publicKeyPath))
                {
                    MessageBox.Show(
                        "License system configuration file not found.\n\n" +
                        "Please ensure 'public_key_for_client.xml' is in the application directory.\n\n" +
                        "Contact support for assistance.",
                        "Configuration Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return false;
                }

                string publicKey = System.IO.File.ReadAllText(publicKeyPath);
                var validator = new LicenseSystem.LicenseValidator(publicKey);

                // Check if valid license exists
                if (validator.IsLicenseValid())
                {
                    // License is valid - check if it's expiring soon
                    var licenseData = validator.GetCurrentLicenseData();

                    if (licenseData != null)
                    {
                        int daysRemaining = licenseData.DaysRemaining();

                        // Warn if license expires in 7 days or less
                        if (daysRemaining <= 7 && daysRemaining > 0)
                        {
                            MessageBox.Show(
                                $"⚠️ License Expiration Warning\n\n" +
                                $"Your license will expire in {daysRemaining} day(s).\n" +
                                $"Expiry Date: {licenseData.ExpiryDate:yyyy-MM-dd}\n\n" +
                                $"Please contact us to renew your license.",
                                "License Expiring Soon",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                        }
                    }

                    return true;
                }

                // No valid license - show activation form
                using (var activationForm = new ActivationForm())
                {
                    if (activationForm.ShowDialog() == DialogResult.OK && activationForm.IsActivated)
                    {
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error checking license:\n\n{ex.Message}\n\n" +
                    "The application will now close.",
                    "License Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }
    }
}