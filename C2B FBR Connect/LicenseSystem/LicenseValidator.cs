using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace LicenseSystem
{
    public class LicenseValidator
    {
        private RSA _rsa;
        private readonly string _licenseFilePath;
        private readonly string _publicKey;

        public LicenseValidator(string publicKeyXml)
        {
            _publicKey = publicKeyXml;
            _rsa = RSA.Create();
            _rsa.FromXmlString(_publicKey);

            // Store license in AppData
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appDataPath, "C2B_FBR_Connect");

            if (!Directory.Exists(appFolder))
                Directory.CreateDirectory(appFolder);

            _licenseFilePath = Path.Combine(appFolder, "license.lic");
        }

        public bool ValidateLicense(string licenseKey, out LicenseData licenseData, out string errorMessage)
        {
            licenseData = null;
            errorMessage = string.Empty;

            try
            {
                var parts = licenseKey.Split('.');
                if (parts.Length != 2)
                {
                    errorMessage = "Invalid license key format";
                    return false;
                }

                byte[] dataBytes = Convert.FromBase64String(parts[0]);
                byte[] signature = Convert.FromBase64String(parts[1]);

                // Verify signature with public key
                bool isValid = _rsa.VerifyData(dataBytes, signature,
                    HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                if (!isValid)
                {
                    errorMessage = "License key signature is invalid";
                    return false;
                }

                // Parse license data
                string data = Encoding.UTF8.GetString(dataBytes);
                licenseData = LicenseData.Parse(data);

                // Check if license is for this hardware
                string currentHardwareId = HardwareInfo.GetHardwareId();
                if (!licenseData.HardwareId.Equals(currentHardwareId, StringComparison.OrdinalIgnoreCase))
                {
                    errorMessage = "License key is not valid for this hardware";
                    return false;
                }

                // Check if license is expired
                if (licenseData.IsExpired())
                {
                    errorMessage = $"License expired on {licenseData.ExpiryDate:yyyy-MM-dd}";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"Error validating license: {ex.Message}";
                return false;
            }
        }

        public void SaveLicense(string licenseKey)
        {
            try
            {
                // Encrypt the license before saving (optional extra security)
                byte[] encryptedData = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(licenseKey),
                    null,
                    DataProtectionScope.CurrentUser);

                File.WriteAllBytes(_licenseFilePath, encryptedData);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to save license: {ex.Message}");
            }
        }

        public string LoadLicense()
        {
            try
            {
                if (!File.Exists(_licenseFilePath))
                    return null;

                byte[] encryptedData = File.ReadAllBytes(_licenseFilePath);
                byte[] decryptedData = ProtectedData.Unprotect(
                    encryptedData,
                    null,
                    DataProtectionScope.CurrentUser);

                return Encoding.UTF8.GetString(decryptedData);
            }
            catch
            {
                return null;
            }
        }

        public bool IsLicenseValid()
        {
            string license = LoadLicense();
            if (string.IsNullOrEmpty(license))
                return false;

            LicenseData licenseData;
            string errorMessage;
            return ValidateLicense(license, out licenseData, out errorMessage);
        }

        public LicenseData GetCurrentLicenseData()
        {
            string license = LoadLicense();
            if (string.IsNullOrEmpty(license))
                return null;

            LicenseData licenseData;
            string errorMessage;

            if (ValidateLicense(license, out licenseData, out errorMessage))
                return licenseData;

            return null;
        }

        public void DeleteLicense()
        {
            try
            {
                if (File.Exists(_licenseFilePath))
                    File.Delete(_licenseFilePath);
            }
            catch { }
        }
    }
}