using System;
using System.Management;
using System.Security.Cryptography;
using System.Text;

namespace LicenseSystem
{
    public class HardwareInfo
    {
        public static string GetHardwareId()
        {
            try
            {
                string cpuId = GetCpuId();
                string motherboardId = GetMotherboardId();
                string diskId = GetDiskId();

                // Combine hardware identifiers
                string combined = $"{cpuId}-{motherboardId}-{diskId}";

                // Create a hash for consistent length
                using (SHA256 sha256 = SHA256.Create())
                {
                    byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(combined));
                    return BitConverter.ToString(hashBytes).Replace("-", "").Substring(0, 32);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to get hardware ID: " + ex.Message);
            }
        }

        private static string GetCpuId()
        {
            string cpuId = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor");
                foreach (ManagementObject obj in searcher.Get())
                {
                    cpuId = obj["ProcessorId"]?.ToString() ?? "";
                    break;
                }
            }
            catch { }
            return cpuId;
        }

        private static string GetMotherboardId()
        {
            string motherboardId = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");
                foreach (ManagementObject obj in searcher.Get())
                {
                    motherboardId = obj["SerialNumber"]?.ToString() ?? "";
                    break;
                }
            }
            catch { }
            return motherboardId;
        }

        private static string GetDiskId()
        {
            string diskId = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_DiskDrive");
                foreach (ManagementObject obj in searcher.Get())
                {
                    diskId = obj["SerialNumber"]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(diskId))
                        break;
                }
            }
            catch { }
            return diskId;
        }

        public static string FormatHardwareId(string hardwareId)
        {
            // Format as XXXX-XXXX-XXXX-XXXX-XXXX-XXXX-XXXX-XXXX
            if (string.IsNullOrEmpty(hardwareId) || hardwareId.Length < 32)
                return hardwareId;

            StringBuilder formatted = new StringBuilder();
            for (int i = 0; i < 32; i += 4)
            {
                if (i > 0) formatted.Append("-");
                formatted.Append(hardwareId.Substring(i, 4));
            }
            return formatted.ToString();
        }
    }
}