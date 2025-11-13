using System;

namespace LicenseSystem
{
    [Serializable]
    public class LicenseData
    {
        public string HardwareId { get; set; }
        public DateTime ExpiryDate { get; set; }
        public DateTime IssueDate { get; set; }
        public string CustomerEmail { get; set; }

        public LicenseData()
        {
            IssueDate = DateTime.Now;
        }

        public override string ToString()
        {
            return $"{HardwareId}|{ExpiryDate:yyyy-MM-dd}|{IssueDate:yyyy-MM-dd}|{CustomerEmail}";
        }

        public static LicenseData Parse(string data)
        {
            var parts = data.Split('|');
            if (parts.Length != 4)
                throw new FormatException("Invalid license data format");

            return new LicenseData
            {
                HardwareId = parts[0],
                ExpiryDate = DateTime.Parse(parts[1]),
                IssueDate = DateTime.Parse(parts[2]),
                CustomerEmail = parts[3]
            };
        }

        public bool IsExpired()
        {
            return DateTime.Now > ExpiryDate;
        }

        public int DaysRemaining()
        {
            return (ExpiryDate - DateTime.Now).Days;
        }
    }
}