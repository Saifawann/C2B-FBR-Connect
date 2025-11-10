using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace C2B_FBR_Connect.Services
{
    /// <summary>
    /// Tracks uploaded invoices to avoid redundant QuickBooks queries
    /// Stores invoice upload status in a local JSON file
    /// </summary>
    public class InvoiceTrackingService
    {
        private readonly string _trackingFilePath;
        private Dictionary<string, InvoiceUploadRecord> _uploadedInvoices;
        private readonly object _lockObject = new object();

        public InvoiceTrackingService(string companyName = null)
        {
            // Store tracking file in app data directory
            string appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "C2B_FBR_Connect",
                "InvoiceTracking"
            );

            Directory.CreateDirectory(appDataPath);

            // Use company-specific file if provided
            string fileName = string.IsNullOrEmpty(companyName)
                ? "invoice_tracking.json"
                : $"invoice_tracking_{SanitizeFileName(companyName)}.json";

            _trackingFilePath = Path.Combine(appDataPath, fileName);
            LoadTrackingData();
        }

        #region Public Methods

        /// <summary>
        /// Check if invoice has been successfully uploaded
        /// </summary>
        public bool IsInvoiceUploaded(string qbInvoiceId)
        {
            lock (_lockObject)
            {
                return _uploadedInvoices.ContainsKey(qbInvoiceId) &&
                       _uploadedInvoices[qbInvoiceId].Status == UploadStatus.Success;
            }
        }

        /// <summary>
        /// Mark invoice as successfully uploaded
        /// </summary>
        public void MarkAsUploaded(string qbInvoiceId, string invoiceNumber, string irn, DateTime uploadDate)
        {
            lock (_lockObject)
            {
                _uploadedInvoices[qbInvoiceId] = new InvoiceUploadRecord
                {
                    QuickBooksInvoiceId = qbInvoiceId,
                    InvoiceNumber = invoiceNumber,
                    IRN = irn,
                    Status = UploadStatus.Success,
                    UploadDate = uploadDate,
                    LastAttemptDate = uploadDate
                };

                SaveTrackingData();
            }
        }

        /// <summary>
        /// Mark invoice as failed (for retry logic)
        /// </summary>
        public void MarkAsFailed(string qbInvoiceId, string invoiceNumber, string errorMessage)
        {
            lock (_lockObject)
            {
                if (_uploadedInvoices.TryGetValue(qbInvoiceId, out var existing))
                {
                    existing.Status = UploadStatus.Failed;
                    existing.LastAttemptDate = DateTime.Now;
                    existing.ErrorMessage = errorMessage;
                    existing.RetryCount++;
                }
                else
                {
                    _uploadedInvoices[qbInvoiceId] = new InvoiceUploadRecord
                    {
                        QuickBooksInvoiceId = qbInvoiceId,
                        InvoiceNumber = invoiceNumber,
                        Status = UploadStatus.Failed,
                        LastAttemptDate = DateTime.Now,
                        ErrorMessage = errorMessage,
                        RetryCount = 1
                    };
                }

                SaveTrackingData();
            }
        }

        /// <summary>
        /// Get upload status for an invoice
        /// </summary>
        public InvoiceUploadRecord GetUploadStatus(string qbInvoiceId)
        {
            lock (_lockObject)
            {
                return _uploadedInvoices.TryGetValue(qbInvoiceId, out var record) ? record : null;
            }
        }

        /// <summary>
        /// Get all uploaded invoices
        /// </summary>
        public List<InvoiceUploadRecord> GetAllUploadedInvoices()
        {
            lock (_lockObject)
            {
                return _uploadedInvoices.Values
                    .Where(r => r.Status == UploadStatus.Success)
                    .OrderByDescending(r => r.UploadDate)
                    .ToList();
            }
        }

        /// <summary>
        /// Get failed invoices that can be retried
        /// </summary>
        public List<InvoiceUploadRecord> GetFailedInvoices(int maxRetries = 3)
        {
            lock (_lockObject)
            {
                return _uploadedInvoices.Values
                    .Where(r => r.Status == UploadStatus.Failed && r.RetryCount < maxRetries)
                    .OrderBy(r => r.LastAttemptDate)
                    .ToList();
            }
        }

        /// <summary>
        /// Remove invoice from tracking (e.g., if invoice was deleted in QuickBooks)
        /// </summary>
        public void RemoveInvoice(string qbInvoiceId)
        {
            lock (_lockObject)
            {
                if (_uploadedInvoices.Remove(qbInvoiceId))
                {
                    SaveTrackingData();
                }
            }
        }

        /// <summary>
        /// Clear all tracking data (use with caution!)
        /// </summary>
        public void ClearAll()
        {
            lock (_lockObject)
            {
                _uploadedInvoices.Clear();
                SaveTrackingData();
            }
        }

        /// <summary>
        /// Get statistics about tracked invoices
        /// </summary>
        public TrackingStatistics GetStatistics()
        {
            lock (_lockObject)
            {
                return new TrackingStatistics
                {
                    TotalTracked = _uploadedInvoices.Count,
                    SuccessfulUploads = _uploadedInvoices.Values.Count(r => r.Status == UploadStatus.Success),
                    FailedUploads = _uploadedInvoices.Values.Count(r => r.Status == UploadStatus.Failed),
                    LastUploadDate = _uploadedInvoices.Values
                        .Where(r => r.Status == UploadStatus.Success)
                        .OrderByDescending(r => r.UploadDate)
                        .FirstOrDefault()?.UploadDate
                };
            }
        }

        /// <summary>
        /// Clean up old records (older than specified days)
        /// </summary>
        public int CleanupOldRecords(int daysToKeep = 90)
        {
            lock (_lockObject)
            {
                var cutoffDate = DateTime.Now.AddDays(-daysToKeep);
                var toRemove = _uploadedInvoices
                    .Where(kvp => kvp.Value.UploadDate.HasValue && kvp.Value.UploadDate.Value < cutoffDate)
                    .Select(kvp => kvp.Key)
                    .ToList();

                foreach (var key in toRemove)
                {
                    _uploadedInvoices.Remove(key);
                }

                if (toRemove.Count > 0)
                {
                    SaveTrackingData();
                }

                return toRemove.Count;
            }
        }

        #endregion

        #region Private Methods

        private void LoadTrackingData()
        {
            try
            {
                if (File.Exists(_trackingFilePath))
                {
                    var json = File.ReadAllText(_trackingFilePath);
                    _uploadedInvoices = JsonConvert.DeserializeObject<Dictionary<string, InvoiceUploadRecord>>(json)
                                       ?? new Dictionary<string, InvoiceUploadRecord>();
                }
                else
                {
                    _uploadedInvoices = new Dictionary<string, InvoiceUploadRecord>();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading invoice tracking: {ex.Message}");
                _uploadedInvoices = new Dictionary<string, InvoiceUploadRecord>();
            }
        }

        private void SaveTrackingData()
        {
            try
            {
                var json = JsonConvert.SerializeObject(_uploadedInvoices, Formatting.Indented);
                File.WriteAllText(_trackingFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving invoice tracking: {ex.Message}");
            }
        }

        private string SanitizeFileName(string fileName)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));
        }

        #endregion
    }

    #region Helper Classes

    public class InvoiceUploadRecord
    {
        public string QuickBooksInvoiceId { get; set; }
        public string InvoiceNumber { get; set; }
        public string IRN { get; set; }
        public UploadStatus Status { get; set; }
        public DateTime? UploadDate { get; set; }
        public DateTime LastAttemptDate { get; set; }
        public string ErrorMessage { get; set; }
        public int RetryCount { get; set; }
    }

    public enum UploadStatus
    {
        Pending,
        Success,
        Failed
    }

    public class TrackingStatistics
    {
        public int TotalTracked { get; set; }
        public int SuccessfulUploads { get; set; }
        public int FailedUploads { get; set; }
        public DateTime? LastUploadDate { get; set; }

        public override string ToString()
        {
            return $@"
📊 Invoice Tracking Statistics
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Total Tracked: {TotalTracked}
✅ Successful: {SuccessfulUploads}
❌ Failed: {FailedUploads}
Last Upload: {LastUploadDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Never"}
";
        }
    }

    #endregion
}