using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Managers
{
    public class InvoiceManager
    {
        private readonly DatabaseService _db;
        private readonly QuickBooksService _qb;
        private readonly FBRApiService _fbr;
        private readonly PDFService _pdf;
        private Company _currentCompany;

        public InvoiceManager(DatabaseService db, QuickBooksService qb,
            FBRApiService fbr, PDFService pdf)
        {
            _db = db;
            _qb = qb;
            _fbr = fbr;
            _pdf = pdf;
        }

        /// <summary>
        /// Sets the current company context for operations
        /// </summary>
        public void SetCompany(Company company)
        {
            _currentCompany = company;
        }

        /// <summary>
        /// Fetches invoices from QuickBooks and saves them to the database
        /// </summary>
        /// <param name="fromDate">Optional start date for invoice filtering</param>
        /// <returns>List of invoices from the database</returns>
        public List<Invoice> FetchFromQuickBooks()
        {
            try
            {
                // Ensure QuickBooks is connected with company settings
                if (!_qb.Connect(_currentCompany))
                {
                    throw new Exception("Failed to connect to QuickBooks");
                }

                var qbInvoices = _qb.FetchInvoices();

                foreach (var inv in qbInvoices)
                {
                    _db.SaveInvoice(inv);
                }

                return _db.GetInvoices(_qb.CurrentCompanyName);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error fetching invoices from QuickBooks: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Uploads a single invoice to FBR
        /// </summary>
        public async Task<bool> UploadToFBR(Invoice invoice, string token)
        {
            try
            {
                // Validate company configuration
                if (_currentCompany == null)
                {
                    throw new InvalidOperationException("Company not set. Call SetCompany() first.");
                }

                if (string.IsNullOrEmpty(_currentCompany.SellerNTN))
                {
                    throw new InvalidOperationException("Seller NTN not configured. Please update company settings.");
                }

                // Get full invoice details from QuickBooks
                var details = _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);

                if (details == null)
                {
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed",
                        null, "Could not retrieve invoice details from QuickBooks");
                    return false;
                }

                // Upload to FBR
                var response = await _fbr.UploadInvoice(details, token);

                if (response.Success)
                {
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Uploaded",
                        response.IRN, null);
                    return true;
                }
                else
                {
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed",
                        null, response.ErrorMessage);
                    return false;
                }
            }
            catch (Exception ex)
            {
                _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed",
                    null, $"Exception: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Uploads multiple invoices to FBR in batch
        /// </summary>
        public async Task<Dictionary<string, bool>> UploadMultipleToFBR(List<Invoice> invoices, string token)
        {
            var results = new Dictionary<string, bool>();

            foreach (var invoice in invoices)
            {
                var success = await UploadToFBR(invoice, token);
                results[invoice.InvoiceNumber] = success;
            }

            return results;
        }

        /// <summary>
        /// Generates a PDF for the invoice
        /// </summary>
        public void GeneratePDF(Invoice invoice, string outputPath)
        {
            try
            {
                var details = _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);

                if (details == null)
                {
                    throw new Exception("Could not retrieve invoice details from QuickBooks");
                }

                _pdf.GenerateInvoicePDF(invoice, details, outputPath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error generating PDF: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Gets all invoices for the current company
        /// </summary>
        public List<Invoice> GetInvoices()
        {
            if (string.IsNullOrEmpty(_qb.CurrentCompanyName))
            {
                throw new InvalidOperationException("No company connected. Connect to QuickBooks first.");
            }

            return _db.GetInvoices(_qb.CurrentCompanyName);
        }

        /// <summary>
        /// Gets invoices filtered by status
        /// </summary>
        public List<Invoice> GetInvoicesByStatus(string status)
        {
            return GetInvoices().Where(i => i.Status == status).ToList();
        }

        /// <summary>
        /// Retries failed invoice uploads
        /// </summary>
        public async Task<int> RetryFailedInvoices(string token)
        {
            var failedInvoices = GetInvoicesByStatus("Failed");
            int successCount = 0;

            foreach (var invoice in failedInvoices)
            {
                var success = await UploadToFBR(invoice, token);
                if (success) successCount++;
            }

            return successCount;
        }
    }
}