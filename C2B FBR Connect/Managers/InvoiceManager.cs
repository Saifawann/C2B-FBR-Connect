using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using System;
using System.Collections.Generic;
using System.Linq;
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

        // ✅ Sets the current company context for operations
        public void SetCompany(Company company)
        {
            _currentCompany = company;
        }

        // ✅ Fetch invoices from QuickBooks and save to DB
        public List<Invoice> FetchFromQuickBooks()
        {
            try
            {
                if (!_qb.Connect(_currentCompany))
                    throw new Exception("Failed to connect to QuickBooks");

                var qbInvoices = _qb.FetchInvoices();

                foreach (var inv in qbInvoices)
                    _db.SaveInvoice(inv);

                return _db.GetInvoices(_qb.CurrentCompanyName);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error fetching invoices from QuickBooks: {ex.Message}", ex);
            }
        }

        // Updated UploadToFBR method
        public async Task<bool> UploadToFBR(Invoice invoice, string token)
        {
            try
            {
                if (_currentCompany == null)
                    throw new InvalidOperationException("Company not set. Call SetCompany() first.");

                if (string.IsNullOrEmpty(_currentCompany.SellerNTN))
                    throw new InvalidOperationException("Seller NTN not configured. Please update company settings.");

                // Get invoice details from QuickBooks
                var details = _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);
                if (details == null)
                {
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null,
                        "Could not retrieve invoice details from QuickBooks");
                    return false;
                }

                // ✅ Inject company details into payload
                details.SellerNTN = _currentCompany.SellerNTN;
                details.SellerBusinessName = _currentCompany.CompanyName;
                details.SellerProvince = _currentCompany.SellerProvince;
                details.SellerAddress = _currentCompany.SellerAddress;

                // ✅ Upload to FBR
                var response = await _fbr.UploadInvoice(details, token);

                if (response.Success)
                {
                    // Update invoice with FBR response data
                    invoice.FBR_IRN = response.IRN;
                    invoice.FBR_QRCode = response.IRN; // QR Code content is the IRN

                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Uploaded", response.IRN, null);

                    // ✅ Auto-generate PDF and save to Documents folder
                    try
                    {
                        string outputPath = GetDefaultPDFPath(invoice);
                        _pdf.GenerateInvoicePDF(invoice, details, outputPath);
                        Console.WriteLine($"PDF generated successfully: {outputPath}");
                    }
                    catch (Exception pdfEx)
                    {
                        // Log PDF generation error but don't fail the upload
                        Console.WriteLine($"Warning: PDF generation failed: {pdfEx.Message}");
                    }

                    return true;
                }
                else
                {
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null, response.ErrorMessage);
                    return false;
                }
            }
            catch (Exception ex)
            {
                _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null,
                    $"Exception: {ex.Message}");
                return false;
            }
        }

        // ✅ Upload multiple invoices to FBR in batch
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

        // ✅ Generate a PDF for invoice
        public void GeneratePDF(Invoice invoice, string outputPath)
        {
            try
            {
                var details = _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);
                if (details == null)
                    throw new Exception("Could not retrieve invoice details from QuickBooks");

                _pdf.GenerateInvoicePDF(invoice, details, outputPath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error generating PDF: {ex.Message}", ex);
            }
        }

        public List<Invoice> GetInvoices()
        {
            if (string.IsNullOrEmpty(_qb.CurrentCompanyName))
                throw new InvalidOperationException("No company connected. Connect to QuickBooks first.");

            return _db.GetInvoices(_qb.CurrentCompanyName);
        }

        public List<Invoice> GetInvoicesByStatus(string status)
        {
            return GetInvoices().Where(i => i.Status == status).ToList();
        }

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

        // Add this helper method to InvoiceManager class
        private string GetDefaultPDFPath(Invoice invoice)
        {
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string invoiceFolder = Path.Combine(documentsPath, "FBR_Invoices");

            if (!Directory.Exists(invoiceFolder))
                Directory.CreateDirectory(invoiceFolder);

            string fileName = $"Invoice_{invoice.InvoiceNumber}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            return Path.Combine(invoiceFolder, fileName);
        }

        
    }
}
