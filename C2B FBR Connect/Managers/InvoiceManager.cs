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
        public async Task<FBRResponse> UploadToFBR(Invoice invoice, string token)
        {
            var result = new FBRResponse();

            try
            {
                if (_currentCompany == null)
                    throw new InvalidOperationException("Company not set. Call SetCompany() first.");

                if (string.IsNullOrEmpty(_currentCompany.SellerNTN))
                    throw new InvalidOperationException("Seller NTN not configured. Please update company settings.");

                var details = await _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);
                if (details == null)
                {
                    result.Success = false;
                    result.ErrorMessage = "Could not retrieve invoice details from QuickBooks";
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null, result.ErrorMessage);
                    return result;
                }

                // Add company details
                details.SellerNTN = _currentCompany.SellerNTN;
                details.SellerBusinessName = _currentCompany.CompanyName;
                details.SellerProvince = _currentCompany.SellerProvince;
                details.SellerAddress = _currentCompany.SellerAddress;

                // Call FBR API
                var response = await _fbr.UploadInvoice(details, token);
                result = response;

                if (response.Success)
                {
                    invoice.FBR_IRN = response.IRN;
                    invoice.FBR_QRCode = response.IRN;

                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Uploaded", response.IRN, null);

                    // Generate PDF
                    try
                    {
                        string outputPath = GetDefaultPDFPath(invoice);
                        _pdf.GenerateInvoicePDF(invoice, details, outputPath);
                        Console.WriteLine($"PDF generated successfully: {outputPath}");
                    }
                    catch (Exception pdfEx)
                    {
                        Console.WriteLine($"Warning: PDF generation failed: {pdfEx.Message}");
                    }
                }
                else
                {
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null, response.ErrorMessage);
                }

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null, $"Exception: {ex.Message}");
                return result;
            }
        }


        // ✅ Upload multiple invoices to FBR in batch
        public async Task<Dictionary<string, FBRResponse>> UploadMultipleToFBR(List<Invoice> invoices, string token)
        {
            var results = new Dictionary<string, FBRResponse>();

            foreach (var invoice in invoices)
            {
                var response = await UploadToFBR(invoice, token);
                results[invoice.InvoiceNumber] = response;
            }

            return results;
        }


        // ✅ Generate a PDF for invoice
        public async void GeneratePDF(Invoice invoice, string outputPath)
        {
            try
            {
                var details = await _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);
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
                var response = await UploadToFBR(invoice, token);

                if (response.Success)
                    successCount++;
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
