using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using System;
using System.Collections.Generic;
using System.IO;
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

        public void SetCompany(Company company)
        {
            _currentCompany = company;
        }

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

        public async Task<FBRResponse> UploadToFBR(Invoice invoice, string token)
        {
            var result = new FBRResponse();

            try
            {
                if (_currentCompany == null)
                    throw new InvalidOperationException("Company not set. Call SetCompany() first.");

                if (string.IsNullOrEmpty(_currentCompany.SellerNTN))
                    throw new InvalidOperationException("Seller NTN not configured. Please update company settings.");

                // Get full invoice details from QuickBooks
                var details = await _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);
                if (details == null)
                {
                    result.Success = false;
                    result.ErrorMessage = "Could not retrieve invoice details from QuickBooks";
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null, result.ErrorMessage);
                    return result;
                }

                // Map FBRInvoicePayload data to Invoice model for complete database storage
                invoice.TotalAmount = details.TotalAmount;
                invoice.TaxAmount = details.TaxAmount;
                invoice.DiscountAmount = details.Items?.Sum(i => i.Discount) ?? 0;
                invoice.InvoiceDate = details.InvoiceDate;
                invoice.CustomerAddress = details.BuyerAddress;
                invoice.CustomerPhone = details.BuyerPhone;
                invoice.CustomerNTN = details.CustomerNTN;
                invoice.PaymentMode = "Cash"; // You can get this from details if available

                // Map invoice items
                invoice.Items = details.Items?.Select(i => new InvoiceItem
                {
                    ItemName = i.ItemName,
                    Quantity = (int)i.Quantity,
                    UnitPrice = i.UnitPrice,
                    TotalPrice = i.TotalPrice,
                    TaxRate = i.TaxRate,
                    SalesTaxAmount = i.SalesTaxAmount,
                    TotalValue = i.TotalPrice + i.SalesTaxAmount,
                    HSCode = i.HSCode,
                    UnitOfMeasure = i.UnitOfMeasure,
                    RetailPrice = i.RetailPrice,
                    ExtraTax = i.ExtraTax,
                    FurtherTax = i.FurtherTax,
                    FedPayable = i.FedPayable,
                    SalesTaxWithheldAtSource = i.SalesTaxWithheldAtSource,
                    Discount = i.Discount,
                    SaleType = i.SaleType
                }).ToList() ?? new List<InvoiceItem>();

                // Add company details to FBR payload
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
                    invoice.Status = "Uploaded";
                    invoice.UploadDate = DateTime.Now;

                    // Save complete invoice data to database (including items)
                    _db.SaveInvoiceWithDetails(invoice);

                    // Generate PDF from database data
                    try
                    {
                        var invoiceFromDb = _db.GetInvoiceWithDetails(
                            invoice.QuickBooksInvoiceId,
                            _currentCompany.CompanyName);

                        if (invoiceFromDb != null)
                        {
                            string outputPath = GetDefaultPDFPath(invoice);
                            _pdf.GenerateInvoicePDF(invoiceFromDb, _currentCompany, outputPath);
                            Console.WriteLine($"PDF generated successfully: {outputPath}");
                        }
                    }
                    catch (Exception pdfEx)
                    {
                        Console.WriteLine($"Warning: PDF generation failed: {pdfEx.Message}");
                        // Don't fail the whole operation if PDF fails
                    }
                }
                else
                {
                    invoice.Status = "Failed";
                    invoice.ErrorMessage = response.ErrorMessage;
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

        public void GeneratePDF(Invoice invoice, string outputPath)
        {
            try
            {
                if (_currentCompany == null)
                    throw new InvalidOperationException("Company not set. Call SetCompany() first.");

                // Fetch complete invoice data from database
                var invoiceFromDb = _db.GetInvoiceWithDetails(
                    invoice.QuickBooksInvoiceId,
                    _currentCompany.CompanyName);

                if (invoiceFromDb == null)
                    throw new Exception("Invoice not found in database");

                _pdf.GenerateInvoicePDF(invoiceFromDb, _currentCompany, outputPath);
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