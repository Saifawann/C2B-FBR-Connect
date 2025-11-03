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

        public async Task<List<Invoice>> FetchFromQuickBooks()
        {
            try
            {
                if (!_qb.Connect(_currentCompany))
                    throw new Exception("Failed to connect to QuickBooks");

                // ✅ CLEAR CACHE TO FORCE FRESH DATA
                Console.WriteLine("🔄 Starting invoice fetch from QuickBooks...");
                Console.WriteLine("🗑️ Clearing QuickBooks cache to fetch fresh data...");

                // Fetch basic invoice list from QuickBooks
                var qbInvoices = _qb.FetchInvoices();
                Console.WriteLine($"📥 Found {qbInvoices.Count} invoices in QuickBooks\n");

                int updatedCount = 0;
                int newCount = 0;
                int preservedCount = 0;

                // For each invoice, fetch complete details including line items
                foreach (var inv in qbInvoices)
                {
                    try
                    {
                        Console.WriteLine($"\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                        Console.WriteLine($"🔹 Processing invoice: {inv.InvoiceNumber}");

                        // ✅ STEP 1: Check if invoice exists in database FIRST
                        var existingInvoice = _db.GetInvoiceWithDetails(
                            inv.QuickBooksInvoiceId,
                            _qb.CurrentCompanyName
                        );

                        // ✅ STEP 2: Get fresh detailed invoice data from QuickBooks (includes items with Sale Type)
                        Console.WriteLine($"   🔍 Fetching detailed invoice data from QuickBooks...");
                        var details = await _qb.GetInvoiceDetails(inv.QuickBooksInvoiceId);

                        if (details == null)
                        {
                            Console.WriteLine($"   ⚠️ Could not fetch details for {inv.InvoiceNumber}");
                            continue;
                        }

                        Console.WriteLine($"   ✅ Fetched {details.Items?.Count ?? 0} items from QuickBooks");

                        // ✅ STEP 3: Map invoice header data
                        inv.TotalAmount = details.TotalAmount;
                        inv.TaxAmount = details.TaxAmount;
                        inv.DiscountAmount = details.Items?.Sum(i => i.Discount) ?? 0;
                        inv.InvoiceDate = details.InvoiceDate;
                        inv.CustomerAddress = details.BuyerAddress;
                        inv.CustomerPhone = details.BuyerPhone;
                        inv.CustomerNTN = details.CustomerNTN;
                        inv.CustomerType = details.BuyerRegistrationType;
                        inv.CustomerEmail = details.BuyerEmail ?? "";
                        inv.PaymentMode = "Cash";

                        // ✅ STEP 4: Map invoice items FROM THE DETAILED PAYLOAD (which has Sale Type!)
                        inv.Items = new List<InvoiceItem>();

                        if (details.Items != null && details.Items.Count > 0)
                        {
                            Console.WriteLine($"   📦 Mapping {details.Items.Count} items:");

                            foreach (var fbrItem in details.Items)
                            {
                                Console.WriteLine($"      - {fbrItem.ItemName}");
                                Console.WriteLine($"        Sale Type: '{fbrItem.SaleType}'");
                                Console.WriteLine($"        HS Code: '{fbrItem.HSCode}'");
                                Console.WriteLine($"        Quantity: {fbrItem.Quantity}");

                                var invoiceItem = new InvoiceItem
                                {
                                    ItemName = fbrItem.ItemName ?? "",
                                    ItemDescription = fbrItem.ItemName ?? "", // ✅ Use ItemName as description
                                    Quantity = fbrItem.Quantity,
                                    UnitPrice = fbrItem.UnitPrice,
                                    TotalPrice = fbrItem.TotalPrice,
                                    NetAmount = fbrItem.NetAmount,
                                    TaxRate = fbrItem.TaxRate,
                                    SalesTaxAmount = fbrItem.SalesTaxAmount,
                                    TotalValue = fbrItem.TotalValue,
                                    HSCode = fbrItem.HSCode ?? "",
                                    UnitOfMeasure = fbrItem.UnitOfMeasure ?? "",
                                    RetailPrice = fbrItem.RetailPrice,
                                    ExtraTax = fbrItem.ExtraTax,
                                    FurtherTax = fbrItem.FurtherTax,
                                    FedPayable = fbrItem.FedPayable,
                                    SalesTaxWithheldAtSource = fbrItem.SalesTaxWithheldAtSource,
                                    Discount = fbrItem.Discount,
                                    SaleType = fbrItem.SaleType ?? "Goods at standard rate (default)", // ✅ This comes from QB
                                    SroScheduleNo = fbrItem.SroScheduleNo ?? "",
                                    SroItemSerialNo = fbrItem.SroItemSerialNo ?? ""
                                };

                                inv.Items.Add(invoiceItem);
                            }
                        }
                        else
                        {
                            Console.WriteLine($"   ⚠️ No items found for this invoice");
                        }

                        // ✅ STEP 5: Preserve FBR upload status if invoice was already uploaded
                        if (existingInvoice != null && existingInvoice.Status == "Uploaded")
                        {
                            Console.WriteLine($"   📌 Preserving upload status for {inv.InvoiceNumber}");

                            inv.Status = "Uploaded";
                            inv.FBR_IRN = existingInvoice.FBR_IRN;
                            inv.FBR_QRCode = existingInvoice.FBR_QRCode;
                            inv.UploadDate = existingInvoice.UploadDate;
                            inv.ErrorMessage = existingInvoice.ErrorMessage;

                            preservedCount++;
                        }
                        else if (existingInvoice != null)
                        {
                            Console.WriteLine($"   🔄 Updating existing invoice {inv.InvoiceNumber}");

                            // Keep the existing status if not uploaded
                            inv.Status = existingInvoice.Status;
                            inv.ErrorMessage = existingInvoice.ErrorMessage;

                            updatedCount++;
                        }
                        else
                        {
                            Console.WriteLine($"   ✨ New invoice detected: {inv.InvoiceNumber}");

                            inv.Status = "Pending";
                            newCount++;
                        }

                        // ✅ STEP 6: Save invoice with ALL details to database
                        Console.WriteLine($"   💾 Saving to database...");
                        _db.SaveInvoiceWithDetails(inv);

                        Console.WriteLine($"   ✅ Saved {inv.InvoiceNumber} with {inv.Items?.Count ?? 0} items");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Error processing invoice {inv.InvoiceNumber}: {ex.Message}");
                        Console.WriteLine($"   Stack: {ex.StackTrace}");

                        // Still try to save basic invoice data
                        try
                        {
                            inv.Status = "Error";
                            inv.ErrorMessage = ex.Message;
                            _db.SaveInvoice(inv);
                        }
                        catch (Exception saveEx)
                        {
                            Console.WriteLine($"   ❌ Could not save error state: {saveEx.Message}");
                        }
                    }
                }

                Console.WriteLine($"\n╔════════════════════════════════════════╗");
                Console.WriteLine($"║         FETCH SUMMARY                  ║");
                Console.WriteLine($"╚════════════════════════════════════════╝");
                Console.WriteLine($"✨ New invoices: {newCount}");
                Console.WriteLine($"🔄 Updated invoices: {updatedCount}");
                Console.WriteLine($"📌 Preserved uploaded: {preservedCount}");
                Console.WriteLine($"📝 Total processed: {qbInvoices.Count}");
                Console.WriteLine($"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

                return _db.GetInvoices(_qb.CurrentCompanyName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ FATAL ERROR in FetchFromQuickBooks:");
                Console.WriteLine($"   Message: {ex.Message}");
                Console.WriteLine($"   Stack: {ex.StackTrace}");
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