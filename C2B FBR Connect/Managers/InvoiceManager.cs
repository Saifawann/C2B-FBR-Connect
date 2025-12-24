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

        /// <summary>
        /// Fetch invoices/credit memos from QuickBooks and save to database
        /// Always fetches fresh data - no caching
        /// </summary>
        public async Task<List<Invoice>> FetchFromQuickBooks(DateTime? dateFrom = null, DateTime? dateTo = null,
            bool includeInvoices = true, bool includeCreditMemos = false)
        {
            try
            {
                if (!_qb.Connect(_currentCompany))
                    throw new Exception("Failed to connect to QuickBooks");

                Console.WriteLine("🔄 Fetching transactions from QuickBooks...");

                string dateRangeInfo = "";
                if (dateFrom.HasValue && dateTo.HasValue)
                    dateRangeInfo = $" from {dateFrom.Value:dd-MMM-yyyy} to {dateTo.Value:dd-MMM-yyyy}";
                else if (dateFrom.HasValue)
                    dateRangeInfo = $" from {dateFrom.Value:dd-MMM-yyyy}";
                else if (dateTo.HasValue)
                    dateRangeInfo = $" up to {dateTo.Value:dd-MMM-yyyy}";

                Console.WriteLine($"📅 Date range{dateRangeInfo}");

                // ========================================
                // STEP 1: Fetch basic invoice data from QuickBooks
                // ========================================
                var qbTransactions = _qb.FetchTransactions(dateFrom, dateTo, true, includeInvoices, includeCreditMemos);
                Console.WriteLine($"📥 Found {qbTransactions.Count} transactions in QuickBooks\n");

                if (qbTransactions.Count == 0)
                {
                    Console.WriteLine("No transactions found in QuickBooks for this date range");
                    return new List<Invoice>();
                }

                // ========================================
                // STEP 2: Extract unique customer IDs from invoices
                // ========================================
                Console.WriteLine($"👥 Extracting unique customers...");
                var uniqueCustomerIds = qbTransactions
                    .Where(i => !string.IsNullOrEmpty(i.QuickBooksCustomerId))
                    .Select(i => i.QuickBooksCustomerId)
                    .Distinct()
                    .ToList();

                Console.WriteLine($"📊 Found {uniqueCustomerIds.Count} unique customers\n");

                // ========================================
                // STEP 3: Fetch/Update customer data in database
                // ========================================
                Console.WriteLine($"🔄 Fetching/updating customer data from QuickBooks...");
                var customerMap = new Dictionary<string, Customer>();

                foreach (var qbCustomerId in uniqueCustomerIds)
                {
                    try
                    {
                        // Fetch fresh customer data from QuickBooks
                        var customerData = _qb.FetchCustomerDetails(qbCustomerId);

                        // Get or create customer in database
                        var customer = _db.GetOrCreateCustomer(
                            _qb.CurrentCompanyName,
                            qbCustomerId,
                            customerData
                        );

                        customerMap[qbCustomerId] = customer;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️ Error fetching customer {qbCustomerId}: {ex.Message}");
                        // Continue with other customers
                    }
                }

                Console.WriteLine($"✅ Processed {customerMap.Count} customers\n");

                // ========================================
                // STEP 4: Load existing invoices for status preservation
                // ========================================
                Console.WriteLine($"📊 Loading existing invoices from database...");
                var existingInvoices = _db.GetInvoicesWithDetails(_qb.CurrentCompanyName)
                    .ToDictionary(i => i.QuickBooksInvoiceId, i => i);
                Console.WriteLine($"✅ Loaded {existingInvoices.Count} existing invoices\n");

                // ========================================
                // STEP 5: Process and save invoices with customer references
                // ========================================
                int updatedCount = 0;
                int newCount = 0;

                Console.WriteLine($"💾 Saving invoices with customer references...");

                foreach (var transaction in qbTransactions)
                {
                    try
                    {
                        // ✅ FIXED: Map customer to invoice
                        if (customerMap.TryGetValue(transaction.QuickBooksCustomerId, out var customer))
                        {
                            transaction.CustomerId = customer.Id;
                            transaction.Customer = customer;
                            // ✅ That's it! The Invoice model's helper properties will handle the rest
                        }
                        else
                        {
                            Console.WriteLine($"⚠️ Customer not found for invoice {transaction.InvoiceNumber}");
                            continue;
                        }

                        // Check if invoice exists
                        existingInvoices.TryGetValue(transaction.QuickBooksInvoiceId, out var existingInvoice);

                        if (existingInvoice != null)
                        {
                            // Preserve FBR-specific data
                            transaction.Id = existingInvoice.Id;
                            transaction.Status = existingInvoice.Status;
                            transaction.FBR_IRN = existingInvoice.FBR_IRN;
                            transaction.FBR_QRCode = existingInvoice.FBR_QRCode;
                            transaction.UploadDate = existingInvoice.UploadDate;
                            transaction.ErrorMessage = existingInvoice.ErrorMessage;

                            // For uploaded invoices, preserve complete item details
                            if (existingInvoice.Status == "Uploaded" && existingInvoice.Items != null && existingInvoice.Items.Count > 0)
                            {
                                transaction.Items = existingInvoice.Items;
                            }

                            Console.WriteLine($"   🔄 {transaction.InvoiceNumber} - Updated" +
                                (existingInvoice.Status == "Uploaded" ? " (Uploaded)" :
                                 existingInvoice.Status == "Failed" ? " (Failed)" : " (Pending)"));
                            updatedCount++;
                        }
                        else
                        {
                            transaction.Status = "Pending";
                            Console.WriteLine($"   ✨ {transaction.InvoiceNumber} - New");
                            newCount++;
                        }

                        // Save to database with customer reference
                        _db.SaveInvoice(transaction);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error processing invoice {transaction.InvoiceNumber}: {ex.Message}");
                    }
                }

                Console.WriteLine($"\n╔════════════════════════════════════════╗");
                Console.WriteLine($"║       FETCH SUMMARY                    ║");
                Console.WriteLine($"╚════════════════════════════════════════╝");
                Console.WriteLine($"👥 Customers: {customerMap.Count} processed");
                Console.WriteLine($"✨ New Invoices: {newCount}");
                Console.WriteLine($"🔄 Updated Invoices: {updatedCount}");
                Console.WriteLine($"📝 Total: {qbTransactions.Count}");
                Console.WriteLine($"💾 All data refreshed with normalized customers!");
                Console.WriteLine($"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

                return _db.GetInvoices(_qb.CurrentCompanyName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ FATAL ERROR in FetchFromQuickBooks:");
                Console.WriteLine($"   Message: {ex.Message}");
                throw new Exception($"Error fetching transactions from QuickBooks: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Get full FBR payload for an invoice (always fetches fresh data)
        /// Called by UploadToFBR and ShowInvoiceDetailsAsync
        /// </summary>
        public async Task<FBRInvoicePayload> GetFullInvoicePayload(Invoice invoice)
        {
            // Always fetch fresh details from QuickBooks
            Console.WriteLine($"🔍 Fetching fresh details for {invoice.InvoiceType ?? "Invoice"}: {invoice.InvoiceNumber}");

            FBRInvoicePayload details = null;

            if (invoice.InvoiceType == "Credit Memo")
                details = await _qb.GetCreditMemoDetails(invoice.QuickBooksInvoiceId);
            else
                details = await _qb.GetInvoiceDetails(invoice.QuickBooksInvoiceId);

            if (details != null)
            {
                // Add company details
                details.SellerNTN = _currentCompany?.SellerNTN ?? "";
                details.SellerBusinessName = _currentCompany?.CompanyName ?? "";
                details.SellerProvince = _currentCompany?.SellerProvince ?? "";
                details.SellerAddress = _currentCompany?.SellerAddress ?? "";

                Console.WriteLine($"✅ Retrieved {details.Items?.Count ?? 0} items");
            }

            return details;
        }

        /// <summary>
        /// Upload invoice to FBR (always fetches fresh data before upload)
        /// </summary>
        public async Task<FBRResponse> UploadToFBR(Invoice invoice, string token)
        {
            var result = new FBRResponse();

            try
            {
                ValidateUploadPrerequisites();

                // Get fresh payload from QuickBooks
                var details = await GetFullInvoicePayload(invoice);

                if (details == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Could not retrieve {invoice.InvoiceType ?? "invoice"} details from QuickBooks";
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null, result.ErrorMessage);
                    return result;
                }

                // Map payload to invoice for database storage
                MapPayloadToInvoice(invoice, details);

                string environment = _currentCompany?.Environment ?? "Sandbox";

                // Call FBR API
                Console.WriteLine($"📤 Uploading {invoice.InvoiceType ?? "Invoice"} to FBR...");
                var response = await _fbr.UploadInvoice(details, token, environment);
                result = response;

                if (response.Success)
                {
                    invoice.FBR_IRN = response.IRN;
                    invoice.FBR_QRCode = response.IRN;
                    invoice.Status = "Uploaded";
                    invoice.UploadDate = DateTime.Now;
                    invoice.ErrorMessage = null;

                    // Save complete invoice with fresh items
                    _db.SaveInvoiceWithDetails(invoice);
                    Console.WriteLine($"✅ {invoice.InvoiceType ?? "Invoice"} {invoice.InvoiceNumber} uploaded successfully (IRN: {response.IRN})");

                    // Generate PDF
                    _ = GeneratePDFAsync(invoice);
                }
                else
                {
                    invoice.Status = "Failed";
                    invoice.ErrorMessage = response.ErrorMessage;
                    _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null, response.ErrorMessage);
                    Console.WriteLine($"❌ Upload failed: {response.ErrorMessage}");
                }

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _db.UpdateInvoiceStatus(invoice.QuickBooksInvoiceId, "Failed", null, $"Exception: {ex.Message}");
                Console.WriteLine($"❌ Exception during upload: {ex.Message}");
                return result;
            }
        }

        /// <summary>
        /// Upload multiple invoices
        /// </summary>
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

        /// <summary>
        /// Generate PDF
        /// </summary>
        public void GeneratePDF(Invoice invoice, string outputPath)
        {
            try
            {
                if (_currentCompany == null)
                    throw new InvalidOperationException("Company not set. Call SetCompany() first.");

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

        #region Private Helper Methods

        private void MapPayloadToInvoice(Invoice invoice, FBRInvoicePayload details)
        {
            // ✅ FIXED: Load customer if not already loaded
            if (invoice.Customer == null && invoice.CustomerId > 0)
            {
                invoice.Customer = _db.GetCustomerById(invoice.CustomerId);
            }

            if (invoice.Customer == null)
                throw new Exception("Customer not loaded for invoice");

            // Update invoice fields
            invoice.TotalAmount = details.TotalAmount;
            invoice.TaxAmount = details.TaxAmount;
            invoice.DiscountAmount = details.Items?.Sum(i => i.Discount) ?? 0;
            invoice.InvoiceDate = details.InvoiceDate;
            invoice.PaymentMode = "Cash";

            // ✅ FIXED: Update customer record with fresh data from QuickBooks
            var customer = invoice.Customer;
            customer.CustomerAddress = details.BuyerAddress;
            customer.CustomerPhone = details.BuyerPhone;
            customer.CustomerNTN = details.CustomerNTN;
            customer.CustomerType = details.BuyerRegistrationType;
            customer.CustomerEmail = details.BuyerEmail ?? "";
            customer.ModifiedDate = DateTime.Now;

            // Save updated customer to database
            _db.UpdateCustomer(customer);

            // Map items
            invoice.Items = details.Items?.Select(i => new InvoiceItem
            {
                ItemName = i.ItemName ?? "",
                ItemDescription = i.ItemName ?? "",
                Quantity = i.Quantity,
                UnitPrice = i.UnitPrice,
                TotalPrice = i.TotalPrice,
                NetAmount = i.NetAmount,
                TaxRate = i.TaxRate,
                SalesTaxAmount = i.SalesTaxAmount,
                TotalValue = i.TotalValue,
                HSCode = i.HSCode ?? "",
                UnitOfMeasure = i.UnitOfMeasure ?? "",
                RetailPrice = i.RetailPrice,
                ExtraTax = i.ExtraTax,
                FurtherTax = i.FurtherTax,
                FedPayable = i.FedPayable,
                SalesTaxWithheldAtSource = i.SalesTaxWithheldAtSource,
                Discount = i.Discount,
                SaleType = i.SaleType ?? "Goods at standard rate (default)",
                SroScheduleNo = i.SroScheduleNo ?? "",
                SroItemSerialNo = i.SroItemSerialNo ?? ""
            }).ToList() ?? new List<InvoiceItem>();
        }

        private void ValidateUploadPrerequisites()
        {
            if (_currentCompany == null)
                throw new InvalidOperationException("Company not set. Call SetCompany() first.");

            if (string.IsNullOrEmpty(_currentCompany.SellerNTN))
                throw new InvalidOperationException("Seller NTN not configured.");

            if (string.IsNullOrEmpty(_currentCompany.FBRToken))
                throw new InvalidOperationException("FBR Token not configured.");
        }

        private async Task GeneratePDFAsync(Invoice invoice)
        {
            try
            {
                await Task.Run(() =>
                {
                    var invoiceFromDb = _db.GetInvoiceWithDetails(
                        invoice.QuickBooksInvoiceId,
                        _currentCompany.CompanyName);

                    if (invoiceFromDb != null)
                    {
                        string outputPath = GetDefaultPDFPath(invoice);
                        _pdf.GenerateInvoicePDF(invoiceFromDb, _currentCompany, outputPath);
                        Console.WriteLine($"📄 PDF generated: {outputPath}");
                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ PDF generation failed: {ex.Message}");
            }
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

        #endregion
    }
}