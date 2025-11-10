using C2B_FBR_Connect.Models;
using QBFC16Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Services
{
    public class QuickBooksService : IDisposable
    {
        #region Fields

        private QBSessionManager _sessionManager;
        private bool _isConnected;
        private CompanyInfo _companyInfo;
        private Company _companySettings;
        private FBRApiService _fbr;
        private SroDataService _sroDataService;
        private DataCacheService _cache;
        private InvoiceTrackingService _trackingService;

        private short _qbXmlMajorVersion = 13;
        private short _qbXmlMinorVersion = 0;

        public string CurrentCompanyName { get; private set; }
        public string CurrentCompanyFile { get; private set; }

        #endregion

        #region Constructor

        public QuickBooksService()
        {
            _cache = new DataCacheService();
        }

        #endregion

        #region Connection Management

        public bool Connect(Company companySettings = null)
        {
            try
            {
                _companySettings = companySettings;

                _sessionManager = new QBSessionManager();
                _sessionManager.OpenConnection("", "C2B Smart App");
                _sessionManager.BeginSession("", ENOpenMode.omDontCare);

                DetectQBXMLVersion();
                FetchCompanyInfo();

                _isConnected = true;
                return true;
            }
            catch (Exception ex)
            {
                _isConnected = false;
                LogBoth($"QuickBooks connection error: {ex.Message}");
                throw new Exception($"QuickBooks connection failed: {ex.Message}", ex);
            }
        }

        private void DetectQBXMLVersion()
        {
            try
            {
                string[] supportedVersions = (string[])_sessionManager.QBXMLVersionsForSession;

                if (supportedVersions == null || supportedVersions.Length == 0)
                {
                    _qbXmlMajorVersion = 13;
                    _qbXmlMinorVersion = 0;
                    return;
                }

                string highestVersion = supportedVersions[supportedVersions.Length - 1];
                string[] versionParts = highestVersion.Split('.');

                _qbXmlMajorVersion = short.Parse(versionParts[0]);
                _qbXmlMinorVersion = versionParts.Length > 1 ? short.Parse(versionParts[1]) : (short)0;

                LogBoth($"✅ Using QBXML version: {_qbXmlMajorVersion}.{_qbXmlMinorVersion}");
            }
            catch (Exception ex)
            {
                LogBoth($"Error detecting QBXML version: {ex.Message}");
                _qbXmlMajorVersion = 13;
                _qbXmlMinorVersion = 0;
            }
        }

        private void FetchCompanyInfo()
        {
            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var companyQuery = msgSetRq.AppendCompanyQueryRq();
                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode == 0)
                {
                    var company = (ICompanyRet)response.Detail;

                    _companyInfo = new CompanyInfo
                    {
                        Name = company.CompanyName?.GetValue(),
                        Address = FormatAddress(company.LegalAddress),
                        City = company.LegalAddress?.City?.GetValue(),
                        State = company.LegalAddress?.State?.GetValue(),
                        PostalCode = company.LegalAddress?.PostalCode?.GetValue(),
                        Country = company.LegalAddress?.Country?.GetValue(),
                        Phone = company.Phone?.GetValue(),
                        Email = company.Email?.GetValue()
                    };

                    CurrentCompanyName = _companyInfo.Name;
                    CurrentCompanyFile = _sessionManager.GetCurrentCompanyFileName();

                    LogBoth($"✅ Connected: {CurrentCompanyName}");
                }
            }
            catch (Exception ex)
            {
                LogBoth($"Warning: Could not fetch company info: {ex.Message}");
            }
        }

        public CompanyInfo GetCompanyInfo() => _companyInfo;

        #endregion

        #region Cache Management

        /// <summary>
        /// Clear cache - call this after modifying data in QuickBooks
        /// </summary>
        public void RefreshCache()
        {
            _cache.ClearAll();
            LogBoth("🔄 Cache refreshed - all data will be reloaded from QuickBooks");
        }

        /// <summary>
        /// Get cache statistics for debugging
        /// </summary>
        public string GetCacheStats()
        {
            return _cache.GetCacheStats();
        }

        /// <summary>
        /// Pre-load all customers and items for an invoice (batch optimization)
        /// </summary>
        private void PreloadInvoiceData(IInvoiceRet inv)
        {
            var customerListIDs = new HashSet<string>();
            var itemListIDs = new HashSet<string>();

            // Collect invoice customer
            string invoiceCustomerID = inv.CustomerRef?.ListID?.GetValue();
            if (!string.IsNullOrEmpty(invoiceCustomerID))
            {
                customerListIDs.Add(invoiceCustomerID);
            }

            // Collect all line item IDs
            if (inv.ORInvoiceLineRetList != null)
            {
                for (int i = 0; i < inv.ORInvoiceLineRetList.Count; i++)
                {
                    var lineRet = inv.ORInvoiceLineRetList.GetAt(i);
                    if (lineRet.InvoiceLineRet != null)
                    {
                        string itemListID = lineRet.InvoiceLineRet.ItemRef?.ListID?.GetValue();
                        if (!string.IsNullOrEmpty(itemListID))
                        {
                            itemListIDs.Add(itemListID);
                        }
                    }
                }
            }

            LogBoth($"📊 Invoice has {customerListIDs.Count} customers and {itemListIDs.Count} unique items");

            // Batch load customers not in cache
            var customersToFetch = customerListIDs.Where(id => !_cache.TryGetCustomer(id, out _)).ToList();
            if (customersToFetch.Count > 0)
            {
                LogBoth($"🔍 Batch loading {customersToFetch.Count} customers...");
                BatchFetchCustomers(customersToFetch);
            }

            // Batch load items not in cache
            var itemsToFetch = itemListIDs.Where(id => !_cache.TryGetItem(id, out _)).ToList();
            if (itemsToFetch.Count > 0)
            {
                LogBoth($"🔍 Batch loading {itemsToFetch.Count} items...");
                BatchFetchItems(itemsToFetch);
            }
        }

        #endregion

        #region Invoice Fetching

        public List<Invoice> FetchInvoices(bool excludeUploaded = true)
        {
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

                // ✅ Initialize tracking service if not already done
                if (_trackingService == null)
                {
                    _trackingService = new InvoiceTrackingService(CurrentCompanyName);
                }

                var invoices = new List<Invoice>();
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var invoiceQuery = msgSetRq.AppendInvoiceQueryRq();
                invoiceQuery.IncludeLineItems.SetValue(true);

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode == 0)
                {
                    var invoiceRetList = response.Detail as IInvoiceRetList;
                    if (invoiceRetList != null)
                    {
                        // Collect all unique customer IDs
                        var customerIDs = new HashSet<string>();
                        for (int i = 0; i < invoiceRetList.Count; i++)
                        {
                            var inv = invoiceRetList.GetAt(i);
                            string customerListID = inv.CustomerRef?.ListID?.GetValue();
                            if (!string.IsNullOrEmpty(customerListID))
                            {
                                customerIDs.Add(customerListID);
                            }
                        }

                        // Batch load all customers
                        var customersToFetch = customerIDs.Where(id => !_cache.TryGetCustomer(id, out _)).ToList();
                        if (customersToFetch.Count > 0)
                        {
                            LogBoth($"🔍 Batch loading {customersToFetch.Count} customers for invoice list...");
                            BatchFetchCustomers(customersToFetch);
                        }

                        // Build invoice list
                        int skippedCount = 0;
                        for (int i = 0; i < invoiceRetList.Count; i++)
                        {
                            var inv = invoiceRetList.GetAt(i);
                            string qbInvoiceId = inv.TxnID.GetValue();

                            // ✅ CHECK IF ALREADY UPLOADED
                            if (excludeUploaded && _trackingService.IsInvoiceUploaded(qbInvoiceId))
                            {
                                skippedCount++;
                                continue; // Skip this invoice
                            }

                            string customerListID = inv.CustomerRef?.ListID?.GetValue() ?? "";
                            var customerData = FetchCustomerDetails(customerListID);

                            var uploadRecord = _trackingService.GetUploadStatus(qbInvoiceId);

                            invoices.Add(new Invoice
                            {
                                CompanyName = CurrentCompanyName,
                                QuickBooksInvoiceId = qbInvoiceId,
                                InvoiceNumber = inv.RefNumber?.GetValue() ?? "",
                                CustomerName = inv.CustomerRef?.FullName?.GetValue() ?? "",
                                CustomerNTN = customerData.NTN,
                                Amount = Convert.ToDecimal(inv.Subtotal?.GetValue() ?? 0),
                                Status = uploadRecord?.Status == UploadStatus.Success ? "Uploaded" :
                                        uploadRecord?.Status == UploadStatus.Failed ? "Failed" : "Pending",
                                CreatedDate = inv.TxnDate?.GetValue() ?? DateTime.Now,
                                FBR_IRN = uploadRecord?.IRN,
                                UploadDate = uploadRecord?.UploadDate
                            });
                        }

                        if (excludeUploaded && skippedCount > 0)
                        {
                            LogBoth($"✅ Skipped {skippedCount} already-uploaded invoices (saved {skippedCount} QB queries)");
                        }
                    }
                }

                return invoices;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to fetch invoices: {ex.Message}", ex);
            }
            finally
            {
                if (!wasConnected && _isConnected)
                {
                    CloseSession();
                    CloseConnection();
                }
            }
        }

        /// <summary>
        /// Get invoice details only if not already uploaded
        /// </summary>
        public async Task<FBRInvoicePayload> GetInvoiceDetails(string qbInvoiceId, bool checkUploadStatus = true)
        {
            _fbr = new FBRApiService();
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

                // ✅ Initialize tracking service if not already done
                if (_trackingService == null)
                {
                    _trackingService = new InvoiceTrackingService(CurrentCompanyName);
                }

                // ✅ CHECK IF ALREADY UPLOADED
                if (checkUploadStatus && _trackingService.IsInvoiceUploaded(qbInvoiceId))
                {
                    LogBoth($"⏭️ Invoice {qbInvoiceId} already uploaded - skipping QB query");
                    return null; // Or return cached data if needed
                }

                // ... rest of existing GetInvoiceDetails code ...
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var invoiceQuery = msgSetRq.AppendInvoiceQueryRq();
                invoiceQuery.ORInvoiceQuery.TxnIDList.Add(qbInvoiceId);
                invoiceQuery.IncludeLineItems.SetValue(true);

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode != 0 || response.Detail == null)
                    return null;

                var invoiceRetList = response.Detail as IInvoiceRetList;
                if (invoiceRetList == null || invoiceRetList.Count == 0)
                    return null;

                var inv = invoiceRetList.GetAt(0);

                // PRE-LOAD ALL DATA IN BATCH
                PreloadInvoiceData(inv);

                string customerListID = inv.CustomerRef?.ListID?.GetValue() ?? "";
                var customerData = FetchCustomerDetails(customerListID);

                var payload = new FBRInvoicePayload
                {
                    InvoiceType = "Sale Invoice",
                    InvoiceNumber = inv.RefNumber?.GetValue() ?? "",
                    InvoiceDate = inv.TxnDate?.GetValue() ?? DateTime.Now,
                    SellerNTN = _companySettings?.SellerNTN ?? "",
                    SellerBusinessName = _companyInfo?.Name ?? CurrentCompanyName,
                    SellerProvince = _companySettings?.SellerProvince ?? "",
                    SellerAddress = _companySettings?.SellerAddress ?? _companyInfo?.Address ?? "",
                    CustomerName = inv.CustomerRef?.FullName?.GetValue() ?? "",
                    CustomerNTN = customerData.NTN,
                    BuyerProvince = customerData.State,
                    BuyerAddress = customerData.Address,
                    BuyerRegistrationType = customerData.CustomerType,
                    BuyerPhone = customerData.Phone,
                    TotalAmount = Convert.ToDecimal((inv.Subtotal?.GetValue() ?? 0) + (inv.SalesTaxTotal?.GetValue() ?? 0)),
                    Subtotal = Convert.ToDecimal(inv.Subtotal?.GetValue() ?? 0),
                    TaxAmount = Convert.ToDecimal(inv.SalesTaxTotal?.GetValue() ?? 0),
                    Items = new List<InvoiceItem>()
                };

                double taxRate = inv.SalesTaxPercentage?.GetValue() ?? 0;

                // Process invoice lines
                var lineItemsWithContext = ProcessInvoiceLines(inv);
                ApplyDiscountsContextually(lineItemsWithContext);

                // Build invoice items
                await BuildInvoiceItems(lineItemsWithContext, payload, taxRate);

                // Enrich with SRO data if available
                await EnrichWithSroData(payload);

                return payload;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to get invoice details: {ex.Message}", ex);
            }
            finally
            {
                // Log cache statistics
                LogBoth(_cache.GetCacheStats());

                if (!wasConnected && _isConnected)
                {
                    CloseSession();
                    CloseConnection();
                }
            }
        }

        #endregion

        #region Invoice Line Processing

        private List<LineItemContext> ProcessInvoiceLines(IInvoiceRet inv)
        {
            var lineItems = new List<LineItemContext>();

            if (inv.ORInvoiceLineRetList == null)
                return lineItems;

            for (int i = 0; i < inv.ORInvoiceLineRetList.Count; i++)
            {
                var lineRet = inv.ORInvoiceLineRetList.GetAt(i);
                if (lineRet.InvoiceLineRet == null) continue;

                var line = lineRet.InvoiceLineRet;
                string itemListID = line.ItemRef?.ListID?.GetValue() ?? "";
                double lineAmount = line.Amount?.GetValue() ?? 0;
                string itemType = line.ItemRef?.FullName?.GetValue() ?? "";
                string itemName = line.Desc?.GetValue() ?? "";

                bool isSubtotal = IsLikelySubtotal(line, itemListID, itemType, itemName, lineAmount);
                bool isDiscount = lineAmount < 0 || itemType.Contains("Discount", StringComparison.OrdinalIgnoreCase);

                lineItems.Add(new LineItemContext
                {
                    Index = i,
                    Line = line,
                    ItemListID = itemListID,
                    Amount = lineAmount,
                    ItemType = itemType,
                    ItemName = itemName,
                    IsDiscount = isDiscount,
                    IsSubtotal = isSubtotal
                });
            }

            return lineItems;
        }

        private async Task BuildInvoiceItems(List<LineItemContext> lineItems, FBRInvoicePayload payload, double taxRate)
        {
            LogBoth($"\n🔨 BuildInvoiceItems: Processing {lineItems.Count} line items");
            LogBoth($"   Tax Rate: {taxRate}%");

            foreach (var context in lineItems)
            {
                // Skip non-item lines
                if (context.IsDiscount || context.IsSubtotal)
                    continue;

                if (string.IsNullOrEmpty(context.ItemListID) && string.IsNullOrEmpty(context.ItemType))
                    continue;

                var line = context.Line;

                // Fetch item details from QuickBooks (uses cache)
                var itemData = FetchItemDetails(context.ItemListID);

                // Extract item properties (with line-level overrides)
                string hsCode = itemData.HSCode;
                string retailPrice = itemData.RetailPrice;
                string saleType = itemData.SaleType;

                // Check for line-level custom field overrides
                if (line.DataExtRetList != null && line.DataExtRetList.Count > 0)
                {
                    for (int k = 0; k < line.DataExtRetList.Count; k++)
                    {
                        var dataExt = line.DataExtRetList.GetAt(k);
                        string fieldName = dataExt.DataExtName?.GetValue();
                        string fieldValue = dataExt.DataExtValue?.GetValue();

                        if (fieldName?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                            hsCode = fieldValue?.Trim();
                        else if (fieldName?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                            retailPrice = fieldValue?.Trim();
                        else if (fieldName?.Equals("Sale Type", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                            saleType = fieldValue?.Trim();
                    }
                }

                // Determine item classification
                bool is3rdSchedule = saleType?.Contains("3rd Schedule", StringComparison.OrdinalIgnoreCase) == true;
                bool isStandardRate = saleType?.Equals("Goods at standard rate (default)", StringComparison.OrdinalIgnoreCase) == true ||
                                     saleType?.Equals("Goods at Standard Rate (default)", StringComparison.OrdinalIgnoreCase) == true;
                bool isExempt = saleType?.Contains("Exempt", StringComparison.OrdinalIgnoreCase) == true;
                bool isZeroRated = saleType?.Contains("zero-rate", StringComparison.OrdinalIgnoreCase) == true;
                bool isSRO297 = saleType?.Contains("SRO.297", StringComparison.OrdinalIgnoreCase) == true ||
                                saleType?.Contains("SRO 297", StringComparison.OrdinalIgnoreCase) == true;
                bool isMobilePhones = saleType?.Contains("Mobile Phone", StringComparison.OrdinalIgnoreCase) == true;
                bool isPotassiumChlorate = saleType?.Contains("Potassium Chlorate", StringComparison.OrdinalIgnoreCase) == true;
                bool isCNGSales = saleType?.Contains("CNG Sales", StringComparison.OrdinalIgnoreCase) == true ||
                                  saleType?.Equals("CNG", StringComparison.OrdinalIgnoreCase) == true;

                // Determine effective tax rate
                double effectiveTaxRate = taxRate;
                if (isExempt || isZeroRated)
                {
                    effectiveTaxRate = 0;
                }

                // Calculate base amounts
                int quantity = Convert.ToInt32(line.Quantity?.GetValue() ?? 1);
                decimal lineAmount = Convert.ToDecimal(context.Amount);
                decimal discount = context.ApplicableDiscount;
                decimal netAmount = lineAmount - discount;

                // Validate net amount
                if (netAmount <= 0)
                {
                    netAmount = Math.Abs(lineAmount);
                }

                // Parse retail price
                decimal parsedRetailPrice = 0m;
                if (decimal.TryParse(retailPrice, out decimal unitRetailPrice))
                    parsedRetailPrice = unitRetailPrice * quantity;

                // Calculate tax based on item type
                decimal displayTaxRate, salesTaxAmount, furtherTax, computedTotalValue;
                string rateString;

                if (isExempt)
                {
                    CalculateExemptTax(netAmount, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue);
                }
                else if (isZeroRated)
                {
                    CalculateZeroRatedTax(netAmount, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue);
                }
                else if (isCNGSales)
                {
                    CalculateCNGSalesTax(netAmount, quantity, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue, out parsedRetailPrice);
                }
                else if (isPotassiumChlorate)
                {
                    CalculatePotassiumChlorateTax(netAmount, quantity, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue, out parsedRetailPrice);
                }
                else if (is3rdSchedule && parsedRetailPrice > 0)
                {
                    Calculate3rdScheduleTax(parsedRetailPrice, effectiveTaxRate, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue, out netAmount);
                }
                else if (isSRO297)
                {
                    CalculateSRO297Tax(netAmount, effectiveTaxRate, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue);
                }
                else if (isMobilePhones)
                {
                    CalculateMobilePhonesTax(netAmount, effectiveTaxRate, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue);
                }
                else if (isStandardRate && effectiveTaxRate > 18)
                {
                    CalculateStandardRateSplitTax(netAmount, effectiveTaxRate, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue);
                }
                else
                {
                    CalculateStandardTax(netAmount, effectiveTaxRate, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue);
                }

                // Create invoice item
                string itemName = line.Desc?.GetValue() ?? "";
                decimal unitPrice = quantity > 0 ? lineAmount / quantity : lineAmount;

                var item = new InvoiceItem
                {
                    ItemName = itemName,
                    ItemDescription = itemName,
                    HSCode = hsCode ?? "",
                    Quantity = quantity,
                    UnitOfMeasure = await GetUnitOfMeasure(hsCode, line.UnitOfMeasure?.GetValue()),
                    UnitPrice = unitPrice,
                    TotalPrice = lineAmount,
                    NetAmount = netAmount,
                    TaxRate = displayTaxRate,
                    Rate = rateString,
                    SalesTaxAmount = salesTaxAmount,
                    TotalValue = computedTotalValue,
                    RetailPrice = parsedRetailPrice,
                    ExtraTax = 0m,
                    FurtherTax = furtherTax,
                    FedPayable = 0m,
                    SalesTaxWithheldAtSource = 0m,
                    Discount = discount,
                    SaleType = saleType ?? "Goods at standard rate (default)",
                    SroScheduleNo = "",
                    SroItemSerialNo = ""
                };

                // Set default SRO values
                SetDefaultSroValues(item);

                payload.Items.Add(item);
            }

            LogBoth($"\n✅ BuildInvoiceItems Complete: Added {payload.Items.Count} items to payload\n");
        }

        #region Tax Calculation Methods

        private void CalculateExemptTax(decimal netAmount, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue)
        {
            displayTaxRate = 0m;
            salesTaxAmount = 0m;
            furtherTax = 0m;
            rateString = "Exempt";
            computedTotalValue = netAmount;
        }

        private void CalculateZeroRatedTax(decimal netAmount, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue)
        {
            displayTaxRate = 0m;
            salesTaxAmount = 0m;
            furtherTax = 0m;
            rateString = "0%";
            computedTotalValue = netAmount;
        }

        private void CalculateCNGSalesTax(decimal netAmount, int quantity, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue, out decimal parsedRetailPrice)
        {
            const decimal fixedRatePerUnit = 200m;
            displayTaxRate = 0m;
            salesTaxAmount = fixedRatePerUnit * quantity;
            furtherTax = 0m;
            rateString = "Rs.200";
            parsedRetailPrice = fixedRatePerUnit * quantity;
            computedTotalValue = netAmount + salesTaxAmount;
        }

        private void CalculatePotassiumChlorateTax(decimal netAmount, int quantity, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue, out decimal parsedRetailPrice)
        {
            const decimal fixedRatePerKg = 60m;
            displayTaxRate = 18m;

            decimal percentageTax = Convert.ToDecimal((double)netAmount * 0.18);
            decimal fixedTaxAmount = fixedRatePerKg * quantity;

            salesTaxAmount = percentageTax + fixedTaxAmount;
            furtherTax = 0m;
            rateString = "18% along with rupees 60 per kilogram";
            parsedRetailPrice = fixedRatePerKg * quantity;
            computedTotalValue = netAmount + salesTaxAmount;
        }

        private void Calculate3rdScheduleTax(decimal retailPrice, double effectiveTaxRate, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue, out decimal netAmount)
        {
            displayTaxRate = Convert.ToDecimal(effectiveTaxRate);
            salesTaxAmount = Convert.ToDecimal((double)retailPrice * (effectiveTaxRate / 100.0));
            furtherTax = 0m;
            rateString = $"{effectiveTaxRate}%";
            netAmount = 0m;
            computedTotalValue = retailPrice + salesTaxAmount;
        }

        private void CalculateSRO297Tax(decimal netAmount, double effectiveTaxRate, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue)
        {
            displayTaxRate = Convert.ToDecimal(effectiveTaxRate);
            salesTaxAmount = Convert.ToDecimal((double)netAmount * (effectiveTaxRate / 100.0));
            furtherTax = 0m;
            rateString = $"{effectiveTaxRate}%";
            computedTotalValue = netAmount + salesTaxAmount;
        }

        private void CalculateMobilePhonesTax(decimal netAmount, double effectiveTaxRate, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue)
        {
            displayTaxRate = Convert.ToDecimal(effectiveTaxRate);
            salesTaxAmount = Convert.ToDecimal((double)netAmount * (effectiveTaxRate / 100.0));
            furtherTax = 0m;
            rateString = $"{effectiveTaxRate}%";
            computedTotalValue = netAmount + salesTaxAmount;
        }

        private void CalculateStandardRateSplitTax(decimal netAmount, double effectiveTaxRate, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue)
        {
            displayTaxRate = 18m;
            salesTaxAmount = Convert.ToDecimal((double)netAmount * 0.18);
            decimal totalTaxAmount = Convert.ToDecimal((double)netAmount * (effectiveTaxRate / 100.0));
            furtherTax = totalTaxAmount - salesTaxAmount;
            rateString = "18%";
            computedTotalValue = netAmount + salesTaxAmount + furtherTax;
        }

        private void CalculateStandardTax(decimal netAmount, double effectiveTaxRate, out decimal displayTaxRate, out decimal salesTaxAmount,
            out decimal furtherTax, out string rateString, out decimal computedTotalValue)
        {
            displayTaxRate = Convert.ToDecimal(effectiveTaxRate);
            decimal totalTaxAmount = Convert.ToDecimal((double)netAmount * (effectiveTaxRate / 100.0));
            rateString = $"{effectiveTaxRate}%";

            if (effectiveTaxRate > 18)
            {
                salesTaxAmount = Convert.ToDecimal((double)netAmount * 0.18);
                furtherTax = totalTaxAmount - salesTaxAmount;
            }
            else
            {
                salesTaxAmount = totalTaxAmount;
                furtherTax = 0m;
            }

            computedTotalValue = netAmount + salesTaxAmount + furtherTax;
        }

        #endregion

        private async Task<string> GetUnitOfMeasure(string hsCode, string defaultUom)
        {
            if (_fbr != null && _companySettings != null && !string.IsNullOrEmpty(_companySettings.FBRToken))
            {
                var uom = await _fbr.GetUOMDescriptionAsync(hsCode, _companySettings.FBRToken);
                if (!string.IsNullOrEmpty(uom))
                    return uom;
            }
            return defaultUom;
        }

        private async Task EnrichWithSroData(FBRInvoicePayload payload)
        {
            if (_sroDataService == null || _companySettings == null || string.IsNullOrEmpty(_companySettings.FBRToken))
                return;

            int provinceCode = GetProvinceCode(_companySettings.SellerProvince);
            if (provinceCode <= 0) return;

            try
            {
                await _sroDataService.EnrichInvoiceItemsWithSroDataAsync(
                    payload.Items,
                    provinceCode,
                    payload.InvoiceDate,
                    _companySettings.FBRToken
                );
            }
            catch (Exception ex)
            {
                LogBoth($"⚠️ SRO enrichment failed: {ex.Message}");
            }
        }

        private void SetDefaultSroValues(InvoiceItem item)
        {
            if (string.IsNullOrEmpty(item.SaleType))
                return;

            string normalizedSaleType = ScenarioMapper.NormalizeSaleType(item.SaleType);

            switch (normalizedSaleType)
            {
                case "Electricity Supply to Retailers":
                    item.SroScheduleNo = "1450(I)/2021";
                    item.SroItemSerialNo = "4";
                    LogBoth($"      ✅ Set default SRO for Electricity: Schedule={item.SroScheduleNo}, Serial={item.SroItemSerialNo}");
                    break;
            }
        }

        #endregion

        #region Discount Processing

        private void ApplyDiscountsContextually(List<LineItemContext> lineItems)
        {
            var subtotalIndices = lineItems
                .Select((item, index) => new { item, index })
                .Where(x => x.item.IsSubtotal)
                .Select(x => x.index)
                .ToList();

            foreach (var discount in lineItems.Where(x => x.IsDiscount))
            {
                decimal discountAmount = Math.Abs(Convert.ToDecimal(discount.Amount));
                int discountIndex = discount.Index;

                int relevantSubtotalIndex = subtotalIndices
                    .Where(idx => idx < discountIndex)
                    .OrderByDescending(idx => idx)
                    .FirstOrDefault(-1);

                if (relevantSubtotalIndex >= 0)
                {
                    ApplyDiscountToSubtotalRange(lineItems, relevantSubtotalIndex, discountIndex, discountAmount);
                }
                else
                {
                    ApplyDiscountToPreviousItem(lineItems, discountIndex, discountAmount);
                }
            }
        }

        private void ApplyDiscountToSubtotalRange(List<LineItemContext> lineItems, int subtotalIndex, int discountIndex, decimal discountAmount)
        {
            var itemsToDiscount = lineItems
                .Skip(subtotalIndex + 1)
                .Take(discountIndex - subtotalIndex - 1)
                .Where(x => !x.IsDiscount && !x.IsSubtotal)
                .ToList();

            if (itemsToDiscount.Count == 0)
            {
                int rangeStart = 0;
                for (int j = subtotalIndex - 1; j >= 0; j--)
                {
                    if (lineItems[j].IsSubtotal)
                    {
                        rangeStart = j + 1;
                        break;
                    }
                }

                itemsToDiscount = lineItems
                    .Skip(rangeStart)
                    .Take(subtotalIndex - rangeStart)
                    .Where(x => !x.IsDiscount && !x.IsSubtotal)
                    .ToList();
            }

            DistributeDiscountToItems(itemsToDiscount, discountAmount);
        }

        private void ApplyDiscountToPreviousItem(List<LineItemContext> lineItems, int discountIndex, decimal discountAmount)
        {
            var previousItem = lineItems
                .Take(discountIndex)
                .Where(x => !x.IsDiscount && !x.IsSubtotal)
                .LastOrDefault();

            if (previousItem != null)
            {
                decimal itemAmount = Math.Abs(Convert.ToDecimal(previousItem.Amount));
                previousItem.ApplicableDiscount = Math.Min(
                    previousItem.ApplicableDiscount + discountAmount,
                    itemAmount
                );
            }
        }

        private void DistributeDiscountToItems(List<LineItemContext> items, decimal totalDiscount)
        {
            if (items == null || items.Count == 0 || totalDiscount <= 0)
                return;

            decimal totalAmount = items.Sum(x => Math.Abs(Convert.ToDecimal(x.Amount)));
            if (totalAmount <= 0) return;

            totalDiscount = Math.Min(totalDiscount, totalAmount);

            decimal allocated = 0m;
            for (int i = 0; i < items.Count; i++)
            {
                decimal itemAmount = Math.Abs(Convert.ToDecimal(items[i].Amount));
                decimal proportion = itemAmount / totalAmount;
                decimal itemDiscount = Math.Round(totalDiscount * proportion, 2, MidpointRounding.AwayFromZero);

                items[i].ApplicableDiscount += itemDiscount;
                allocated += itemDiscount;
            }

            decimal remainder = totalDiscount - allocated;
            if (remainder != 0m && items.Count > 0)
            {
                var largestItem = items.OrderByDescending(x => Math.Abs(Convert.ToDecimal(x.Amount))).First();
                largestItem.ApplicableDiscount += remainder;
            }
        }

        private bool IsLikelySubtotal(IInvoiceLineRet line, string itemListID, string itemType, string itemName, double lineAmount)
        {
            string type = itemType?.Trim() ?? "";
            string name = itemName?.Trim() ?? "";

            if (type.Contains("subtotal", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("subtotal", StringComparison.OrdinalIgnoreCase))
                return true;

            bool itemRefEmpty = string.IsNullOrEmpty(itemListID);
            bool quantityMissing = line.Quantity == null || Math.Abs(line.Quantity.GetValue()) == 0;

            return itemRefEmpty && quantityMissing && lineAmount > 0;
        }

        #endregion

        #region Tracking Management

        /// <summary>
        /// Mark invoice as successfully uploaded
        /// </summary>
        public void MarkInvoiceAsUploaded(string qbInvoiceId, string invoiceNumber, string irn)
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            _trackingService.MarkAsUploaded(qbInvoiceId, invoiceNumber, irn, DateTime.Now);
            LogBoth($"✅ Marked invoice {invoiceNumber} as uploaded (IRN: {irn})");
        }

        /// <summary>
        /// Mark invoice as failed
        /// </summary>
        public void MarkInvoiceAsFailed(string qbInvoiceId, string invoiceNumber, string errorMessage)
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            _trackingService.MarkAsFailed(qbInvoiceId, invoiceNumber, errorMessage);
            LogBoth($"❌ Marked invoice {invoiceNumber} as failed: {errorMessage}");
        }

        /// <summary>
        /// Get tracking statistics
        /// </summary>
        public string GetTrackingStatistics()
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            return _trackingService.GetStatistics().ToString();
        }

        /// <summary>
        /// Clear invoice tracking (use with caution!)
        /// </summary>
        public void ClearInvoiceTracking()
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            _trackingService.ClearAll();
            LogBoth("🔄 Cleared all invoice tracking data");
        }

        /// <summary>
        /// Get tracking service for advanced operations
        /// </summary>
        public InvoiceTrackingService GetTrackingService()
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            return _trackingService;
        }

        #endregion

        #region Data Fetching with Cache

        private CustomerData FetchCustomerDetails(string customerListID)
        {
            if (string.IsNullOrEmpty(customerListID))
            {
                return new CustomerData { CustomerType = "Unregistered" };
            }

            // ✅ CHECK CACHE FIRST
            if (_cache.TryGetCustomer(customerListID, out var cachedCustomer))
            {
                return cachedCustomer;
            }

            // ✅ CACHE MISS - Fetch from QuickBooks
            LogBoth($"🔍 Fetching customer from QuickBooks (ListID: {customerListID})...");

            var customerData = new CustomerData();

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var customerQuery = msgSetRq.AppendCustomerQueryRq();
                customerQuery.ORCustomerListQuery.ListIDList.Add(customerListID);
                customerQuery.OwnerIDList.Add("0");

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode == 0)
                {
                    var customerList = response.Detail as ICustomerRetList;
                    if (customerList != null && customerList.Count > 0)
                    {
                        var customer = customerList.GetAt(0);
                        customerData = ExtractCustomerData(customer);

                        // ✅ ADD TO CACHE
                        _cache.AddCustomer(customerListID, customerData);
                    }
                }
            }
            catch (Exception ex)
            {
                LogBoth($"   ❌ Error fetching customer: {ex.Message}");
            }

            return customerData;
        }

        private ItemData FetchItemDetails(string itemListID)
        {
            if (string.IsNullOrEmpty(itemListID))
                return new ItemData { SaleType = "Goods at standard rate (default)" };

            // ✅ CHECK CACHE FIRST
            if (_cache.TryGetItem(itemListID, out var cachedItem))
            {
                return cachedItem;
            }

            // ✅ CACHE MISS - Fetch from QuickBooks
            LogBoth($"🔍 Fetching item from QuickBooks (ListID: {itemListID})...");

            var itemData = new ItemData { SaleType = "Goods at standard rate (default)" };

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var itemQuery = msgSetRq.AppendItemQueryRq();
                itemQuery.ORListQuery.ListIDList.Add(itemListID);
                itemQuery.OwnerIDList.Add("0");

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode == 0)
                {
                    var itemList = response.Detail as IORItemRetList;
                    if (itemList != null && itemList.Count > 0)
                    {
                        var itemRet = itemList.GetAt(0);
                        itemData = ExtractItemData(itemRet, itemListID);

                        // ✅ ADD TO CACHE
                        _cache.AddItem(itemListID, itemData);
                    }
                }
            }
            catch (Exception ex)
            {
                LogBoth($"   ❌ Error: {ex.Message}");
            }

            return itemData;
        }

        #endregion

        #region Batch Fetching Methods

        private void BatchFetchCustomers(List<string> customerListIDs)
        {
            if (customerListIDs == null || customerListIDs.Count == 0)
                return;

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var customerQuery = msgSetRq.AppendCustomerQueryRq();
                foreach (var listID in customerListIDs)
                {
                    customerQuery.ORCustomerListQuery.ListIDList.Add(listID);
                }
                customerQuery.OwnerIDList.Add("0");

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode == 0)
                {
                    var customerList = response.Detail as ICustomerRetList;
                    if (customerList != null)
                    {
                        for (int i = 0; i < customerList.Count; i++)
                        {
                            var customer = customerList.GetAt(i);
                            string listID = customer.ListID?.GetValue();

                            if (string.IsNullOrEmpty(listID))
                                continue;

                            var customerData = ExtractCustomerData(customer);
                            _cache.AddCustomer(listID, customerData);
                        }

                        LogBoth($"✅ Batch loaded {customerList.Count} customers");
                    }
                }
            }
            catch (Exception ex)
            {
                LogBoth($"❌ Error batch fetching customers: {ex.Message}");
            }
        }

        private void BatchFetchItems(List<string> itemListIDs)
        {
            if (itemListIDs == null || itemListIDs.Count == 0)
                return;

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var itemQuery = msgSetRq.AppendItemQueryRq();
                foreach (var listID in itemListIDs)
                {
                    itemQuery.ORListQuery.ListIDList.Add(listID);
                }
                itemQuery.OwnerIDList.Add("0");

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode == 0)
                {
                    var itemList = response.Detail as IORItemRetList;
                    if (itemList != null)
                    {
                        for (int i = 0; i < itemList.Count; i++)
                        {
                            var itemRet = itemList.GetAt(i);
                            string listID = GetItemListID(itemRet);

                            if (string.IsNullOrEmpty(listID))
                                continue;

                            var itemData = ExtractItemData(itemRet, listID);
                            _cache.AddItem(listID, itemData);
                        }

                        LogBoth($"✅ Batch loaded {itemList.Count} items");
                    }
                }
            }
            catch (Exception ex)
            {
                LogBoth($"❌ Error batch fetching items: {ex.Message}");
            }
        }

        #endregion

        #region Data Extraction Methods

        private CustomerData ExtractCustomerData(ICustomerRet customer)
        {
            var customerData = new CustomerData
            {
                Address = FormatAddress(customer.BillAddress),
                Phone = customer.Phone?.GetValue() ?? "N/A"
            };

            string customerTypeFromQB = customer.CustomerTypeRef?.FullName?.GetValue() ?? "";

            if (customer.DataExtRetList != null)
            {
                for (int i = 0; i < customer.DataExtRetList.Count; i++)
                {
                    var dataExt = customer.DataExtRetList.GetAt(i);
                    string fieldName = dataExt.DataExtName?.GetValue();
                    string fieldValue = dataExt.DataExtValue?.GetValue();

                    if (IsNTNField(fieldName) && !string.IsNullOrEmpty(fieldValue))
                    {
                        customerData.NTN = fieldValue.Trim();
                    }

                    if (fieldName?.Equals("Province", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                    {
                        customerData.State = fieldValue.Trim();
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(customerTypeFromQB))
            {
                customerData.CustomerType = customerTypeFromQB.Trim();
            }
            else if (!string.IsNullOrWhiteSpace(customerData.NTN) && customerData.NTN.Length >= 7)
            {
                customerData.CustomerType = "Registered";
            }
            else
            {
                customerData.CustomerType = "Unregistered";
            }

            return customerData;
        }

        private ItemData ExtractItemData(IORItemRet itemRet, string itemListID)
        {
            var itemData = new ItemData
            {
                SaleType = "Goods at standard rate (default)"
            };

            IDataExtRetList customFields = GetCustomFieldsFromItem(itemRet);

            if (customFields != null)
            {
                for (int i = 0; i < customFields.Count; i++)
                {
                    var dataExt = customFields.GetAt(i);
                    string fieldName = dataExt.DataExtName?.GetValue();
                    string fieldValue = dataExt.DataExtValue?.GetValue();

                    if (fieldName?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        itemData.HSCode = fieldValue ?? "";
                    }
                    else if (fieldName?.Equals("Sale Type", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        if (!string.IsNullOrWhiteSpace(fieldValue))
                        {
                            itemData.SaleType = fieldValue.Trim();
                        }
                    }
                    else if (fieldName?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        itemData.RetailPrice = fieldValue ?? "";
                    }
                }
            }

            if (string.IsNullOrEmpty(itemData.RetailPrice))
            {
                itemData.RetailPrice = FetchItemRetailPriceFromPriceLists(itemListID);
            }

            return itemData;
        }

        private IDataExtRetList GetCustomFieldsFromItem(IORItemRet itemRet)
        {
            if (itemRet.ItemServiceRet != null)
                return itemRet.ItemServiceRet.DataExtRetList;
            if (itemRet.ItemInventoryRet != null)
                return itemRet.ItemInventoryRet.DataExtRetList;
            if (itemRet.ItemNonInventoryRet != null)
                return itemRet.ItemNonInventoryRet.DataExtRetList;
            if (itemRet.ItemOtherChargeRet != null)
                return itemRet.ItemOtherChargeRet.DataExtRetList;
            if (itemRet.ItemInventoryAssemblyRet != null)
                return itemRet.ItemInventoryAssemblyRet.DataExtRetList;
            if (itemRet.ItemGroupRet != null)
                return itemRet.ItemGroupRet.DataExtRetList;
            if (itemRet.ItemDiscountRet != null)
                return itemRet.ItemDiscountRet.DataExtRetList;
            if (itemRet.ItemPaymentRet != null)
                return itemRet.ItemPaymentRet.DataExtRetList;
            if (itemRet.ItemSalesTaxRet != null)
                return itemRet.ItemSalesTaxRet.DataExtRetList;
            if (itemRet.ItemSalesTaxGroupRet != null)
                return itemRet.ItemSalesTaxGroupRet.DataExtRetList;
            if (itemRet.ItemSubtotalRet != null)
                return itemRet.ItemSubtotalRet.DataExtRetList;
            if (itemRet.ItemFixedAssetRet != null)
                return itemRet.ItemFixedAssetRet.DataExtRetList;

            return null;
        }

        private string GetItemListID(IORItemRet itemRet)
        {
            if (itemRet.ItemServiceRet != null)
                return itemRet.ItemServiceRet.ListID?.GetValue();
            if (itemRet.ItemInventoryRet != null)
                return itemRet.ItemInventoryRet.ListID?.GetValue();
            if (itemRet.ItemNonInventoryRet != null)
                return itemRet.ItemNonInventoryRet.ListID?.GetValue();
            if (itemRet.ItemOtherChargeRet != null)
                return itemRet.ItemOtherChargeRet.ListID?.GetValue();
            if (itemRet.ItemInventoryAssemblyRet != null)
                return itemRet.ItemInventoryAssemblyRet.ListID?.GetValue();
            if (itemRet.ItemGroupRet != null)
                return itemRet.ItemGroupRet.ListID?.GetValue();
            if (itemRet.ItemDiscountRet != null)
                return itemRet.ItemDiscountRet.ListID?.GetValue();
            if (itemRet.ItemPaymentRet != null)
                return itemRet.ItemPaymentRet.ListID?.GetValue();
            if (itemRet.ItemSalesTaxRet != null)
                return itemRet.ItemSalesTaxRet.ListID?.GetValue();
            if (itemRet.ItemSalesTaxGroupRet != null)
                return itemRet.ItemSalesTaxGroupRet.ListID?.GetValue();
            if (itemRet.ItemSubtotalRet != null)
                return itemRet.ItemSubtotalRet.ListID?.GetValue();
            if (itemRet.ItemFixedAssetRet != null)
                return itemRet.ItemFixedAssetRet.ListID?.GetValue();

            return null;
        }

        private string FetchItemRetailPriceFromPriceLists(string itemListID)
        {
            try
            {
                // Check cache first
                if (!_cache.TryGetPriceLevels(out var priceLevels))
                {
                    priceLevels = FetchPriceLevels();
                    _cache.SetPriceLevels(priceLevels);
                }

                var retailPriceLevel = priceLevels.FirstOrDefault(p =>
                    p.Name?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true);

                if (retailPriceLevel?.Items != null)
                {
                    var itemPrice = retailPriceLevel.Items.FirstOrDefault(i => i.ItemListID == itemListID);
                    if (itemPrice != null && itemPrice.CustomPrice > 0)
                        return itemPrice.CustomPrice.ToString("0.##");
                }
            }
            catch (Exception ex)
            {
                LogBoth($"Error fetching retail price: {ex.Message}");
            }

            return "";
        }

        public List<PriceLevel> FetchPriceLevels()
        {
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

                var priceLevels = new List<PriceLevel>();
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                msgSetRq.AppendPriceLevelQueryRq();
                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode != 0 || response.Detail == null)
                    return priceLevels;

                var priceLevelList = response.Detail as IPriceLevelRetList;
                if (priceLevelList == null) return priceLevels;

                for (int i = 0; i < priceLevelList.Count; i++)
                {
                    var priceLevel = priceLevelList.GetAt(i);
                    var level = new PriceLevel
                    {
                        ListID = priceLevel.ListID?.GetValue(),
                        Name = priceLevel.Name?.GetValue(),
                        IsActive = priceLevel.IsActive?.GetValue() ?? true,
                        PriceLevelType = priceLevel.PriceLevelType?.GetValue().ToString() ?? "",
                        Items = new List<PriceLevelItem>()
                    };

                    ExtractPriceLevelItems(priceLevel, level);
                    priceLevels.Add(level);
                }

                return priceLevels;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to fetch price levels: {ex.Message}", ex);
            }
            finally
            {
                if (!wasConnected && _isConnected)
                {
                    CloseSession();
                    CloseConnection();
                }
            }
        }

        private void ExtractPriceLevelItems(IPriceLevelRet priceLevel, PriceLevel level)
        {
            IORPriceLevelRet orPriceLevel = priceLevel.ORPriceLevelRet;
            if (orPriceLevel == null) return;

            if (orPriceLevel.ortype == ENORPriceLevelRet.orplrPriceLevelFixedPercentage)
            {
                level.FixedPercentage = orPriceLevel.PriceLevelFixedPercentage?.GetValue() ?? 0;
            }
            else if (orPriceLevel.ortype == ENORPriceLevelRet.orplrPriceLevelPerItemRetCurrency)
            {
                var perItemCurrency = orPriceLevel.PriceLevelPerItemRetCurrency;
                var itemList = perItemCurrency?.PriceLevelPerItemRetList;

                if (itemList != null)
                {
                    for (int j = 0; j < itemList.Count; j++)
                    {
                        var itemPrice = itemList.GetAt(j);
                        var customPriceObj = itemPrice.ORORCustomPrice;

                        level.Items.Add(new PriceLevelItem
                        {
                            ItemListID = itemPrice.ItemRef?.ListID?.GetValue(),
                            ItemFullName = itemPrice.ItemRef?.FullName?.GetValue(),
                            CustomPrice = customPriceObj?.CustomPrice != null ?
                                Convert.ToDecimal(customPriceObj.CustomPrice.GetValue()) : 0,
                            CustomPricePercent = customPriceObj?.CustomPricePercent?.GetValue() ?? 0
                        });
                    }
                }
            }
        }

        #endregion

        #region Helper Methods

        private string FormatAddress(IAddress address)
        {
            if (address == null) return "";

            var lines = new List<string>();
            AddAddressLine(lines, address.Addr1);
            AddAddressLine(lines, address.Addr2);
            AddAddressLine(lines, address.Addr3);
            AddAddressLine(lines, address.Addr4);
            AddAddressLine(lines, address.Addr5);
            AddAddressLine(lines, address.City);
            AddAddressLine(lines, address.State);
            AddAddressLine(lines, address.PostalCode);
            AddAddressLine(lines, address.Country);

            return string.Join(", ", lines);
        }

        private void AddAddressLine(List<string> lines, IQBStringType addressField)
        {
            if (addressField != null)
            {
                var value = addressField.GetValue();
                if (!string.IsNullOrEmpty(value))
                    lines.Add(value.TrimEnd(',', ' '));
            }
        }

        private bool IsNTNField(string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName)) return false;

            return fieldName.Equals("NTN", StringComparison.OrdinalIgnoreCase) ||
                   fieldName.Equals("CNIC", StringComparison.OrdinalIgnoreCase) ||
                   fieldName.Equals("NTN/CNIC", StringComparison.OrdinalIgnoreCase);
        }

        private int GetProvinceCode(string provinceName)
        {
            if (string.IsNullOrWhiteSpace(provinceName)) return 0;

            var provinceMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "Punjab", 7 }, { "Sindh", 1 }, { "Khyber Pakhtunkhwa", 2 },
                { "Balochistan", 3 }, { "Azad Jammu and Kashmir", 4 },
                { "Islamabad Capital Territory", 5 }, { "Gilgit-Baltistan", 6 }
            };

            if (provinceMap.TryGetValue(provinceName.Trim(), out int code))
                return code;

            var n = provinceName.Trim().ToLowerInvariant();
            if (n.Contains("punjab")) return 7;
            if (n.Contains("sindh")) return 1;
            if (n.Contains("khyber") || n.Contains("kpk")) return 2;
            if (n.Contains("baloch")) return 3;
            if (n.Contains("azad")) return 4;
            if (n.Contains("islamabad")) return 5;
            if (n.Contains("gilgit")) return 6;

            return 0;
        }

        public void SetSroDataService(SroDataService sroDataService)
        {
            _sroDataService = sroDataService;
        }

        private void LogBoth(string message)
        {
            Console.WriteLine(message);
            System.Diagnostics.Debug.WriteLine(message);
        }

        #endregion

        #region Cleanup

        public void CloseSession()
        {
            if (_sessionManager != null && _isConnected)
            {
                try
                {
                    _sessionManager.EndSession();
                    _isConnected = false;
                }
                catch (Exception ex)
                {
                    LogBoth($"Error ending session: {ex.Message}");
                }
            }
        }

        public void CloseConnection()
        {
            if (_sessionManager != null)
            {
                try
                {
                    _sessionManager.CloseConnection();
                }
                catch (Exception ex)
                {
                    LogBoth($"Error closing connection: {ex.Message}");
                }
            }
        }

        public void Dispose()
        {
            try
            {
                CloseSession();
                CloseConnection();
            }
            finally
            {
                if (_sessionManager != null)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_sessionManager);
                    }
                    catch { }

                    _sessionManager = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        #endregion
    }

    #region Helper Classes

    public class CompanyInfo
    {
        public string Name { get; set; }
        public string NTN { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string PostalCode { get; set; }
        public string Country { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
    }

    public class CustomerData
    {
        public string NTN { get; set; } = "";
        public string CustomerType { get; set; } = "Unregistered";
        public string Address { get; set; } = "";
        public string State { get; set; } = "";
        public string Phone { get; set; } = "";
    }

    public class ItemData
    {
        public string HSCode { get; set; } = "";
        public string RetailPrice { get; set; } = "";
        public string SaleType { get; set; } = "Goods at standard rate (default)";
    }

    public class PriceLevel
    {
        public string ListID { get; set; }
        public string Name { get; set; }
        public bool IsActive { get; set; }
        public string PriceLevelType { get; set; }
        public List<PriceLevelItem> Items { get; set; }
        public double FixedPercentage { get; set; }
    }

    public class PriceLevelItem
    {
        public string ItemListID { get; set; }
        public string ItemFullName { get; set; }
        public decimal CustomPrice { get; set; }
        public double CustomPricePercent { get; set; }
    }

    public class LineItemContext
    {
        public int Index { get; set; }
        public IInvoiceLineRet Line { get; set; }
        public string ItemListID { get; set; }
        public double Amount { get; set; }
        public string ItemType { get; set; }
        public string ItemName { get; set; }
        public bool IsDiscount { get; set; }
        public bool IsSubtotal { get; set; }
        public decimal ApplicableDiscount { get; set; } = 0;
    }

    #endregion
}