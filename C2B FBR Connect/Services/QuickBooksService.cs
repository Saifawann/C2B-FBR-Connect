using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using iText.Layout.Element;
using Microsoft.VisualBasic.ApplicationServices;
using Org.BouncyCastle.Utilities;
using QBFC16Lib;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace C2B_FBR_Connect.Services
{
    public class QuickBooksService : IDisposable
    {
        private QBSessionManager _sessionManager;
        private bool _isConnected;
        private CompanyInfo _companyInfo;
        private Company _companySettings;
        private FBRApiService _fbr;
        private SroDataService _sroDataService;

        // Store detected QBXML version
        private short _qbXmlMajorVersion = 13;
        private short _qbXmlMinorVersion = 0;

        public string CurrentCompanyName { get; private set; }
        public string CurrentCompanyFile { get; private set; }

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
                System.Diagnostics.Debug.WriteLine($"QuickBooks connection error: {ex.Message}");
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

                System.Diagnostics.Debug.WriteLine($"✅ Using QBXML version: {_qbXmlMajorVersion}.{_qbXmlMinorVersion}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error detecting QBXML version: {ex.Message}");
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

                    System.Diagnostics.Debug.WriteLine($"✅ Connected: {CurrentCompanyName}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Warning: Could not fetch company info: {ex.Message}");
            }
        }

        public CompanyInfo GetCompanyInfo() => _companyInfo;

        #endregion

        #region Invoice Fetching

        public List<Invoice> FetchInvoices()
        {
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

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
                        for (int i = 0; i < invoiceRetList.Count; i++)
                        {
                            var inv = invoiceRetList.GetAt(i);
                            string customerListID = inv.CustomerRef?.ListID?.GetValue() ?? "";
                            var customerData = FetchCustomerDetails(customerListID);

                            invoices.Add(new Invoice
                            {
                                CompanyName = CurrentCompanyName,
                                QuickBooksInvoiceId = inv.TxnID.GetValue(),
                                InvoiceNumber = inv.RefNumber?.GetValue() ?? "",
                                CustomerName = inv.CustomerRef?.FullName?.GetValue() ?? "",
                                CustomerNTN = customerData.NTN,
                                Amount = Convert.ToDecimal(inv.Subtotal?.GetValue() ?? 0),
                                Status = "Pending",
                                CreatedDate = inv.TxnDate?.GetValue() ?? DateTime.Now
                            });
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

        public async Task<FBRInvoicePayload> GetInvoiceDetails(string qbInvoiceId)
        {
            _fbr = new FBRApiService();
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

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
            Console.WriteLine($"\n🔨 BuildInvoiceItems: Processing {lineItems.Count} line items");
            Console.WriteLine($"   Tax Rate: {taxRate}%");

            foreach (var context in lineItems)
            {
                if (context.IsDiscount || context.IsSubtotal)
                {
                    Console.WriteLine($"   ⏭️ Skipping line {context.Index}: IsDiscount={context.IsDiscount}, IsSubtotal={context.IsSubtotal}");
                    continue;
                }

                if (string.IsNullOrEmpty(context.ItemListID) && string.IsNullOrEmpty(context.ItemType))
                {
                    Console.WriteLine($"   ⏭️ Skipping line {context.Index}: No ItemListID or ItemType");
                    continue;
                }

                var line = context.Line;

                Console.WriteLine($"\n   📦 Processing Line {context.Index}:");
                Console.WriteLine($"      Description: '{line.Desc?.GetValue() ?? "[NULL]"}'");
                Console.WriteLine($"      ItemListID: '{context.ItemListID}'");
                Console.WriteLine($"      Amount: {context.Amount}");

                Console.WriteLine($"      🔍 Fetching item details from QuickBooks...");
                var itemData = FetchItemDetails(context.ItemListID);

                Console.WriteLine($"      📋 Item Data Retrieved:");
                Console.WriteLine($"         HS Code: '{itemData.HSCode}'");
                Console.WriteLine($"         Sale Type: '{itemData.SaleType}'");
                Console.WriteLine($"         Retail Price: '{itemData.RetailPrice}'");

                string hsCode = itemData.HSCode;
                string retailPrice = itemData.RetailPrice;
                string saleType = itemData.SaleType;

                if (line.DataExtRetList != null && line.DataExtRetList.Count > 0)
                {
                    Console.WriteLine($"      🔍 Checking {line.DataExtRetList.Count} line-level custom fields...");
                    for (int k = 0; k < line.DataExtRetList.Count; k++)
                    {
                        var dataExt = line.DataExtRetList.GetAt(k);
                        string fieldName = dataExt.DataExtName?.GetValue();
                        string fieldValue = dataExt.DataExtValue?.GetValue();

                        Console.WriteLine($"         Line Field: '{fieldName}' = '{fieldValue}'");

                        if (fieldName?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                            hsCode = fieldValue?.Trim();
                        else if (fieldName?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                            retailPrice = fieldValue?.Trim();
                        else if (fieldName?.Equals("Sale Type", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                            saleType = fieldValue?.Trim();
                    }
                }

                bool is3rdSchedule = saleType?.Trim().Contains("3rd Schedule", StringComparison.OrdinalIgnoreCase) == true;
                bool isStandardRate = saleType?.Trim().Equals("Goods at standard rate (default)", StringComparison.OrdinalIgnoreCase) == true ||
                                     saleType?.Trim().Equals("Goods at Standard Rate (default)", StringComparison.OrdinalIgnoreCase) == true;
                bool isExempt = saleType?.Trim().Contains("Exempt", StringComparison.OrdinalIgnoreCase) == true;
                bool isZeroRated = saleType?.Trim().Contains("zero-rate", StringComparison.OrdinalIgnoreCase) == true;

                double effectiveTaxRate = taxRate;
                if (isExempt || isZeroRated)
                {
                    effectiveTaxRate = 0;
                    Console.WriteLine($"      ⚙️ Auto-detected Exempt/Zero-Rated goods - overriding tax rate to 0%");
                }

                int quantity = Convert.ToInt32(line.Quantity?.GetValue() ?? 1);

                decimal lineAmount = Convert.ToDecimal(context.Amount);
                decimal discount = context.ApplicableDiscount;
                decimal netAmount = lineAmount - discount;

                Console.WriteLine($"      💰 Calculations:");
                Console.WriteLine($"         Line Amount: {lineAmount}");
                Console.WriteLine($"         Discount: {discount}");
                Console.WriteLine($"         Net Amount (before tax): {netAmount}");
                Console.WriteLine($"         Quantity: {quantity}");

                if (netAmount <= 0)
                {
                    Console.WriteLine($"      ⚠️ WARNING: NetAmount is {netAmount}, setting to line amount {lineAmount}");
                    netAmount = Math.Abs(lineAmount);
                }

                decimal parsedRetailPrice = 0m;
                decimal unitRetailPrice = 0m;
                if (decimal.TryParse(retailPrice, out unitRetailPrice))
                    parsedRetailPrice = unitRetailPrice * quantity;

                Console.WriteLine($"         Unit Retail Price: {unitRetailPrice}");
                Console.WriteLine($"         Total Retail Price: {parsedRetailPrice}");

                decimal displayTaxRate = 0m;
                decimal salesTaxAmount = 0m;
                decimal furtherTax = 0m;
                string rateString = "0%";
                decimal computedTotalValue = 0m; // <--- will hold final TotalValue for the invoice item

                if (isExempt)
                {
                    Console.WriteLine($"      📋 Exempt Goods detected (SN006) - no sales tax");
                    displayTaxRate = 0m;
                    salesTaxAmount = 0m;
                    furtherTax = 0m;
                    rateString = "Exempt";

                    // For normal items total = netAmount + taxes
                    computedTotalValue = netAmount + salesTaxAmount + furtherTax;
                }
                else if (isZeroRated)
                {
                    Console.WriteLine($"      📋 Zero-Rated Goods detected (SN007) - 0% tax");
                    displayTaxRate = 0m;
                    salesTaxAmount = 0m;
                    furtherTax = 0m;
                    rateString = "0%";

                    computedTotalValue = netAmount + salesTaxAmount + furtherTax;
                }
                else if (is3rdSchedule && parsedRetailPrice > 0)
                {
                    Console.WriteLine($"      📋 3rd Schedule Goods detected - FBR calculates tax on RETAIL PRICE");
                    displayTaxRate = Convert.ToDecimal(effectiveTaxRate);
                    salesTaxAmount = Convert.ToDecimal((double)parsedRetailPrice * (effectiveTaxRate / 100.0));
                    furtherTax = 0;
                    rateString = $"{effectiveTaxRate}%";

                    Console.WriteLine($"         🧮 FBR's Calculation for 3rd Schedule:");
                    Console.WriteLine($"         Retail Price (fixedNotifiedValue): {parsedRetailPrice}");
                    Console.WriteLine($"         Tax Rate: {displayTaxRate}%");
                    Console.WriteLine($"         Sales Tax = Retail Price × Rate = {salesTaxAmount}");

                    // FBR expects valueSalesExcludingST = 0.0
                    netAmount = 0m;

                    // BUT the total invoice value must be Retail Price + SalesTax (not netAmount + tax)
                    computedTotalValue = parsedRetailPrice + salesTaxAmount;

                    // Ensure extra numeric fields are numeric and label is correct
                    // (these are assigned to the InvoiceItem later)
                }
                else if (isStandardRate && effectiveTaxRate > 18)
                {
                    Console.WriteLine($"      📋 Standard Rate Goods with rate > 18% - displaying as 18% + further tax");
                    displayTaxRate = 18m;
                    salesTaxAmount = Convert.ToDecimal((double)netAmount * 0.18);
                    decimal totalTaxAmount = Convert.ToDecimal((double)netAmount * (effectiveTaxRate / 100.0));
                    furtherTax = totalTaxAmount - salesTaxAmount;
                    rateString = "18%";

                    computedTotalValue = netAmount + salesTaxAmount + furtherTax;
                }
                else
                {
                    Console.WriteLine($"      📋 Other Goods - standard tax calculation");
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

                Console.WriteLine($"         Rate String (for API): '{rateString}'");
                Console.WriteLine($"         Sales Tax: {salesTaxAmount}");
                Console.WriteLine($"         Further Tax: {furtherTax}");
                Console.WriteLine($"         Total Value (with tax): {computedTotalValue}");

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
                    NetAmount = netAmount, // this maps to valueSalesExcludingST (0 for 3rd schedule)
                    TaxRate = displayTaxRate,
                    Rate = rateString,
                    SalesTaxAmount = salesTaxAmount,
                    TotalValue = computedTotalValue, // <-- use computedTotalValue (retail + tax for 3rd schedule)
                    RetailPrice = parsedRetailPrice,
                    ExtraTax = 0m, // ensure numeric 0 not empty string
                    FurtherTax = furtherTax,
                    FedPayable = 0m,
                    SalesTaxWithheldAtSource = 0m,
                    Discount = discount,
                    SaleType = saleType ?? "Goods at standard rate (default)",
                    SroScheduleNo = is3rdSchedule ? "Third Schedule" : "",
                    SroItemSerialNo = ""
                };

                Console.WriteLine($"      ✅ Item created:");
                Console.WriteLine($"         NetAmount: {item.NetAmount}");
                Console.WriteLine($"         TaxRate: {item.TaxRate}%");
                Console.WriteLine($"         Rate (API): '{item.Rate}'");
                Console.WriteLine($"         SalesTaxAmount: {item.SalesTaxAmount}");
                Console.WriteLine($"         FurtherTax: {item.FurtherTax}");
                Console.WriteLine($"         TotalValue: {item.TotalValue}");
                Console.WriteLine($"         RetailPrice: {item.RetailPrice}");

                payload.Items.Add(item);
            }

            Console.WriteLine($"\n✅ BuildInvoiceItems Complete: Added {payload.Items.Count} items to payload\n");
        }


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
                System.Diagnostics.Debug.WriteLine($"⚠️ SRO enrichment failed: {ex.Message}");
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

                // Find most recent subtotal before this discount
                int relevantSubtotalIndex = subtotalIndices
                    .Where(idx => idx < discountIndex)
                    .OrderByDescending(idx => idx)
                    .FirstOrDefault(-1);

                if (relevantSubtotalIndex >= 0)
                {
                    // Apply to items in subtotal range
                    ApplyDiscountToSubtotalRange(lineItems, relevantSubtotalIndex, discountIndex, discountAmount);
                }
                else
                {
                    // Apply to immediate previous item
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
                // Look before subtotal
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

            // Cap discount
            totalDiscount = Math.Min(totalDiscount, totalAmount);

            // Proportional distribution
            decimal allocated = 0m;
            for (int i = 0; i < items.Count; i++)
            {
                decimal itemAmount = Math.Abs(Convert.ToDecimal(items[i].Amount));
                decimal proportion = itemAmount / totalAmount;
                decimal itemDiscount = Math.Round(totalDiscount * proportion, 2, MidpointRounding.AwayFromZero);

                items[i].ApplicableDiscount += itemDiscount;
                allocated += itemDiscount;
            }

            // Handle remainder
            decimal remainder = totalDiscount - allocated;
            if (remainder != 0m && items.Count > 0)
            {
                // Add remainder to largest item
                var largestItem = items.OrderByDescending(x => Math.Abs(Convert.ToDecimal(x.Amount))).First();
                largestItem.ApplicableDiscount += remainder;
            }
        }

        private bool IsLikelySubtotal(IInvoiceLineRet line, string itemListID, string itemType, string itemName, double lineAmount)
        {
            string type = itemType?.Trim() ?? "";
            string name = itemName?.Trim() ?? "";

            // Text markers
            if (type.Contains("subtotal", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("subtotal", StringComparison.OrdinalIgnoreCase))
                return true;

            // Empty ItemRef with no quantity
            bool itemRefEmpty = string.IsNullOrEmpty(itemListID);
            bool quantityMissing = line.Quantity == null || Math.Abs(line.Quantity.GetValue()) == 0;

            return itemRefEmpty && quantityMissing && lineAmount > 0;
        }

        #endregion

        #region Data Fetching (NO CACHING)

        private CustomerData FetchCustomerDetails(string customerListID)
        {
            if (string.IsNullOrEmpty(customerListID))
            {
                LogBoth($"⚠️ Empty customerListID, returning default CustomerData");
                return new CustomerData { CustomerType = "Unregistered" };
            }

            LogBoth($"🔍 Fetching customer details from QuickBooks (ListID: {customerListID})...");

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

                        customerData.Address = FormatAddress(customer.BillAddress);
                        customerData.Phone = customer.Phone?.GetValue() ?? "N/A";

                        // ✅ STEP 1: Get the CustomerType field from QuickBooks
                        string customerTypeFromQB = customer.CustomerTypeRef?.FullName?.GetValue() ?? "";
                        LogBoth($"   📋 QB CustomerType Field: '{customerTypeFromQB}'");

                        // ✅ STEP 2: Extract custom fields (NTN, Province)
                        if (customer.DataExtRetList != null)
                        {
                            LogBoth($"   🔍 Found {customer.DataExtRetList.Count} custom fields:");

                            for (int i = 0; i < customer.DataExtRetList.Count; i++)
                            {
                                var dataExt = customer.DataExtRetList.GetAt(i);
                                string fieldName = dataExt.DataExtName?.GetValue();
                                string fieldValue = dataExt.DataExtValue?.GetValue();

                                LogBoth($"      {i + 1}. '{fieldName}' = '{fieldValue}'");

                                if (IsNTNField(fieldName) && !string.IsNullOrEmpty(fieldValue))
                                {
                                    customerData.NTN = fieldValue.Trim();
                                    LogBoth($"      ✅ NTN extracted: {customerData.NTN}");
                                }

                                if (fieldName?.Equals("Province", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                                {
                                    customerData.State = fieldValue.Trim();
                                    LogBoth($"      ✅ Province extracted: {customerData.State}");
                                }
                            }
                        }
                        else
                        {
                            LogBoth($"   ⚠️ No custom fields found for this customer");
                        }

                        // ✅ STEP 3: Determine CustomerType using PRIORITY ORDER
                        LogBoth($"\n   🎯 Determining CustomerType...");

                        // Priority 1: Use CustomerType from QuickBooks if it's set
                        if (!string.IsNullOrWhiteSpace(customerTypeFromQB))
                        {
                            customerData.CustomerType = customerTypeFromQB.Trim();
                            LogBoth($"   ✅ Using CustomerType from QuickBooks: '{customerData.CustomerType}'");
                        }
                        // Priority 2: Fallback - determine from NTN if CustomerType not set in QB
                        else if (!string.IsNullOrWhiteSpace(customerData.NTN) && customerData.NTN.Length >= 7)
                        {
                            customerData.CustomerType = "Registered";
                            LogBoth($"   ✅ CustomerType not set in QB, determined from NTN: 'Registered'");
                        }
                        // Priority 3: Default to Unregistered
                        else
                        {
                            customerData.CustomerType = "Unregistered";
                            LogBoth($"   ⚠️ CustomerType not set in QB and no valid NTN, defaulting to: 'Unregistered'");
                        }

                        LogBoth($"\n   📦 Final CustomerData:");
                        LogBoth($"      CustomerType: '{customerData.CustomerType}'");
                        LogBoth($"      NTN: '{customerData.NTN ?? "[EMPTY]"}'");
                        LogBoth($"      Address: '{customerData.Address ?? "[EMPTY]"}'");
                        LogBoth($"      State: '{customerData.State ?? "[EMPTY]"}'");
                        LogBoth($"      Phone: '{customerData.Phone}'");
                    }
                }
                else
                {
                    LogBoth($"   ❌ QuickBooks query failed: {response.StatusMessage}");
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

            LogBoth($"🔍 Fetching item details from QuickBooks (ListID: {itemListID})...");

            var itemData = new ItemData
            {
                SaleType = "Goods at standard rate (default)"
            };

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
                        IDataExtRetList customFields = GetCustomFieldsFromItem(itemRet);

                        if (customFields != null)
                        {
                            LogBoth($"   Found {customFields.Count} custom fields");

                            for (int i = 0; i < customFields.Count; i++)
                            {
                                var dataExt = customFields.GetAt(i);
                                string fieldName = dataExt.DataExtName?.GetValue();
                                string fieldValue = dataExt.DataExtValue?.GetValue();

                                LogBoth($"   Custom Field #{i + 1}: '{fieldName}' = '{fieldValue}'");

                                if (fieldName?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    itemData.HSCode = fieldValue ?? "";
                                }
                                else if (fieldName?.Equals("Sale Type", StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    if (!string.IsNullOrWhiteSpace(fieldValue))
                                    {
                                        itemData.SaleType = fieldValue.Trim();
                                        LogBoth($"   ✅ Sale Type SET to: '{itemData.SaleType}'");
                                    }
                                    else
                                    {
                                        LogBoth($"   ⚠️ Sale Type field exists but is EMPTY");
                                    }
                                }
                                else if (fieldName?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    itemData.RetailPrice = fieldValue ?? "";
                                }
                            }
                        }
                        else
                        {
                            LogBoth($"   ⚠️ NO custom fields found for this item");
                        }

                        // Check price lists for retail price if not in custom fields
                        if (string.IsNullOrEmpty(itemData.RetailPrice))
                        {
                            itemData.RetailPrice = FetchItemRetailPriceFromPriceLists(itemListID);
                        }
                    }
                }

                LogBoth($"   📦 Final ItemData: SaleType='{itemData.SaleType}', HSCode='{itemData.HSCode}'");
            }
            catch (Exception ex)
            {
                LogBoth($"   ❌ Error: {ex.Message}");
            }

            return itemData;
        }

        private IDataExtRetList GetCustomFieldsFromItem(IORItemRet itemRet)
        {
            Console.WriteLine($"   🔍 Checking item type for custom fields...");

            if (itemRet.ItemServiceRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemServiceRet");
                return itemRet.ItemServiceRet.DataExtRetList;
            }
            if (itemRet.ItemInventoryRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemInventoryRet");
                return itemRet.ItemInventoryRet.DataExtRetList;
            }
            if (itemRet.ItemNonInventoryRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemNonInventoryRet");
                return itemRet.ItemNonInventoryRet.DataExtRetList;
            }
            if (itemRet.ItemOtherChargeRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemOtherChargeRet");
                return itemRet.ItemOtherChargeRet.DataExtRetList;
            }
            if (itemRet.ItemInventoryAssemblyRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemInventoryAssemblyRet");
                return itemRet.ItemInventoryAssemblyRet.DataExtRetList;
            }
            if (itemRet.ItemGroupRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemGroupRet");
                return itemRet.ItemGroupRet.DataExtRetList;
            }
            if (itemRet.ItemDiscountRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemDiscountRet");
                return itemRet.ItemDiscountRet.DataExtRetList;
            }
            if (itemRet.ItemPaymentRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemPaymentRet");
                return itemRet.ItemPaymentRet.DataExtRetList;
            }
            if (itemRet.ItemSalesTaxRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemSalesTaxRet");
                return itemRet.ItemSalesTaxRet.DataExtRetList;
            }
            if (itemRet.ItemSalesTaxGroupRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemSalesTaxGroupRet");
                return itemRet.ItemSalesTaxGroupRet.DataExtRetList;
            }
            if (itemRet.ItemSubtotalRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemSubtotalRet");
                return itemRet.ItemSubtotalRet.DataExtRetList;
            }
            if (itemRet.ItemFixedAssetRet != null)
            {
                Console.WriteLine($"      ✅ Found ItemFixedAssetRet");
                return itemRet.ItemFixedAssetRet.DataExtRetList;
            }

            Console.WriteLine($"      ⚠️ Unknown or unsupported item type");
            return null;
        }

        private string FetchItemRetailPriceFromPriceLists(string itemListID)
        {
            try
            {
                var priceLevels = FetchPriceLevels();
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
                System.Diagnostics.Debug.WriteLine($"Error fetching retail price: {ex.Message}");
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

            // Fallback fuzzy match
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
                    System.Diagnostics.Debug.WriteLine($"Error ending session: {ex.Message}");
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
                    System.Diagnostics.Debug.WriteLine($"Error closing connection: {ex.Message}");
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

    internal class CustomerData
    {
        public string NTN { get; set; } = "";
        public string CustomerType { get; set; } = "Unregistered";
        public string Address { get; set; } = "";
        public string State { get; set; } = "";
        public string Phone { get; set; } = "";
    }

    internal class ItemData
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