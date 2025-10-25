using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using Microsoft.VisualBasic;
using QBFC16Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Services
{
    public class QuickBooksService : IDisposable
    {
        private QBSessionManager _sessionManager;
        private bool _isConnected;
        private CompanyInfo _companyInfo;
        private Company _companySettings;
        private CompanyManager _companyManager;
        private FBRApiService _fbr;

        // Store detected QBXML version
        private short _qbXmlMajorVersion = 13;
        private short _qbXmlMinorVersion = 0;

        public string CurrentCompanyName { get; private set; }
        public string CurrentCompanyFile { get; private set; }

        public bool Connect(Company companySettings = null)
        {
            try
            {
                _companySettings = companySettings;

                _sessionManager = new QBSessionManager();
                _sessionManager.OpenConnection("", "C2B Smart App");
                _sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // Detect and store supported QBXML version
                DetectQBXMLVersion();

                // Get company info
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
                // Get supported versions from QuickBooks
                string[] supportedVersions = (string[])_sessionManager.QBXMLVersionsForSession;

                if (supportedVersions == null || supportedVersions.Length == 0)
                {
                    System.Diagnostics.Debug.WriteLine("No supported QBXML versions found, using default 13.0");
                    _qbXmlMajorVersion = 13;
                    _qbXmlMinorVersion = 0;
                    return;
                }

                // Use the highest supported version
                string highestVersion = supportedVersions[supportedVersions.Length - 1];
                string[] versionParts = highestVersion.Split('.');

                _qbXmlMajorVersion = short.Parse(versionParts[0]);
                _qbXmlMinorVersion = versionParts.Length > 1 ? short.Parse(versionParts[1]) : (short)0;

                System.Diagnostics.Debug.WriteLine($"✅ Using QBXML version: {_qbXmlMajorVersion}.{_qbXmlMinorVersion}");
                System.Diagnostics.Debug.WriteLine($"Available versions: {string.Join(", ", supportedVersions)}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error detecting QBXML version: {ex.Message}. Using default 13.0");
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

                    System.Diagnostics.Debug.WriteLine($"✅ Successfully fetched company info: {CurrentCompanyName}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Warning: Could not fetch company information: {ex.Message}");
            }
        }

        /// Gets the current QuickBooks company information
        /// <returns>Company information including address and city</returns>
        public CompanyInfo GetCompanyInfo()
        {
            return _companyInfo;
        }

        public List<Invoice> FetchInvoices()
        {
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

                var invoices = new List<Invoice>();

                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                // Create the invoice query
                var invoiceQuery = msgSetRq.AppendInvoiceQueryRq();
                invoiceQuery.IncludeLineItems.SetValue(true);

                // Send request to QuickBooks
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

                            string customerName = inv.CustomerRef?.FullName?.GetValue() ?? "";
                            string customerListID = inv.CustomerRef?.ListID?.GetValue() ?? "";

                            // Fetch customer details for NTN
                            var customerData = FetchCustomerDetails(customerListID);

                            var invoice = new Invoice
                            {
                                CompanyName = CurrentCompanyName,
                                QuickBooksInvoiceId = inv.TxnID.GetValue(),
                                InvoiceNumber = inv.RefNumber?.GetValue() ?? "",
                                CustomerName = customerName,
                                CustomerNTN = customerData.NTN,
                                Amount = Convert.ToDecimal(inv.Subtotal?.GetValue() ?? 0),
                                Status = "Pending",
                                CreatedDate = inv.TxnDate?.GetValue() ?? DateTime.Now
                            };

                            invoices.Add(invoice);
                        }
                    }
                }

                return invoices;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to fetch all invoices: {ex.Message}", ex);
            }
            finally
            {
                // Close session if we opened it in this method
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

                if (response.StatusCode == 0)
                {
                    var invoiceRetList = response.Detail as IInvoiceRetList;
                    if (invoiceRetList != null && invoiceRetList.Count > 0)
                    {
                        var inv = invoiceRetList.GetAt(0);
                        string customerListID = inv.CustomerRef?.ListID?.GetValue() ?? "";

                        // Fetch comprehensive customer details
                        var customerData = FetchCustomerDetails(customerListID);

                        var payload = new FBRInvoicePayload
                        {
                            // Invoice Type
                            InvoiceType = "Sale Invoice",

                            // Invoice Information
                            InvoiceNumber = inv.RefNumber?.GetValue() ?? "",
                            InvoiceDate = inv.TxnDate?.GetValue() ?? DateTime.Now,

                            // Seller Information
                            SellerNTN = _companySettings?.SellerNTN ?? _companyInfo?.NTN ?? "",
                            SellerBusinessName = _companyInfo?.Name ?? CurrentCompanyName,
                            SellerProvince = _companySettings?.SellerProvince ?? "",
                            SellerAddress = _companySettings?.SellerAddress ?? _companyInfo?.Address ?? "",

                            // Buyer Information
                            CustomerName = inv.CustomerRef?.FullName?.GetValue() ?? "",
                            CustomerNTN = customerData.NTN,
                            BuyerProvince = customerData.State,
                            BuyerAddress = customerData.Address,
                            BuyerRegistrationType = customerData.CustomerType,
                            BuyerPhone = customerData.Phone,

                            // Financial Summary
                            TotalAmount = Convert.ToDecimal((inv.Subtotal?.GetValue() ?? 0) + (inv.SalesTaxTotal?.GetValue() ?? 0)),
                            Subtotal = Convert.ToDecimal(inv.Subtotal?.GetValue() ?? 0),
                            TaxAmount = Convert.ToDecimal(inv.SalesTaxTotal?.GetValue() ?? 0),

                            Items = new List<InvoiceItem>()
                        };

                        // Calculate tax rate
                        double taxRate = 0;
                        if (inv.SalesTaxPercentage != null)
                        {
                            taxRate = inv.SalesTaxPercentage.GetValue();
                        }

                        // ==========================================
                        // NEW LOGIC: Process lines with contextual discount tracking
                        // ==========================================

                        var lineItemsWithContext = new List<LineItemContext>();
                        decimal lastSubtotalValue = 0;
                        int lastSubtotalIndex = -1;

                        if (inv.ORInvoiceLineRetList != null)
                        {
                            for (int i = 0; i < inv.ORInvoiceLineRetList.Count; i++)
                            {
                                var lineRet = inv.ORInvoiceLineRetList.GetAt(i);

                                if (lineRet.InvoiceLineRet != null)
                                {
                                    var line = lineRet.InvoiceLineRet;
                                    string itemListID = line.ItemRef?.ListID?.GetValue() ?? "";
                                    double lineAmount = line.Amount?.GetValue() ?? 0;
                                    string itemType = line.ItemRef?.FullName?.GetValue() ?? "";
                                    string itemName = line.Desc?.GetValue() ?? "";

                                    // IMPROVED SUBTOTAL DETECTION - Robust helper
                                    bool isSubtotal = IsLikelySubtotal(line, itemListID, itemType, itemName, lineAmount);
                                    if (isSubtotal)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"  → Subtotal detected (Robust): ItemRef='{itemListID}', Name='{itemType}', Desc='{itemName}', Amount={lineAmount}");
                                    }

                                    // DISCOUNT DETECTION
                                    bool isDiscount = lineAmount < 0 ||
                                                     (!string.IsNullOrEmpty(itemType) &&
                                                      itemType.Contains("Discount", StringComparison.OrdinalIgnoreCase));

                                    var context = new LineItemContext
                                    {
                                        Index = i,
                                        Line = line,
                                        ItemListID = itemListID,
                                        Amount = lineAmount,
                                        ItemType = itemType,
                                        ItemName = itemName,
                                        IsDiscount = isDiscount,
                                        IsSubtotal = isSubtotal
                                    };

                                    lineItemsWithContext.Add(context);

                                    // Track subtotal positions
                                    if (isSubtotal)
                                    {
                                        lastSubtotalValue = Math.Abs(Convert.ToDecimal(lineAmount));
                                        lastSubtotalIndex = i;
                                        System.Diagnostics.Debug.WriteLine(
                                            $"✓ SUBTOTAL at index {i}: " +
                                            $"Value={lastSubtotalValue}, " +
                                            $"ItemType='{itemType}', " +
                                            $"Desc='{itemName}'");
                                    }

                                    System.Diagnostics.Debug.WriteLine(
                                        $"Line {i}: " +
                                        $"Name='{itemName}', " +
                                        $"Type='{itemType}', " +
                                        $"ListID='{(string.IsNullOrEmpty(itemListID) ? "EMPTY" : "HAS_ID")}', " +
                                        $"Amount={lineAmount:F2}, " +
                                        $"IsDiscount={isDiscount}, " +
                                        $"IsSubtotal={isSubtotal}");
                                }
                            }
                        }

                        // Apply discounts based on QuickBooks contextual rules
                        ApplyDiscountsContextually(lineItemsWithContext, lastSubtotalIndex);

                        // Now process regular items and build the payload
                        for (int i = 0; i < lineItemsWithContext.Count; i++)
                        {
                            var context = lineItemsWithContext[i];

                            // Skip discount items
                            if (context.IsDiscount)
                            {
                                System.Diagnostics.Debug.WriteLine($"⊗ SKIPPING discount at index {i}: {context.ItemName}");
                                continue;
                            }

                            // Skip subtotal items
                            if (context.IsSubtotal)
                            {
                                System.Diagnostics.Debug.WriteLine($"⊗ SKIPPING subtotal at index {i}: {context.ItemName}");
                                continue;
                            }

                            // Additional safety check: Skip lines with no valid ItemRef
                            if (string.IsNullOrEmpty(context.ItemListID) &&
                                string.IsNullOrEmpty(context.ItemType))
                            {
                                System.Diagnostics.Debug.WriteLine($"⊗ SKIPPING line with no ItemRef at index {i}: {context.ItemName}");
                                continue;
                            }

                            var line = context.Line;

                            // Fetch item details for HS Code and Retail Price
                            var itemData = FetchItemDetails(context.ItemListID);
                            string hsCode = itemData.HSCode;
                            string retailPrice = itemData.RetailPrice;

                            System.Diagnostics.Debug.WriteLine($"Item: {line.Desc?.GetValue()}, Retail Price: '{retailPrice}'");

                            // Check line-level custom fields
                            if (line.DataExtRetList != null && line.DataExtRetList.Count > 0)
                            {
                                for (int k = 0; k < line.DataExtRetList.Count; k++)
                                {
                                    var dataExt = line.DataExtRetList.GetAt(k);
                                    string fieldName = dataExt.DataExtName?.GetValue();
                                    string fieldValue = dataExt.DataExtValue?.GetValue();

                                    if (fieldName?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                                    {
                                        hsCode = fieldValue;
                                    }
                                    else if (fieldName?.Equals("Residential", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                                    {
                                        retailPrice = fieldValue;
                                    }
                                }
                            }

                            // Calculate line-level tax with split logic
                            double actualTaxRate = taxRate;
                            double salesTaxAmount = context.Amount * (taxRate / 100);

                            decimal displayTaxRate;
                            decimal standardTaxAmount;
                            decimal furtherTax;

                            if (actualTaxRate > 18)
                            {
                                displayTaxRate = 18;
                                standardTaxAmount = Convert.ToDecimal(context.Amount * 0.18);
                                furtherTax = Convert.ToDecimal(salesTaxAmount) - standardTaxAmount;
                            }
                            else
                            {
                                displayTaxRate = Convert.ToDecimal(actualTaxRate);
                                standardTaxAmount = Convert.ToDecimal(salesTaxAmount);
                                furtherTax = 0;
                            }

                            var item = new InvoiceItem
                            {
                                ItemName = line.Desc?.GetValue() ?? "",
                                HSCode = hsCode,
                                Quantity = Convert.ToInt32(line.Quantity?.GetValue() ?? 1),
                                UnitOfMeasure = await _fbr.GetUOMDescriptionAsync(hsCode, _companySettings.FBRToken) ?? line.UnitOfMeasure?.GetValue(),
                                UnitPrice = Convert.ToDecimal(line.Amount?.GetValue() ?? 0),
                                TotalPrice = Convert.ToDecimal(context.Amount),
                                TaxRate = displayTaxRate,
                                SalesTaxAmount = standardTaxAmount,
                                TotalValue = Convert.ToDecimal(context.Amount) + standardTaxAmount + furtherTax,
                                RetailPrice = decimal.TryParse(retailPrice, out var parsedRetailPrice) ? parsedRetailPrice : 0,
                                ExtraTax = furtherTax,
                                FurtherTax = 0,
                                Discount = context.ApplicableDiscount // Apply the calculated discount from context
                            };

                            payload.Items.Add(item);

                            System.Diagnostics.Debug.WriteLine(
                                $"Added item: {item.ItemName}, " +
                                $"TotalPrice: {item.TotalPrice}, " +
                                $"Discount: {item.Discount}");
                        }

                        System.Diagnostics.Debug.WriteLine($"Total items in payload: {payload.Items.Count}");

                        return payload;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to get invoice details: {ex.Message}", ex);
            }
            finally
            {
                // Close session if we opened it in this method
                if (!wasConnected && _isConnected)
                {
                    CloseSession();
                    CloseConnection();
                }
            }
        }

        private CustomerData FetchCustomerDetails(string customerListID)
        {
            var customerData = new CustomerData();

            if (string.IsNullOrEmpty(customerListID)) return customerData;

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var customerQuery = msgSetRq.AppendCustomerQueryRq();
                customerQuery.ORCustomerListQuery.ListIDList.Add(customerListID);
                customerQuery.OwnerIDList.Add("0"); // Include custom fields

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode == 0)
                {
                    var customerList = response.Detail as ICustomerRetList;

                    if (customerList != null && customerList.Count > 0)
                    {
                        var customer = customerList.GetAt(0);

                        // Get address information
                        customerData.Address = FormatAddress(customer.BillAddress);

                        // Get customer type from QuickBooks
                        customerData.CustomerType = customer.CustomerTypeRef?.FullName?.GetValue() ?? "Unregistered";
                        customerData.Phone = customer.Phone?.GetValue() ?? "N/A";

                        // Read custom fields (NTN and Province)
                        if (customer.DataExtRetList != null && customer.DataExtRetList.Count > 0)
                        {
                            for (int i = 0; i < customer.DataExtRetList.Count; i++)
                            {
                                var dataExt = customer.DataExtRetList.GetAt(i);
                                string fieldName = dataExt.DataExtName?.GetValue();
                                string fieldValue = dataExt.DataExtValue?.GetValue();

                                // Check for NTN/CNIC
                                if ((fieldName?.Equals("NTN", StringComparison.OrdinalIgnoreCase) == true ||
                                     fieldName?.Equals("CNIC", StringComparison.OrdinalIgnoreCase) == true ||
                                     fieldName?.Equals("NTN/CNIC", StringComparison.OrdinalIgnoreCase) == true) &&
                                    !string.IsNullOrEmpty(fieldValue))
                                {
                                    customerData.NTN = fieldValue;
                                }

                                // Check for Province custom field
                                if (fieldName?.Equals("Province", StringComparison.OrdinalIgnoreCase) == true &&
                                    !string.IsNullOrEmpty(fieldValue))
                                {
                                    customerData.State = fieldValue;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error fetching customer details: {ex.Message}");
            }

            return customerData;
        }

        private ItemData FetchItemDetails(string itemListID)
        {
            var itemData = new ItemData();

            if (string.IsNullOrEmpty(itemListID)) return itemData;

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var itemQuery = msgSetRq.AppendItemQueryRq();
                itemQuery.ORListQuery.ListIDList.Add(itemListID);
                itemQuery.OwnerIDList.Add("0"); // Include custom fields

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode == 0)
                {
                    var itemList = response.Detail as IORItemRetList;

                    if (itemList != null && itemList.Count > 0)
                    {
                        var itemRet = itemList.GetAt(0);
                        IDataExtRetList customFields = null;
                        string itemName = "";

                        // Check which type of item it is and get custom fields
                        if (itemRet.ItemServiceRet != null)
                        {
                            customFields = itemRet.ItemServiceRet.DataExtRetList;
                            itemName = itemRet.ItemServiceRet.Name?.GetValue();
                        }
                        else if (itemRet.ItemInventoryRet != null)
                        {
                            customFields = itemRet.ItemInventoryRet.DataExtRetList;
                            itemName = itemRet.ItemInventoryRet.Name?.GetValue();
                        }
                        else if (itemRet.ItemNonInventoryRet != null)
                        {
                            customFields = itemRet.ItemNonInventoryRet.DataExtRetList;
                            itemName = itemRet.ItemNonInventoryRet.Name?.GetValue();
                        }
                        else if (itemRet.ItemOtherChargeRet != null)
                        {
                            customFields = itemRet.ItemOtherChargeRet.DataExtRetList;
                            itemName = itemRet.ItemOtherChargeRet.Name?.GetValue();
                        }
                        else if (itemRet.ItemInventoryAssemblyRet != null)
                        {
                            customFields = itemRet.ItemInventoryAssemblyRet.DataExtRetList;
                            itemName = itemRet.ItemInventoryAssemblyRet.Name?.GetValue();
                        }
                        else if (itemRet.ItemGroupRet != null)
                        {
                            customFields = itemRet.ItemGroupRet.DataExtRetList;
                            itemName = itemRet.ItemGroupRet.Name?.GetValue();
                        }

                        System.Diagnostics.Debug.WriteLine($"Fetching details for item: {itemName} (ID: {itemListID})");

                        // First, check custom fields for HS Code and Retail Price
                        bool retailPriceFound = false;

                        if (customFields != null && customFields.Count > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"  Checking {customFields.Count} custom fields...");

                            for (int i = 0; i < customFields.Count; i++)
                            {
                                var dataExt = customFields.GetAt(i);
                                string fieldName = dataExt.DataExtName?.GetValue();
                                string fieldValue = dataExt.DataExtValue?.GetValue();

                                System.Diagnostics.Debug.WriteLine($"    Custom field: {fieldName} = '{fieldValue}'");

                                if (fieldName?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true &&
                                    !string.IsNullOrEmpty(fieldValue))
                                {
                                    itemData.HSCode = fieldValue;
                                    System.Diagnostics.Debug.WriteLine($"    HS Code found: {fieldValue}");
                                }

                                if (fieldName?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true &&
                                    !string.IsNullOrEmpty(fieldValue))
                                {
                                    itemData.RetailPrice = fieldValue;
                                    retailPriceFound = true;
                                    System.Diagnostics.Debug.WriteLine($"    Retail Price found in custom field: {fieldValue}");
                                }
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"  No custom fields found");
                        }

                        // If retail price not found in custom fields, check price lists
                        if (!retailPriceFound)
                        {
                            System.Diagnostics.Debug.WriteLine($"  Retail price not in custom fields, checking price lists...");
                            var priceInfo = FetchItemRetailPriceFromPriceLists(itemListID);
                            if (!string.IsNullOrEmpty(priceInfo))
                            {
                                itemData.RetailPrice = priceInfo;
                                System.Diagnostics.Debug.WriteLine($"    Retail Price found in price list: {priceInfo}");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"    No retail price found in price lists");
                            }
                        }

                        System.Diagnostics.Debug.WriteLine($"  Final HS Code: '{itemData.HSCode}'");
                        System.Diagnostics.Debug.WriteLine($"  Final Retail Price: '{itemData.RetailPrice}'");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error fetching item details: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }

            return itemData;
        }

        private string FetchItemRetailPriceFromPriceLists(string itemListID)
        {
            try
            {
                // Get all price levels
                var allPriceLevels = FetchPriceLevels();

                System.Diagnostics.Debug.WriteLine($"    Found {allPriceLevels.Count} price levels total");

                // Look specifically for "Retail Price" price level
                foreach (var priceLevel in allPriceLevels)
                {
                    System.Diagnostics.Debug.WriteLine($"    Checking price level: '{priceLevel.Name}'");

                    if (priceLevel.Name?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true &&
                        priceLevel.Items != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"      Found 'Retail Price' level with {priceLevel.Items.Count} items");

                        var itemPrice = priceLevel.Items.FirstOrDefault(i => i.ItemListID == itemListID);
                        if (itemPrice != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"      Item found in Retail Price list!");
                            System.Diagnostics.Debug.WriteLine($"      CustomPrice: {itemPrice.CustomPrice}, CustomPercent: {itemPrice.CustomPricePercent}");

                            if (itemPrice.CustomPrice > 0)
                            {
                                return itemPrice.CustomPrice.ToString("0.##");
                            }
                            else if (itemPrice.CustomPricePercent != 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"      Using percent-based price calculation");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"      Item {itemListID} not found in Retail Price list");
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"    No 'Retail Price' price level found or item not in it");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error fetching retail price from price lists: {ex.Message}");
            }

            return "";
        }

        private string FormatAddress(IAddress address)
        {
            if (address == null) return "";
            try
            {
                var lines = new List<string>();

                // Add all address lines that have data
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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FormatAddress error: {ex.Message}");
                return "";
            }
        }

        private void AddAddressLine(List<string> lines, IQBStringType addressField)
        {
            if (addressField != null)
            {
                var value = addressField.GetValue();
                if (!string.IsNullOrEmpty(value))
                {
                    lines.Add(value.TrimEnd(',', ' '));
                }
            }
        }

        /// Fetches all price levels from QuickBooks
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
                if (priceLevelList == null)
                    return priceLevels;

                for (int i = 0; i < priceLevelList.Count; i++)
                {
                    var priceLevel = priceLevelList.GetAt(i);
                    var level = new PriceLevel
                    {
                        ListID = priceLevel.ListID?.GetValue(),
                        Name = priceLevel.Name?.GetValue(),
                        IsActive = priceLevel.IsActive?.GetValue() ?? true,
                        PriceLevelType = priceLevel.PriceLevelType?.GetValue().ToString() ?? ""
                    };

                    level.Items = new List<PriceLevelItem>();
                    level.FixedPercentage = 0;

                    // Access ORPriceLevelRet
                    IORPriceLevelRet orPriceLevel = priceLevel.ORPriceLevelRet;
                    if (orPriceLevel != null)
                    {
                        // Check the type using ortype enum
                        ENORPriceLevelRet priceType = orPriceLevel.ortype;

                        if (priceType == ENORPriceLevelRet.orplrPriceLevelFixedPercentage)
                        {
                            // Handle Fixed Percentage
                            IQBPercentType fixedPct = orPriceLevel.PriceLevelFixedPercentage;
                            if (fixedPct != null)
                            {
                                level.FixedPercentage = fixedPct.GetValue();
                                System.Diagnostics.Debug.WriteLine($"✅ Fixed percentage level '{level.Name}': {level.FixedPercentage}%");
                            }
                        }
                        else if (priceType == ENORPriceLevelRet.orplrPriceLevelPerItemRetCurrency)
                        {
                            // Handle Per-Item pricing
                            IPriceLevelPerItemRetCurrency perItemCurrency = orPriceLevel.PriceLevelPerItemRetCurrency;

                            if (perItemCurrency != null)
                            {
                                IPriceLevelPerItemRetList itemList = perItemCurrency.PriceLevelPerItemRetList;

                                if (itemList != null)
                                {
                                    for (int j = 0; j < itemList.Count; j++)
                                    {
                                        IPriceLevelPerItemRet itemPrice = itemList.GetAt(j);

                                        string itemListID = itemPrice.ItemRef?.ListID?.GetValue();
                                        string itemFullName = itemPrice.ItemRef?.FullName?.GetValue();

                                        decimal customPrice = 0;
                                        double customPricePercent = 0;

                                        // Access ORORCustomPrice
                                        IORORCustomPrice customPriceObj = itemPrice.ORORCustomPrice;
                                        if (customPriceObj != null)
                                        {
                                            if (customPriceObj.CustomPrice != null)
                                            {
                                                customPrice = Convert.ToDecimal(customPriceObj.CustomPrice.GetValue());
                                            }
                                            else if (customPriceObj.CustomPricePercent != null)
                                            {
                                                customPricePercent = customPriceObj.CustomPricePercent.GetValue();
                                            }
                                        }

                                        level.Items.Add(new PriceLevelItem
                                        {
                                            ItemListID = itemListID,
                                            ItemFullName = itemFullName,
                                            CustomPrice = customPrice,
                                            CustomPricePercent = customPricePercent
                                        });
                                    }
                                }
                            }
                        }
                    }

                    priceLevels.Add(level);
                }

                System.Diagnostics.Debug.WriteLine($"✅ Successfully fetched {priceLevels.Count} price levels from QuickBooks.");
                foreach (var lvl in priceLevels)
                {
                    if (lvl.PriceLevelType == "pltPerItem")
                    {
                        System.Diagnostics.Debug.WriteLine($"  Per-Item Level: {lvl.Name} - {lvl.Items?.Count ?? 0} items");
                    }
                    else if (lvl.PriceLevelType == "pltFixedPercentage")
                    {
                        System.Diagnostics.Debug.WriteLine($"  Fixed Percentage Level: {lvl.Name} - {lvl.FixedPercentage}%");
                    }
                }

                return priceLevels;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to fetch price levels: {ex.Message}", ex);
            }
            finally
            {
                // Close session if we opened it in this method
                if (!wasConnected && _isConnected)
                {
                    CloseSession();
                    CloseConnection();
                }
            }
        }

        private void ApplyDiscountsContextually(List<LineItemContext> lineItems, int lastSubtotalIndex)
        {
            System.Diagnostics.Debug.WriteLine("\n╔══════════════════════════════════════════════════════════╗");
            System.Diagnostics.Debug.WriteLine("║    STARTING CONTEXTUAL DISCOUNT APPLICATION             ║");
            System.Diagnostics.Debug.WriteLine("╚══════════════════════════════════════════════════════════╝\n");

            // Track all subtotal positions for proper range handling
            var subtotalIndices = new List<int>();
            for (int i = 0; i < lineItems.Count; i++)
            {
                if (lineItems[i].IsSubtotal)
                {
                    subtotalIndices.Add(i);
                }
            }

            if (subtotalIndices.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"📊 Found {subtotalIndices.Count} subtotal(s) at indices: {string.Join(", ", subtotalIndices)}");
            }

            // Track discount statistics
            int discountsProcessed = 0;
            int discountsApplied = 0;
            int discountsSkipped = 0;

            for (int i = 0; i < lineItems.Count; i++)
            {
                var currentItem = lineItems[i];

                // Skip if not a discount
                if (!currentItem.IsDiscount)
                    continue;

                discountsProcessed++;
                decimal discountAmount = Math.Abs(Convert.ToDecimal(currentItem.Amount));

                System.Diagnostics.Debug.WriteLine($"\n──────────────────────────────────────────────────────────");
                System.Diagnostics.Debug.WriteLine($"💰 Processing Discount #{discountsProcessed} at index {i}");
                System.Diagnostics.Debug.WriteLine($"   Name: {currentItem.ItemName}");
                System.Diagnostics.Debug.WriteLine($"   Amount: {discountAmount:C}");

                // Check for consecutive discounts (unusual but possible)
                if (i + 1 < lineItems.Count && lineItems[i + 1].IsDiscount)
                {
                    System.Diagnostics.Debug.WriteLine($"   ⚠ Note: Multiple consecutive discounts detected");
                }

                // Find the most recent subtotal BEFORE this discount
                int relevantSubtotalIndex = -1;
                for (int j = i - 1; j >= 0; j--)
                {
                    if (lineItems[j].IsSubtotal)
                    {
                        relevantSubtotalIndex = j;
                        break;
                    }
                }

                // RULE 1: If there's a subtotal before this discount
                if (relevantSubtotalIndex >= 0)
                {
                    System.Diagnostics.Debug.WriteLine($"   📍 Found subtotal at index {relevantSubtotalIndex}");

                    // Check if discount is immediately after subtotal (no items between them)
                    bool discountImmediatelyAfterSubtotal = true;
                    for (int j = relevantSubtotalIndex + 1; j < i; j++)
                    {
                        if (!lineItems[j].IsDiscount && !lineItems[j].IsSubtotal)
                        {
                            discountImmediatelyAfterSubtotal = false;
                            break;
                        }
                    }

                    if (discountImmediatelyAfterSubtotal)
                    {
                        // Find the start of the subtotal range
                        int rangeStart = 0;
                        for (int j = relevantSubtotalIndex - 1; j >= 0; j--)
                        {
                            if (lineItems[j].IsSubtotal)
                            {
                                rangeStart = j + 1;
                                break;
                            }
                        }

                        // Find all items between rangeStart and the subtotal
                        var itemsInSubtotalRange = new List<LineItemContext>();
                        for (int j = rangeStart; j < relevantSubtotalIndex; j++)
                        {
                            if (!lineItems[j].IsDiscount && !lineItems[j].IsSubtotal)
                            {
                                itemsInSubtotalRange.Add(lineItems[j]);
                            }
                        }

                        if (itemsInSubtotalRange.Count > 0)
                        {
                            string discountType = GetDiscountType(currentItem, itemsInSubtotalRange);
                            System.Diagnostics.Debug.WriteLine($"   ✓ RULE 1A: {discountType} immediately after subtotal");
                            System.Diagnostics.Debug.WriteLine($"   ✓ Applying to {itemsInSubtotalRange.Count} items BEFORE subtotal");
                            System.Diagnostics.Debug.WriteLine($"   ✓ Item range: [{rangeStart}..{relevantSubtotalIndex - 1}]");

                            DistributeDiscountToItems(itemsInSubtotalRange, discountAmount);
                            discountsApplied++;
                            continue;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"   ⚠ No items found in subtotal range");
                            discountsSkipped++;
                        }
                    }
                    else
                    {
                        // There are items between subtotal and discount
                        var itemsBetweenSubtotalAndDiscount = new List<LineItemContext>();
                        for (int j = relevantSubtotalIndex + 1; j < i; j++)
                        {
                            if (!lineItems[j].IsDiscount && !lineItems[j].IsSubtotal)
                            {
                                itemsBetweenSubtotalAndDiscount.Add(lineItems[j]);
                            }
                        }

                        if (itemsBetweenSubtotalAndDiscount.Count > 0)
                        {
                            string discountType = GetDiscountType(currentItem, itemsBetweenSubtotalAndDiscount);
                            System.Diagnostics.Debug.WriteLine($"   ✓ RULE 1B: {discountType} with items between subtotal and discount");
                            System.Diagnostics.Debug.WriteLine($"   ✓ Applying to {itemsBetweenSubtotalAndDiscount.Count} items AFTER subtotal");
                            System.Diagnostics.Debug.WriteLine($"   ✓ Item range: [{relevantSubtotalIndex + 1}..{i - 1}]");

                            DistributeDiscountToItems(itemsBetweenSubtotalAndDiscount, discountAmount);
                            discountsApplied++;
                            continue;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"   ⚠ No items found between subtotal and discount");
                            discountsSkipped++;
                        }
                    }
                }

                // RULE 2: No subtotal before discount - apply to immediate previous item
                LineItemContext previousItem = null;
                int previousItemIndex = -1;
                for (int j = i - 1; j >= 0; j--)
                {
                    if (!lineItems[j].IsDiscount && !lineItems[j].IsSubtotal)
                    {
                        previousItem = lineItems[j];
                        previousItemIndex = j;
                        break;
                    }
                }

                if (previousItem != null)
                {
                    System.Diagnostics.Debug.WriteLine($"   ✓ RULE 2: No subtotal found, applying to previous item");
                    System.Diagnostics.Debug.WriteLine($"   ✓ Target: {previousItem.ItemName} (index {previousItemIndex})");

                    // Validate discount doesn't exceed item amount
                    decimal itemAmount = Math.Abs(Convert.ToDecimal(previousItem.Amount));
                    decimal existingDiscount = previousItem.ApplicableDiscount;
                    decimal totalDiscountForItem = existingDiscount + discountAmount;

                    if (totalDiscountForItem > itemAmount)
                    {
                        System.Diagnostics.Debug.WriteLine($"   ⚠ WARNING: Total discount ({totalDiscountForItem:C}) exceeds item amount ({itemAmount:C})");
                        System.Diagnostics.Debug.WriteLine($"   ⚠ Capping discount at {itemAmount:C}");
                        previousItem.ApplicableDiscount = itemAmount;
                    }
                    else
                    {
                        previousItem.ApplicableDiscount += discountAmount;
                        System.Diagnostics.Debug.WriteLine($"   ✓ Applied {discountAmount:C} (Total: {previousItem.ApplicableDiscount:C})");
                    }

                    discountsApplied++;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"   ❌ ERROR: No applicable item found for discount!");
                    System.Diagnostics.Debug.WriteLine($"   ❌ This discount will not be applied");
                    discountsSkipped++;
                }
            }

            // Generate comprehensive summary
            System.Diagnostics.Debug.WriteLine($"\n──────────────────────────────────────────────────────────");

            decimal totalDiscountsApplied = 0m;
            int itemsWithDiscounts = 0;
            var itemDiscountDetails = new List<string>();

            foreach (var item in lineItems)
            {
                if (!item.IsDiscount && !item.IsSubtotal && item.ApplicableDiscount > 0)
                {
                    totalDiscountsApplied += item.ApplicableDiscount;
                    itemsWithDiscounts++;
                    itemDiscountDetails.Add($"   • {item.ItemName}: {item.ApplicableDiscount:C}");
                }
            }

            System.Diagnostics.Debug.WriteLine("\n╔══════════════════════════════════════════════════════════╗");
            System.Diagnostics.Debug.WriteLine("║        DISCOUNT APPLICATION SUMMARY                      ║");
            System.Diagnostics.Debug.WriteLine("╚══════════════════════════════════════════════════════════╝");
            System.Diagnostics.Debug.WriteLine($"📊 Discounts processed: {discountsProcessed}");
            System.Diagnostics.Debug.WriteLine($"✅ Discounts applied: {discountsApplied}");
            System.Diagnostics.Debug.WriteLine($"⚠  Discounts skipped: {discountsSkipped}");
            System.Diagnostics.Debug.WriteLine($"💵 Total discount amount: {totalDiscountsApplied:C}");
            System.Diagnostics.Debug.WriteLine($"📦 Items receiving discounts: {itemsWithDiscounts}");
            System.Diagnostics.Debug.WriteLine($"📋 Total items processed: {lineItems.Count(x => !x.IsDiscount && !x.IsSubtotal)}");

            if (itemDiscountDetails.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"\n💰 Item-by-item breakdown:");
                foreach (var detail in itemDiscountDetails)
                {
                    System.Diagnostics.Debug.WriteLine(detail);
                }
            }

            System.Diagnostics.Debug.WriteLine("\n╚══════════════════════════════════════════════════════════╝\n");
        }

        private void DistributeDiscountToItems(List<LineItemContext> items, decimal totalDiscount)
        {
            System.Diagnostics.Debug.WriteLine($"\n   ┌─ Discount Distribution ─────────────────────────────");

            // Enhanced validation
            if (items == null || items.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine($"   │ ⚠ ERROR: No items to distribute discount to");
                System.Diagnostics.Debug.WriteLine($"   └─────────────────────────────────────────────────────\n");
                return;
            }

            if (totalDiscount <= 0)
            {
                System.Diagnostics.Debug.WriteLine($"   │ ⚠ ERROR: Invalid discount amount: {totalDiscount}");
                System.Diagnostics.Debug.WriteLine($"   └─────────────────────────────────────────────────────\n");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"   │ Items to process: {items.Count}");
            System.Diagnostics.Debug.WriteLine($"   │ Total discount: {totalDiscount:C}");

            // Calculate total amount with validation
            decimal totalAmount = 0m;
            bool hasNegativeAmounts = false;

            foreach (var item in items)
            {
                decimal amt = Math.Abs(Convert.ToDecimal(item.Amount));
                if (item.Amount < 0)
                {
                    System.Diagnostics.Debug.WriteLine($"   │ ⚠ WARNING: Negative amount for {item.ItemName}: {item.Amount}");
                    hasNegativeAmounts = true;
                }
                totalAmount += amt;
            }

            if (totalAmount <= 0)
            {
                System.Diagnostics.Debug.WriteLine($"   │ ⚠ ERROR: Total amount is zero or negative: {totalAmount}");
                System.Diagnostics.Debug.WriteLine($"   └─────────────────────────────────────────────────────\n");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"   │ Total items amount: {totalAmount:C}");

            // Cap discount at total amount
            if (totalDiscount > totalAmount)
            {
                System.Diagnostics.Debug.WriteLine($"   │ ⚠ WARNING: Discount ({totalDiscount:C}) > Total ({totalAmount:C})");
                System.Diagnostics.Debug.WriteLine($"   │ ⚠ Capping discount at total amount");
                totalDiscount = totalAmount;
            }

            // Calculate discount percentage
            decimal discountPercentage = (totalDiscount / totalAmount) * 100;
            System.Diagnostics.Debug.WriteLine($"   │ Discount percentage: {discountPercentage:F2}%");
            System.Diagnostics.Debug.WriteLine($"   │");

            // Proportional distribution
            var discounts = new decimal[items.Count];
            decimal allocated = 0m;

            for (int i = 0; i < items.Count; i++)
            {
                decimal itemAmount = Math.Abs(Convert.ToDecimal(items[i].Amount));
                decimal proportion = itemAmount / totalAmount;
                discounts[i] = Math.Round(totalDiscount * proportion, 2, MidpointRounding.AwayFromZero);
                allocated += discounts[i];
            }

            // Improved remainder handling
            decimal remainder = totalDiscount - allocated;

            if (remainder != 0m)
            {
                System.Diagnostics.Debug.WriteLine($"   │ 🔄 Rounding remainder: {remainder:C}");

                if (Math.Abs(remainder) <= 0.02m * items.Count)
                {
                    // Small remainder - distribute one cent at a time
                    System.Diagnostics.Debug.WriteLine($"   │ 📊 Using penny distribution strategy");
                    int startIdx = 0;
                    decimal sign = Math.Sign(remainder);
                    int iterations = 0;
                    int maxIterations = items.Count * 10; // Safety limit

                    while (remainder != 0m && iterations < maxIterations)
                    {
                        discounts[startIdx] += 0.01m * sign;
                        allocated += 0.01m * sign;
                        remainder = totalDiscount - allocated;
                        startIdx = (startIdx + 1) % items.Count;
                        iterations++;
                    }

                    if (iterations >= maxIterations)
                    {
                        System.Diagnostics.Debug.WriteLine($"   │ ⚠ Hit iteration limit, remaining: {remainder:C}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"   │ ✓ Distributed in {iterations} iterations");
                    }
                }
                else
                {
                    // Large remainder - add to largest item (more natural)
                    System.Diagnostics.Debug.WriteLine($"   │ 📊 Using largest-item strategy");
                    int maxIdx = 0;
                    decimal maxAmt = Math.Abs(Convert.ToDecimal(items[0].Amount));

                    for (int i = 1; i < items.Count; i++)
                    {
                        decimal amt = Math.Abs(Convert.ToDecimal(items[i].Amount));
                        if (amt > maxAmt)
                        {
                            maxAmt = amt;
                            maxIdx = i;
                        }
                    }

                    discounts[maxIdx] += remainder;
                    System.Diagnostics.Debug.WriteLine($"   │ ✓ Added {remainder:C} to {items[maxIdx].ItemName}");
                }
            }

            System.Diagnostics.Debug.WriteLine($"   │");
            System.Diagnostics.Debug.WriteLine($"   │ 📋 Distribution breakdown:");

            // Apply discounts with detailed logging
            for (int i = 0; i < items.Count; i++)
            {
                items[i].ApplicableDiscount += discounts[i];

                decimal itemAmount = Math.Abs(Convert.ToDecimal(items[i].Amount));
                decimal proportion = itemAmount / totalAmount;

                System.Diagnostics.Debug.WriteLine(
                    $"   │   • {items[i].ItemName,-35} " +
                    $"Amt: {itemAmount,8:C}  " +
                    $"({proportion,5:P1})  " +
                    $"Disc: {discounts[i],7:C}");
            }

            System.Diagnostics.Debug.WriteLine($"   └─────────────────────────────────────────────────────\n");
        }

        private string GetDiscountType(LineItemContext discount, List<LineItemContext> affectedItems)
        {
            if (affectedItems == null || affectedItems.Count == 0)
                return "Unknown Discount";

            decimal totalAmount = affectedItems.Sum(x => Math.Abs(Convert.ToDecimal(x.Amount)));
            decimal discountAmount = Math.Abs(Convert.ToDecimal(discount.Amount));

            if (totalAmount == 0)
                return "Unknown Discount";

            decimal percentage = (discountAmount / totalAmount) * 100;

            // Check if it's a round percentage (within 0.1%)
            decimal roundedPercentage = Math.Round(percentage, 0);
            if (Math.Abs(percentage - roundedPercentage) < 0.1m)
            {
                return $"~{roundedPercentage}% Discount";
            }

            return $"Fixed {discountAmount:C} Discount";
        }


        /// <summary>
        /// Robust heuristic to determine whether an invoice line is a subtotal row (QuickBooks can represent these in different ways).
        /// Returns true if the line appears to be a subtotal (should not be exported as a normal item).
        /// </summary>
        private bool IsLikelySubtotal(IInvoiceLineRet line, string itemListID, string itemType, string itemName, double lineAmount)
        {
            try
            {
                string type = itemType?.Trim() ?? "";
                string name = itemName?.Trim() ?? "";

                // Common textual markers
                if (!string.IsNullOrEmpty(type) && type.IndexOf("subtotal", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                if (!string.IsNullOrEmpty(name) && name.IndexOf("subtotal", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                if (!string.IsNullOrEmpty(type) && type.IndexOf("sub-total", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                if (!string.IsNullOrEmpty(name) && name.IndexOf("sub-total", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;

                // QuickBooks sometimes uses empty ItemRef and description for subtotal lines
                bool itemRefEmpty = string.IsNullOrEmpty(itemListID) || itemListID.Trim() == "0";
                bool descEmpty = string.IsNullOrWhiteSpace(name);

                // Quantity missing or zero is a good indicator the line isn't a normal sale item
                bool quantityMissing = (line.Quantity == null || Math.Abs((double)(line.Quantity?.GetValue() ?? 0)) == 0);

                if (itemRefEmpty && (descEmpty || string.IsNullOrEmpty(type)) && lineAmount > 0 && quantityMissing)
                    return true;

                // If ItemRef exists but it is explicitly named "Subtotal"
                if (!string.IsNullOrEmpty(itemListID) && !string.IsNullOrEmpty(type) && type.IndexOf("subtotal", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;

                // Fallback - not a subtotal
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"IsLikelySubtotal error: {ex.Message}");
                return false;
            }
        }

        public void CloseSession()
        {
            if (_sessionManager != null && _isConnected)
            {
                try
                {
                    _sessionManager.EndSession();
                    _isConnected = false;
                    System.Diagnostics.Debug.WriteLine("✅ QuickBooks session closed successfully");
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
                    System.Diagnostics.Debug.WriteLine("✅ QuickBooks connection closed successfully");
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

    }

    // Helper classes to store data
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
        public string CustomerType { get; set; } = "";
        public string Address { get; set; } = "";
        public string State { get; set; } = "";
        public string Phone { get; set; } = "";
    }

    internal class ItemData
    {
        public string HSCode { get; set; } = "";
        public string RetailPrice { get; set; } = "";
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
}