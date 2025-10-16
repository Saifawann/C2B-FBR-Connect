using C2B_FBR_Connect.Models;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using QBFC16Lib;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Services
{
    public class QuickBooksService : IDisposable
    {
        private QBSessionManager _sessionManager;
        private bool _isConnected;
        private CompanyInfo _companyInfo;
        private Company _companySettings;

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

                // Get company info
                FetchCompanyInfo();

                _isConnected = true;
                return true;
            }
            catch (Exception ex)
            {
                _isConnected = false;
                throw new Exception($"QuickBooks connection failed: {ex.Message}", ex);
            }
        }

        private void FetchCompanyInfo()
        {
            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", 16, 0);
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
                        Phone = company.Email?.GetValue(),
                        Email = company.Phone?.GetValue()
                    };

                    CurrentCompanyName = _companyInfo.Name;
                    CurrentCompanyFile = _sessionManager.GetCurrentCompanyFileName();
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
            if (!_isConnected) Connect();

            var invoices = new List<Invoice>();

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", 16, 0);
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
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to fetch all invoices: {ex.Message}", ex);
            }

            return invoices;
        }

        public FBRInvoicePayload GetInvoiceDetails(string qbInvoiceId)
        {
            if (!_isConnected) Connect();

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", 16, 0);
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

                        // Add line items with FBR required fields
                        if (inv.ORInvoiceLineRetList != null)
                        {
                            for (int i = 0; i < inv.ORInvoiceLineRetList.Count; i++)
                            {
                                var lineRet = inv.ORInvoiceLineRetList.GetAt(i);

                                if (lineRet.InvoiceLineRet != null)
                                {
                                    var line = lineRet.InvoiceLineRet;
                                    string itemListID = line.ItemRef?.ListID?.GetValue() ?? "";

                                    // Fetch item details for HS Code and Retail Price
                                    var itemData = FetchItemDetails(itemListID);
                                    string hsCode = itemData.HSCode;
                                    string retailPrice = itemData.RetailPrice;

                                    // Debug log to check retail price
                                    System.Diagnostics.Debug.WriteLine($"Item: {line.Desc?.GetValue()}, Retail Price from FetchItemDetails: '{retailPrice}'");

                                    // Check line-level custom fields as well
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
                                    double lineAmount = line.Amount?.GetValue() ?? 0;
                                    double actualTaxRate = taxRate;
                                    double salesTaxAmount = lineAmount * (taxRate / 100);

                                    decimal displayTaxRate;
                                    decimal standardTaxAmount;
                                    decimal furtherTax;

                                    if (actualTaxRate > 18)
                                    {
                                        displayTaxRate = 18;
                                        standardTaxAmount = Convert.ToDecimal(lineAmount * 0.18);
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
                                        UnitOfMeasure = line.UnitOfMeasure?.GetValue() ?? "",
                                        UnitPrice = Convert.ToDecimal(line.Amount?.GetValue() ?? 0),
                                        TotalPrice = Convert.ToDecimal(lineAmount),
                                        TaxRate = displayTaxRate,
                                        SalesTaxAmount = standardTaxAmount,
                                        TotalValue = Convert.ToDecimal(lineAmount) + standardTaxAmount + furtherTax,
                                        RetailPrice = decimal.TryParse(retailPrice, out var parsedRetailPrice) ? parsedRetailPrice : 0,
                                        ExtraTax = furtherTax,
                                        FurtherTax = 0
                                    };

                                    payload.Items.Add(item);
                                }
                                else if (lineRet.InvoiceLineGroupRet != null)
                                {
                                    var lineGroup = lineRet.InvoiceLineGroupRet;
                                    string hsCode = "";

                                    // Check for HS Code in line group custom fields
                                    if (lineGroup.DataExtRetList != null && lineGroup.DataExtRetList.Count > 0)
                                    {
                                        for (int k = 0; k < lineGroup.DataExtRetList.Count; k++)
                                        {
                                            var dataExt = lineGroup.DataExtRetList.GetAt(k);
                                            if (dataExt.DataExtName?.GetValue()?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true)
                                            {
                                                hsCode = dataExt.DataExtValue?.GetValue();
                                            }
                                        }
                                    }

                                    var item = new InvoiceItem
                                    {
                                        ItemName = lineGroup.Desc?.GetValue() ?? "",
                                        HSCode = hsCode,
                                        Quantity = Convert.ToInt32(lineGroup.Quantity?.GetValue() ?? 1),
                                        TotalPrice = Convert.ToDecimal(lineGroup.TotalAmount?.GetValue() ?? 0)
                                    };

                                    payload.Items.Add(item);
                                }
                            }
                        }

                        return payload;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to get invoice details: {ex.Message}", ex);
            }

            return null;
        }

        private CustomerData FetchCustomerDetails(string customerListID)
        {
            var customerData = new CustomerData();

            if (string.IsNullOrEmpty(customerListID)) return customerData;

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", 16, 0);
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
                        customerData.CustomerType = customer.CustomerTypeRef?.FullName?.GetValue() ?? "UnRegistered";

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
                                     fieldName?.Equals("NTN_CNIC", StringComparison.OrdinalIgnoreCase) == true) &&
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
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", 16, 0);
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
                                // Note: We'd need the base price here for percentage calculation
                                // This might be why retail price is missing - we need the base price
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
            if (!_isConnected) Connect();

            var priceLevels = new List<PriceLevel>();

            try
            {
                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", 16, 0);
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
                    level.FixedPercentage = 0; // Add this property to your PriceLevel class

                    object orPriceLevelObj = priceLevel.ORPriceLevelRet;
                    if (orPriceLevelObj != null)
                    {
                        var type = orPriceLevelObj.GetType();

                        // 🧩 Handle Per-Item Price Levels
                        var perItemProp = type.GetProperty("PriceLevelPerItemRetList");
                        if (perItemProp != null)
                        {
                            var itemList = perItemProp.GetValue(orPriceLevelObj);
                            if (itemList != null)
                            {
                                int count = (int)itemList.GetType().GetProperty("Count").GetValue(itemList);
                                for (int j = 0; j < count; j++)
                                {
                                    var itemPrice = itemList.GetType().GetMethod("GetAt").Invoke(itemList, new object[] { j });

                                    string itemListID = null;
                                    string itemFullName = null;
                                    decimal customPrice = 0;
                                    double customPricePercent = 0;

                                    try
                                    {
                                        var itemRefProp = itemPrice.GetType().GetProperty("ItemRef");
                                        var itemRef = itemRefProp?.GetValue(itemPrice);
                                        if (itemRef != null)
                                        {
                                            var listIDObj = itemRef.GetType().GetProperty("ListID")?.GetValue(itemRef);
                                            if (listIDObj != null)
                                            {
                                                var valMethod = listIDObj.GetType().GetMethod("GetValue");
                                                if (valMethod != null)
                                                    itemListID = valMethod.Invoke(listIDObj, null)?.ToString();
                                            }

                                            var fullNameObj = itemRef.GetType().GetProperty("FullName")?.GetValue(itemRef);
                                            if (fullNameObj != null)
                                            {
                                                var valMethod = fullNameObj.GetType().GetMethod("GetValue");
                                                if (valMethod != null)
                                                    itemFullName = valMethod.Invoke(fullNameObj, null)?.ToString();
                                            }
                                        }

                                        // Extract price fields
                                        var customPriceObj = itemPrice.GetType().GetProperty("CustomPrice")?.GetValue(itemPrice);
                                        if (customPriceObj != null)
                                        {
                                            var valMethod = customPriceObj.GetType().GetMethod("GetValue");
                                            if (valMethod != null)
                                                customPrice = Convert.ToDecimal(valMethod.Invoke(customPriceObj, null));
                                        }

                                        var customPercentObj = itemPrice.GetType().GetProperty("CustomPricePercent")?.GetValue(itemPrice);
                                        if (customPercentObj != null)
                                        {
                                            var valMethod = customPercentObj.GetType().GetMethod("GetValue");
                                            if (valMethod != null)
                                                customPricePercent = Convert.ToDouble(valMethod.Invoke(customPercentObj, null));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error reading price level item: {ex.Message}");
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

                        // 🧩 Handle Fixed Percentage Price Levels (NEW - This is what you're missing)
                        var fixedPctProp = type.GetProperty("PriceLevelFixedPercentage");
                        if (fixedPctProp != null)
                        {
                            var fixedPctObj = fixedPctProp.GetValue(orPriceLevelObj);
                            if (fixedPctObj != null)
                            {
                                var pctValueProp = fixedPctObj.GetType().GetProperty("Value");
                                if (pctValueProp != null)
                                {
                                    var pctVal = pctValueProp.GetValue(fixedPctObj);
                                    if (pctVal != null)
                                    {
                                        level.FixedPercentage = Convert.ToDouble(pctVal);
                                        System.Diagnostics.Debug.WriteLine($"✅ Fixed percentage level '{level.Name}': {level.FixedPercentage}%");
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
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to fetch price levels: {ex.Message}", ex);
            }

            return priceLevels;
        }

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
        public string Phone { get; set; }      // Add this
        public string Email { get; set; }       // Add this
    }

    internal class CustomerData
    {
        public string NTN { get; set; } = "";
        public string CustomerType { get; set; } = "";
        public string Address { get; set; } = "";
        public string State { get; set; } = "";
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
        public double FixedPercentage { get; set; } // Add this
    }

    public class PriceLevelItem
    {
        public string ItemListID { get; set; }
        public string ItemFullName { get; set; }
        public decimal CustomPrice { get; set; }
        public double CustomPricePercent { get; set; }
    }
}