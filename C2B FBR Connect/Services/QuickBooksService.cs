using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using QBXMLRP2Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace C2B_FBR_Connect.Services
{
    public class QuickBooksService : IDisposable
    {
        private RequestProcessor2 _rp;
        private string _ticket;
        private bool _isConnected;
        private CompanyInfo _companyInfo;
        private Company _companySettings;
        private CompanyManager _companyManager;
        private FBRApiService _fbr;

        public string CurrentCompanyName { get; private set; }
        public string CurrentCompanyFile { get; private set; }

        public bool Connect(Company companySettings = null)
        {
            try
            {
                _companySettings = companySettings;

                _rp = new RequestProcessor2();
                _rp.OpenConnection2("", "C2B Smart App", QBXMLRPConnectionType.localQBD);
                _ticket = _rp.BeginSession("", QBFileMode.qbFileOpenDoNotCare);

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

        private void FetchCompanyInfo()
        {
            try
            {
                string requestXML = BuildCompanyQueryRequest();
                string responseXML = _rp.ProcessRequest(_ticket, requestXML);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseXML);

                XmlNode companyRet = doc.SelectSingleNode("//CompanyRet");
                if (companyRet != null)
                {
                    _companyInfo = new CompanyInfo
                    {
                        Name = GetXmlValue(companyRet, "CompanyName"),
                        City = GetXmlValue(companyRet, "LegalAddress/City"),
                        State = GetXmlValue(companyRet, "LegalAddress/State"),
                        PostalCode = GetXmlValue(companyRet, "LegalAddress/PostalCode"),
                        Country = GetXmlValue(companyRet, "LegalAddress/Country"),
                        Phone = GetXmlValue(companyRet, "Phone"),
                        Email = GetXmlValue(companyRet, "Email")
                    };

                    _companyInfo.Address = FormatAddressFromXml(companyRet.SelectSingleNode("LegalAddress"));

                    // Get NTN from custom fields
                    XmlNodeList dataExtList = companyRet.SelectNodes("DataExtRet");
                    foreach (XmlNode dataExt in dataExtList)
                    {
                        string fieldName = GetXmlValue(dataExt, "DataExtName");
                        string fieldValue = GetXmlValue(dataExt, "DataExtValue");

                        if (fieldName?.Equals("NTN", StringComparison.OrdinalIgnoreCase) == true ||
                            fieldName?.Equals("CNIC", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            _companyInfo.NTN = fieldValue;
                        }
                    }

                    CurrentCompanyName = _companyInfo.Name;
                    CurrentCompanyFile = _rp.GetCurrentCompanyFileName(_ticket);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Warning: Could not fetch company information: {ex.Message}");
            }
        }

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
                string requestXML = BuildInvoiceQueryRequest(null, true);
                string responseXML = _rp.ProcessRequest(_ticket, requestXML);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseXML);

                XmlNodeList invoiceList = doc.SelectNodes("//InvoiceRet");

                foreach (XmlNode inv in invoiceList)
                {
                    string customerName = GetXmlValue(inv, "CustomerRef/FullName");
                    string customerListID = GetXmlValue(inv, "CustomerRef/ListID");
                    string txnID = GetXmlValue(inv, "TxnID");
                    string refNumber = GetXmlValue(inv, "RefNumber");
                    string txnDate = GetXmlValue(inv, "TxnDate");

                    var customerData = FetchCustomerDetails(customerListID);

                    var invoice = new Invoice
                    {
                        CompanyName = CurrentCompanyName,
                        QuickBooksInvoiceId = txnID,
                        InvoiceNumber = refNumber,
                        CustomerName = customerName,
                        CustomerNTN = customerData.NTN,
                        Amount = Convert.ToDecimal(ParseDouble(GetXmlValue(inv, "Subtotal"))),
                        Status = "Pending",
                        CreatedDate = string.IsNullOrEmpty(txnDate) ? DateTime.Now : DateTime.Parse(txnDate)
                    };

                    invoices.Add(invoice);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to fetch all invoices: {ex.Message}", ex);
            }

            return invoices;
        }

        public async System.Threading.Tasks.Task<FBRInvoicePayload> GetInvoiceDetails(string qbInvoiceId)
        {
            _fbr = new FBRApiService();
            if (!_isConnected) Connect();

            try
            {
                string requestXML = BuildInvoiceQueryRequest(qbInvoiceId, true);
                string responseXML = _rp.ProcessRequest(_ticket, requestXML);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseXML);

                XmlNode inv = doc.SelectSingleNode("//InvoiceRet");
                if (inv != null)
                {
                    string customerListID = GetXmlValue(inv, "CustomerRef/ListID");
                    var customerData = FetchCustomerDetails(customerListID);

                    string invoiceDate = GetXmlValue(inv, "TxnDate");
                    DateTime txnDate = string.IsNullOrEmpty(invoiceDate) ? DateTime.Now : DateTime.Parse(invoiceDate);

                    double subtotal = ParseDouble(GetXmlValue(inv, "Subtotal"));
                    double salesTaxTotal = ParseDouble(GetXmlValue(inv, "SalesTaxTotal"));
                    double taxRate = ParseDouble(GetXmlValue(inv, "SalesTaxPercentage"));

                    var payload = new FBRInvoicePayload
                    {
                        InvoiceType = "Sale Invoice",
                        InvoiceNumber = GetXmlValue(inv, "RefNumber"),
                        InvoiceDate = txnDate,

                        SellerNTN = _companySettings?.SellerNTN ?? _companyInfo?.NTN ?? "",
                        SellerBusinessName = _companyInfo?.Name ?? CurrentCompanyName,
                        SellerProvince = _companySettings?.SellerProvince ?? "",
                        SellerAddress = _companySettings?.SellerAddress ?? _companyInfo?.Address ?? "",

                        CustomerName = GetXmlValue(inv, "CustomerRef/FullName"),
                        CustomerNTN = customerData.NTN,
                        BuyerProvince = customerData.State,
                        BuyerAddress = customerData.Address,
                        BuyerRegistrationType = customerData.CustomerType,
                        BuyerPhone = customerData.Phone,

                        TotalAmount = Convert.ToDecimal(subtotal + salesTaxTotal),
                        Subtotal = Convert.ToDecimal(subtotal),
                        TaxAmount = Convert.ToDecimal(salesTaxTotal),

                        Items = new List<InvoiceItem>()
                    };

                    // Process line items
                    XmlNodeList lineItems = inv.SelectNodes("InvoiceLineRet");
                    foreach (XmlNode line in lineItems)
                    {
                        string itemListID = GetXmlValue(line, "ItemRef/ListID");
                        var itemData = FetchItemDetails(itemListID);

                        string hsCode = itemData.HSCode;
                        string retailPrice = itemData.RetailPrice;

                        System.Diagnostics.Debug.WriteLine($"Item: {GetXmlValue(line, "Desc")}, Retail Price from FetchItemDetails: '{retailPrice}'");

                        // Check line-level custom fields
                        XmlNodeList lineDataExt = line.SelectNodes("DataExtRet");
                        foreach (XmlNode dataExt in lineDataExt)
                        {
                            string fieldName = GetXmlValue(dataExt, "DataExtName");
                            string fieldValue = GetXmlValue(dataExt, "DataExtValue");

                            if (fieldName?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                            {
                                hsCode = fieldValue;
                            }
                            else if (fieldName?.Equals("Residential", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                            {
                                retailPrice = fieldValue;
                            }
                        }

                        double lineAmount = ParseDouble(GetXmlValue(line, "Amount"));
                        double quantity = ParseDouble(GetXmlValue(line, "Quantity"));
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
                            ItemName = GetXmlValue(line, "Desc"),
                            HSCode = hsCode,
                            Quantity = Convert.ToInt32(quantity),
                            UnitOfMeasure = await _fbr.GetUOMDescriptionAsync(hsCode, _companySettings.FBRToken) ?? GetXmlValue(line, "UnitOfMeasure"),
                            UnitPrice = Convert.ToDecimal(lineAmount),
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

                    // Process line groups
                    XmlNodeList lineGroups = inv.SelectNodes("InvoiceLineGroupRet");
                    foreach (XmlNode lineGroup in lineGroups)
                    {
                        string hsCode = "";
                        XmlNodeList groupDataExt = lineGroup.SelectNodes("DataExtRet");
                        foreach (XmlNode dataExt in groupDataExt)
                        {
                            if (GetXmlValue(dataExt, "DataExtName")?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                hsCode = GetXmlValue(dataExt, "DataExtValue");
                            }
                        }

                        var item = new InvoiceItem
                        {
                            ItemName = GetXmlValue(lineGroup, "Desc"),
                            HSCode = hsCode,
                            Quantity = Convert.ToInt32(ParseDouble(GetXmlValue(lineGroup, "Quantity"))),
                            TotalPrice = Convert.ToDecimal(ParseDouble(GetXmlValue(lineGroup, "TotalAmount")))
                        };

                        payload.Items.Add(item);
                    }

                    return payload;
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
                string requestXML = BuildCustomerQueryRequest(customerListID);
                string responseXML = _rp.ProcessRequest(_ticket, requestXML);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseXML);

                XmlNode customer = doc.SelectSingleNode("//CustomerRet");
                if (customer != null)
                {
                    customerData.Address = FormatAddressFromXml(customer.SelectSingleNode("BillAddress"));
                    customerData.State = GetXmlValue(customer, "BillAddress/State");
                    customerData.CustomerType = GetXmlValue(customer, "CustomerTypeRef/FullName");
                    if (string.IsNullOrEmpty(customerData.CustomerType))
                        customerData.CustomerType = "Unregistered";

                    customerData.Phone = GetXmlValue(customer, "Phone");
                    if (string.IsNullOrEmpty(customerData.Phone))
                        customerData.Phone = "N/A";

                    // Read custom fields
                    XmlNodeList dataExtList = customer.SelectNodes("DataExtRet");
                    foreach (XmlNode dataExt in dataExtList)
                    {
                        string fieldName = GetXmlValue(dataExt, "DataExtName");
                        string fieldValue = GetXmlValue(dataExt, "DataExtValue");

                        if ((fieldName?.Equals("NTN", StringComparison.OrdinalIgnoreCase) == true ||
                             fieldName?.Equals("CNIC", StringComparison.OrdinalIgnoreCase) == true ||
                             fieldName?.Equals("NTN_CNIC", StringComparison.OrdinalIgnoreCase) == true) &&
                            !string.IsNullOrEmpty(fieldValue))
                        {
                            customerData.NTN = fieldValue;
                        }

                        if (fieldName?.Equals("Province", StringComparison.OrdinalIgnoreCase) == true &&
                            !string.IsNullOrEmpty(fieldValue))
                        {
                            customerData.State = fieldValue;
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
                string requestXML = BuildItemQueryRequest(itemListID);
                string responseXML = _rp.ProcessRequest(_ticket, requestXML);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseXML);

                // Check for different item types
                XmlNode item = doc.SelectSingleNode("//ItemServiceRet | //ItemInventoryRet | //ItemNonInventoryRet | //ItemOtherChargeRet | //ItemInventoryAssemblyRet | //ItemGroupRet");

                if (item != null)
                {
                    string itemName = GetXmlValue(item, "Name");
                    System.Diagnostics.Debug.WriteLine($"Fetching details for item: {itemName} (ID: {itemListID})");

                    bool retailPriceFound = false;
                    XmlNodeList customFields = item.SelectNodes("DataExtRet");

                    if (customFields != null && customFields.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"  Checking {customFields.Count} custom fields...");

                        foreach (XmlNode dataExt in customFields)
                        {
                            string fieldName = GetXmlValue(dataExt, "DataExtName");
                            string fieldValue = GetXmlValue(dataExt, "DataExtValue");

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
                var allPriceLevels = FetchPriceLevels();

                System.Diagnostics.Debug.WriteLine($"    Found {allPriceLevels.Count} price levels total");

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

        public List<PriceLevel> FetchPriceLevels()
        {
            if (!_isConnected) Connect();

            var priceLevels = new List<PriceLevel>();

            try
            {
                string requestXML = BuildPriceLevelQueryRequest();
                string responseXML = _rp.ProcessRequest(_ticket, requestXML);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseXML);

                XmlNodeList priceLevelNodes = doc.SelectNodes("//PriceLevelRet");

                foreach (XmlNode plNode in priceLevelNodes)
                {
                    var level = new PriceLevel
                    {
                        ListID = GetXmlValue(plNode, "ListID"),
                        Name = GetXmlValue(plNode, "Name"),
                        IsActive = GetXmlValue(plNode, "IsActive") == "true",
                        PriceLevelType = GetXmlValue(plNode, "PriceLevelType")
                    };

                    level.Items = new List<PriceLevelItem>();
                    level.FixedPercentage = 0;

                    // Check for fixed percentage
                    string fixedPct = GetXmlValue(plNode, "PriceLevelFixedPercentage");
                    if (!string.IsNullOrEmpty(fixedPct))
                    {
                        level.FixedPercentage = ParseDouble(fixedPct);
                        System.Diagnostics.Debug.WriteLine($"✅ Fixed percentage level '{level.Name}': {level.FixedPercentage}%");
                    }

                    // Check for per-item pricing
                    XmlNodeList perItemNodes = plNode.SelectNodes("PriceLevelPerItem");
                    foreach (XmlNode perItem in perItemNodes)
                    {
                        string itemListID = GetXmlValue(perItem, "ItemRef/ListID");
                        string itemFullName = GetXmlValue(perItem, "ItemRef/FullName");

                        decimal customPrice = 0;
                        double customPricePercent = 0;

                        string priceStr = GetXmlValue(perItem, "CustomPrice");
                        if (!string.IsNullOrEmpty(priceStr))
                        {
                            customPrice = Convert.ToDecimal(ParseDouble(priceStr));
                        }

                        string percentStr = GetXmlValue(perItem, "CustomPricePercent");
                        if (!string.IsNullOrEmpty(percentStr))
                        {
                            customPricePercent = ParseDouble(percentStr);
                        }

                        level.Items.Add(new PriceLevelItem
                        {
                            ItemListID = itemListID,
                            ItemFullName = itemFullName,
                            CustomPrice = customPrice,
                            CustomPricePercent = customPricePercent
                        });
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

        // XML Request Builders
        private string BuildCompanyQueryRequest()
        {
            return @"<?xml version=""1.0"" encoding=""utf-8""?>
<?qbxml version=""13.0""?>
<QBXML>
    <QBXMLMsgsRq onError=""continueOnError"">
        <CompanyQueryRq requestID=""1"">
            <OwnerID>0</OwnerID>
        </CompanyQueryRq>
    </QBXMLMsgsRq>
</QBXML>";
        }

        private string BuildInvoiceQueryRequest(string txnID, bool includeLineItems)
        {
            string txnIDFilter = string.IsNullOrEmpty(txnID) ? "" : $"<TxnID>{txnID}</TxnID>";
            string includeLines = includeLineItems ? "<IncludeLineItems>true</IncludeLineItems>" : "";

            return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<?qbxml version=""13.0""?>
<QBXML>
    <QBXMLMsgsRq onError=""continueOnError"">
        <InvoiceQueryRq requestID=""1"">
            {txnIDFilter}
            {includeLines}
            <OwnerID>0</OwnerID>
        </InvoiceQueryRq>
    </QBXMLMsgsRq>
</QBXML>";
        }

        private string BuildCustomerQueryRequest(string listID)
        {
            return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<?qbxml version=""13.0""?>
<QBXML>
    <QBXMLMsgsRq onError=""continueOnError"">
        <CustomerQueryRq requestID=""1"">
            <ListID>{listID}</ListID>
            <OwnerID>0</OwnerID>
        </CustomerQueryRq>
    </QBXMLMsgsRq>
</QBXML>";
        }

        private string BuildItemQueryRequest(string listID)
        {
            return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<?qbxml version=""13.0""?>
<QBXML>
    <QBXMLMsgsRq onError=""continueOnError"">
        <ItemQueryRq requestID=""1"">
            <ListID>{listID}</ListID>
            <OwnerID>0</OwnerID>
        </ItemQueryRq>
    </QBXMLMsgsRq>
</QBXML>";
        }

        private string BuildPriceLevelQueryRequest()
        {
            return @"<?xml version=""1.0"" encoding=""utf-8""?>
<?qbxml version=""13.0""?>
<QBXML>
    <QBXMLMsgsRq onError=""continueOnError"">
        <PriceLevelQueryRq requestID=""1"">
        </PriceLevelQueryRq>
    </QBXMLMsgsRq>
</QBXML>";
        }

        // Helper Methods
        private string GetXmlValue(XmlNode node, string xpath)
        {
            if (node == null) return "";
            XmlNode childNode = node.SelectSingleNode(xpath);
            return childNode?.InnerText ?? "";
        }

        private double ParseDouble(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            double result;
            double.TryParse(value, out result);
            return result;
        }

        private string FormatAddressFromXml(XmlNode addressNode)
        {
            if (addressNode == null) return "";

            List<string> lines = new List<string>();

            for (int i = 1; i <= 5; i++)
            {
                string addr = GetXmlValue(addressNode, $"Addr{i}");
                if (!string.IsNullOrEmpty(addr))
                    lines.Add(addr.TrimEnd(',', ' '));
            }

            string city = GetXmlValue(addressNode, "City");
            string state = GetXmlValue(addressNode, "State");
            string postalCode = GetXmlValue(addressNode, "PostalCode");
            string country = GetXmlValue(addressNode, "Country");

            if (!string.IsNullOrEmpty(city)) lines.Add(city.TrimEnd(',', ' '));
            if (!string.IsNullOrEmpty(state)) lines.Add(state.TrimEnd(',', ' '));
            if (!string.IsNullOrEmpty(postalCode)) lines.Add(postalCode.TrimEnd(',', ' '));
            if (!string.IsNullOrEmpty(country)) lines.Add(country.TrimEnd(',', ' '));

            return string.Join(", ", lines);
        }

        public void CloseSession()
        {
            if (_rp != null && _isConnected && !string.IsNullOrEmpty(_ticket))
            {
                try
                {
                    _rp.EndSession(_ticket);
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
            if (_rp != null)
            {
                try
                {
                    _rp.CloseConnection();
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
                if (_rp != null)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_rp);
                    }
                    catch { }

                    _rp = null;
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
}