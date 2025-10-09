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
        private Company _companySettings; // Store company settings from database

        public string CurrentCompanyName { get; private set; }
        public string CurrentCompanyFile { get; private set; }

        public bool Connect(Company companySettings = null)
        {
            try
            {
                _companySettings = companySettings;

                _sessionManager = new QBSessionManager();
                _sessionManager.OpenConnection("", "FBR Digital Invoicing App");
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
                        Country = company.LegalAddress?.Country?.GetValue()
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

        /// <summary>
        /// Gets the current QuickBooks company information
        /// </summary>
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

                // ✅ Do not apply any date filter — fetch all invoices
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

                            // Seller Information - Use settings from database first, then fallback to QuickBooks
                            SellerNTN = _companySettings?.SellerNTN ?? _companyInfo?.NTN ?? "",
                            SellerBusinessName = _companyInfo?.Name ?? CurrentCompanyName,
                            SellerProvince = _companySettings?.SellerProvince ?? "",
                            SellerAddress = _companySettings?.SellerAddress ?? _companyInfo?.Address ?? "",

                            // Buyer Information (Customer)
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

                                    // Check line-level custom fields as well
                                    string hsCode = itemData.HSCode;
                                    string retailPrice = itemData.RetailPrice;

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
                                            else if (fieldName?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(fieldValue))
                                            {
                                                retailPrice = fieldValue;
                                            }
                                        }
                                    }

                                    // Calculate line-level tax

                                    double lineAmount = line.Amount?.GetValue() ?? 0;
                                    double actualTaxRate = taxRate; // Original tax rate from QB
                                    double salesTaxAmount = lineAmount * (taxRate / 100);

                                    // ✅ Split tax logic: If rate > 18%, split into standard tax and extra tax
                                    decimal displayTaxRate;
                                    decimal standardTaxAmount;
                                    decimal furtherTax;

                                    if (actualTaxRate > 18)
                                    {
                                        // Tax rate is above 18% - split it
                                        displayTaxRate = 18; // Show only 18% as the rate
                                        standardTaxAmount = Convert.ToDecimal(lineAmount * 0.18); // Calculate 18% of line amount
                                        furtherTax = Convert.ToDecimal(salesTaxAmount) - standardTaxAmount; // Remaining goes to extra tax

                                        System.Diagnostics.Debug.WriteLine($"Tax split - Original: {actualTaxRate}%, Amount: {salesTaxAmount}");
                                        System.Diagnostics.Debug.WriteLine($"  Standard (18%): {standardTaxAmount}, Extra: {furtherTax}");
                                    }
                                    else
                                    {
                                        // Tax rate is 18% or below - use as is
                                        displayTaxRate = Convert.ToDecimal(actualTaxRate);
                                        standardTaxAmount = Convert.ToDecimal(salesTaxAmount);
                                        furtherTax = 0;
                                    }

                                    var item = new InvoiceItem
                                    {
                                        ItemName = line.Desc?.GetValue() ?? "",
                                        HSCode = hsCode,
                                        Quantity = Convert.ToInt32(line.Quantity?.GetValue() ?? 1),
                                        UnitOfMeasure = line.UnitOfMeasure?.GetValue() ?? "Numbers, pieces, units",
                                        UnitPrice = Convert.ToDecimal(line.Amount?.GetValue() ?? 0),
                                        TotalPrice = Convert.ToDecimal(lineAmount),
                                        TaxRate = displayTaxRate, // Show 18% max
                                        SalesTaxAmount = standardTaxAmount, // Standard tax (max 18%)
                                        TotalValue = Convert.ToDecimal(lineAmount) + standardTaxAmount + furtherTax, // Total includes both
                                        RetailPrice = retailPrice,
                                        ExtraTax = furtherTax, // Amount above 18%
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

                        // Get address information (from Addr1, Addr2, etc.)
                        customerData.Address = FormatAddress(customer.BillAddress);

                        // Get customer type from QuickBooks
                        customerData.CustomerType = customer.CustomerTypeRef?.FullName?.GetValue() ?? "UnRegistered";

                        // ✅ READ CUSTOM FIELDS (NTN and Province)
                        if (customer.DataExtRetList != null && customer.DataExtRetList.Count > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"Customer: {customer.Name?.GetValue()} - Custom fields count: {customer.DataExtRetList.Count}");

                            for (int i = 0; i < customer.DataExtRetList.Count; i++)
                            {
                                var dataExt = customer.DataExtRetList.GetAt(i);
                                string fieldName = dataExt.DataExtName?.GetValue();
                                string fieldValue = dataExt.DataExtValue?.GetValue();

                                System.Diagnostics.Debug.WriteLine($"  Custom field: {fieldName} = '{fieldValue}'");

                                // Check for NTN/CNIC
                                if ((fieldName?.Equals("NTN", StringComparison.OrdinalIgnoreCase) == true ||
                                     fieldName?.Equals("CNIC", StringComparison.OrdinalIgnoreCase) == true ||
                                     fieldName?.Equals("NTN_CNIC", StringComparison.OrdinalIgnoreCase) == true) &&
                                    !string.IsNullOrEmpty(fieldValue))
                                {
                                    customerData.NTN = fieldValue;
                                }

                                // ✅ Check for Province custom field
                                if (fieldName?.Equals("Province", StringComparison.OrdinalIgnoreCase) == true &&
                                    !string.IsNullOrEmpty(fieldValue))
                                {
                                    customerData.State = fieldValue;
                                }
                            }
                        }

                        System.Diagnostics.Debug.WriteLine($"  Final NTN: '{customerData.NTN}'");
                        System.Diagnostics.Debug.WriteLine($"  Final Province: '{customerData.State}'");
                        System.Diagnostics.Debug.WriteLine($"  Final Address: '{customerData.Address}'");
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
                        string itemType = "Unknown";

                        // Check which type of item it is and get custom fields
                        if (itemRet.ItemServiceRet != null)
                        {
                            customFields = itemRet.ItemServiceRet.DataExtRetList;
                            itemType = "Service";
                        }
                        else if (itemRet.ItemInventoryRet != null)
                        {
                            customFields = itemRet.ItemInventoryRet.DataExtRetList;
                            itemType = "Inventory";
                        }
                        else if (itemRet.ItemNonInventoryRet != null)
                        {
                            customFields = itemRet.ItemNonInventoryRet.DataExtRetList;
                            itemType = "NonInventory";
                        }
                        else if (itemRet.ItemOtherChargeRet != null)
                        {
                            customFields = itemRet.ItemOtherChargeRet.DataExtRetList;
                            itemType = "OtherCharge";
                        }
                        else if (itemRet.ItemInventoryAssemblyRet != null)
                        {
                            // ✅ INVENTORY ASSEMBLY SUPPORT
                            customFields = itemRet.ItemInventoryAssemblyRet.DataExtRetList;
                            itemType = "InventoryAssembly";
                            System.Diagnostics.Debug.WriteLine($"Item Type: Inventory Assembly");
                        }
                        else if (itemRet.ItemGroupRet != null)
                        {
                            customFields = itemRet.ItemGroupRet.DataExtRetList;
                            itemType = "Group";
                        }

                        System.Diagnostics.Debug.WriteLine($"Fetching item details - Type: {itemType}");

                        // Extract HS Code from custom fields
                        if (customFields != null && customFields.Count > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"  Custom fields found: {customFields.Count}");

                            for (int i = 0; i < customFields.Count; i++)
                            {
                                var dataExt = customFields.GetAt(i);
                                string fieldName = dataExt.DataExtName?.GetValue();
                                string fieldValue = dataExt.DataExtValue?.GetValue();

                                System.Diagnostics.Debug.WriteLine($"  Custom field: {fieldName} = '{fieldValue}'");

                                if (fieldName?.Equals("HS Code", StringComparison.OrdinalIgnoreCase) == true &&
                                    !string.IsNullOrEmpty(fieldValue))
                                {
                                    itemData.HSCode = fieldValue;
                                    System.Diagnostics.Debug.WriteLine($"  HS Code found: {fieldValue}");
                                }

                                if (fieldName?.Equals("Retail Price", StringComparison.OrdinalIgnoreCase) == true &&
                                    !string.IsNullOrEmpty(fieldValue))
                                {
                                    itemData.RetailPrice = fieldValue;
                                    System.Diagnostics.Debug.WriteLine($"  Retail Price found: {fieldValue}");
                                }
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"  No custom fields found for {itemType} item");
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

        private string FormatAddress(IAddress address)
        {
            if (address == null) return "";

            try
            {
                var addr = address as dynamic;
                var lines = new List<string>();

                // Add all address lines that have data
                if (addr.Addr1 != null && !string.IsNullOrEmpty(addr.Addr1.GetValue()))
                    lines.Add(addr.Addr1.GetValue());
                if (addr.Addr2 != null && !string.IsNullOrEmpty(addr.Addr2.GetValue()))
                    lines.Add(addr.Addr2.GetValue());
                if (addr.Addr3 != null && !string.IsNullOrEmpty(addr.Addr3.GetValue()))
                    lines.Add(addr.Addr3.GetValue());
                if (addr.Addr4 != null && !string.IsNullOrEmpty(addr.Addr4.GetValue()))
                    lines.Add(addr.Addr4.GetValue());
                if (addr.Addr5 != null && !string.IsNullOrEmpty(addr.Addr5.GetValue()))
                    lines.Add(addr.Addr5.GetValue());

                // Add country if available
                if (addr.Country != null && !string.IsNullOrEmpty(addr.Country.GetValue()))
                    lines.Add(addr.Country.GetValue());

                return string.Join(", ", lines);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FormatAddress error: {ex.Message}");
                return "";
            }
        }

        public void Dispose()
        {
            if (_sessionManager != null && _isConnected)
            {
                try
                {
                    _sessionManager.EndSession();
                    _sessionManager.CloseConnection();
                }
                catch { }
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
}