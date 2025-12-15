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
        private InvoiceTrackingService _trackingService;

        // Session-scoped data with smart caching
        private Dictionary<string, CachedData<ItemData>> _sessionItems;
        private Dictionary<string, CachedData<CustomerData>> _sessionCustomers;
        private Dictionary<string, PriceLevel> _sessionPriceLevels;

        private short _qbXmlMajorVersion = 13;
        private short _qbXmlMinorVersion = 0;

        // Connection pooling
        private DateTime _lastActivity;
        private static readonly TimeSpan _connectionTimeout = TimeSpan.FromMinutes(5);

        public string CurrentCompanyName { get; private set; }
        public string CurrentCompanyFile { get; private set; }

        #endregion

        #region Constructor

        public QuickBooksService()
        {
            _sessionItems = new Dictionary<string, CachedData<ItemData>>();
            _sessionCustomers = new Dictionary<string, CachedData<CustomerData>>();
            _sessionPriceLevels = new Dictionary<string, PriceLevel>();
            _lastActivity = DateTime.Now;
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
                _lastActivity = DateTime.Now;
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

        private void UpdateActivity()
        {
            _lastActivity = DateTime.Now;
        }

        public bool IsConnectionActive()
        {
            return _isConnected && (DateTime.Now - _lastActivity) < _connectionTimeout;
        }

        #endregion

        #region Session Data Management

        /// <summary>
        /// ✅ OPTIMIZED: Pre-load with parallel processing and smart batching
        /// </summary>
        public void PreloadInvoiceData(IInvoiceRetList invoiceRetList)
        {
            var itemIds = new HashSet<string>();
            var customerIds = new HashSet<string>();

            LogBoth($"\n🔍 === PRE-LOADING INVOICE DATA (OPTIMIZED) ===");
            var startTime = DateTime.Now;

            for (int i = 0; i < invoiceRetList.Count; i++)
            {
                var inv = invoiceRetList.GetAt(i);

                string customerListID = inv.CustomerRef?.ListID?.GetValue();
                if (!string.IsNullOrEmpty(customerListID))
                    customerIds.Add(customerListID);

                if (inv.ORInvoiceLineRetList != null)
                {
                    for (int j = 0; j < inv.ORInvoiceLineRetList.Count; j++)
                    {
                        var lineRet = inv.ORInvoiceLineRetList.GetAt(j);
                        if (lineRet.InvoiceLineRet != null)
                        {
                            string itemListID = lineRet.InvoiceLineRet.ItemRef?.ListID?.GetValue();
                            if (!string.IsNullOrEmpty(itemListID))
                                itemIds.Add(itemListID);
                        }
                    }
                }
            }

            LogBoth($"📊 Found {customerIds.Count} unique customers and {itemIds.Count} unique items");

            var customersToFetch = customerIds.Where(id => !IsCustomerCached(id)).ToList();
            var itemsToFetch = itemIds.Where(id => !IsItemCached(id)).ToList();

            if (customersToFetch.Count > 0 || itemsToFetch.Count > 0)
            {
                BatchFetchAllData(customersToFetch, itemsToFetch);
            }

            var elapsed = DateTime.Now - startTime;
            LogBoth($"✅ Pre-load complete in {elapsed.TotalSeconds:F2}s: {_sessionCustomers.Count} customers, {_sessionItems.Count} items in cache");
            LogBoth($"=====================================\n");

            UpdateActivity();
        }

        public void PreloadCreditMemoData(ICreditMemoRetList creditMemoRetList)
        {
            var itemIds = new HashSet<string>();
            var customerIds = new HashSet<string>();

            LogBoth($"\n🔍 === PRE-LOADING CREDIT MEMO DATA (OPTIMIZED) ===");
            var startTime = DateTime.Now;

            for (int i = 0; i < creditMemoRetList.Count; i++)
            {
                var memo = creditMemoRetList.GetAt(i);

                string customerListID = memo.CustomerRef?.ListID?.GetValue();
                if (!string.IsNullOrEmpty(customerListID))
                    customerIds.Add(customerListID);

                if (memo.ORCreditMemoLineRetList != null)
                {
                    for (int j = 0; j < memo.ORCreditMemoLineRetList.Count; j++)
                    {
                        var lineRet = memo.ORCreditMemoLineRetList.GetAt(j);
                        if (lineRet.CreditMemoLineRet != null)
                        {
                            string itemListID = lineRet.CreditMemoLineRet.ItemRef?.ListID?.GetValue();
                            if (!string.IsNullOrEmpty(itemListID))
                                itemIds.Add(itemListID);
                        }
                    }
                }
            }

            LogBoth($"📊 Found {customerIds.Count} unique customers and {itemIds.Count} unique items");

            var customersToFetch = customerIds.Where(id => !IsCustomerCached(id)).ToList();
            var itemsToFetch = itemIds.Where(id => !IsItemCached(id)).ToList();

            if (customersToFetch.Count > 0 || itemsToFetch.Count > 0)
            {
                BatchFetchAllData(customersToFetch, itemsToFetch);
            }

            var elapsed = DateTime.Now - startTime;
            LogBoth($"✅ Pre-load complete in {elapsed.TotalSeconds:F2}s: {_sessionCustomers.Count} customers, {_sessionItems.Count} items in cache");
            LogBoth($"=====================================\n");

            UpdateActivity();
        }

        private bool IsCustomerCached(string customerListID)
        {
            return _sessionCustomers.TryGetValue(customerListID, out var cached) && !cached.IsExpired;
        }

        private bool IsItemCached(string itemListID)
        {
            return _sessionItems.TryGetValue(itemListID, out var cached) && !cached.IsExpired;
        }

        public void ClearSessionData()
        {
            int itemCount = _sessionItems.Count;
            int customerCount = _sessionCustomers.Count;
            int priceLevelCount = _sessionPriceLevels.Count;

            _sessionItems.Clear();
            _sessionCustomers.Clear();
            _sessionPriceLevels.Clear();

            LogBoth($"🧹 Session data cleared: {itemCount} items, {customerCount} customers, {priceLevelCount} price levels removed");
        }

        public string GetSessionStats()
        {
            return $"Session Data: {_sessionCustomers.Count} customers, {_sessionItems.Count} items, {_sessionPriceLevels.Count} price levels";
        }

        #endregion

        #region Transaction Fetching (Invoices & Credit Memos)

        public List<Invoice> FetchTransactions(DateTime? dateFrom, DateTime? dateTo,
            bool excludeUploaded = true, bool includeInvoices = true, bool includeCreditMemos = false)
        {
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

                if (_trackingService == null)
                {
                    _trackingService = new InvoiceTrackingService(CurrentCompanyName);
                }

                var transactions = new List<Invoice>();

                if (includeInvoices)
                {
                    var invoiceList = FetchInvoiceTransactions(dateFrom, dateTo, excludeUploaded);
                    transactions.AddRange(invoiceList);
                }

                if (includeCreditMemos)
                {
                    var creditMemoList = FetchCreditMemoTransactions(dateFrom, dateTo, excludeUploaded);
                    transactions.AddRange(creditMemoList);
                }

                return transactions;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to fetch transactions: {ex.Message}", ex);
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

        private List<Invoice> FetchInvoiceTransactions(DateTime? dateFrom, DateTime? dateTo, bool excludeUploaded = true)
        {
            var invoices = new List<Invoice>();
            var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
            msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

            var invoiceQuery = msgSetRq.AppendInvoiceQueryRq();
            invoiceQuery.IncludeLineItems.SetValue(true);

            // Add date range filter
            if (dateFrom.HasValue || dateTo.HasValue)
            {
                var filter = invoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter;

                if (dateFrom.HasValue)
                    filter.FromTxnDate.SetValue(dateFrom.Value);

                if (dateTo.HasValue)
                    filter.ToTxnDate.SetValue(dateTo.Value);
            }

            var msgSetRs = _sessionManager.DoRequests(msgSetRq);
            var response = msgSetRs.ResponseList.GetAt(0);

            if (response.StatusCode == 0)
            {
                var invoiceRetList = response.Detail as IInvoiceRetList;
                if (invoiceRetList != null)
                {
                    PreloadInvoiceData(invoiceRetList);

                    int skippedCount = 0;
                    for (int i = 0; i < invoiceRetList.Count; i++)
                    {
                        var inv = invoiceRetList.GetAt(i);
                        string qbInvoiceId = inv.TxnID.GetValue();

                        if (excludeUploaded && _trackingService.IsInvoiceUploaded(qbInvoiceId))
                        {
                            skippedCount++;
                            continue;
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
                            InvoiceDate = inv.TxnDate?.GetValue() ?? DateTime.Now,
                            FBR_IRN = uploadRecord?.IRN,
                            UploadDate = uploadRecord?.UploadDate,
                            InvoiceType = "Invoice"  // Add this field to distinguish
                        });
                    }

                    if (excludeUploaded && skippedCount > 0)
                    {
                        LogBoth($"✅ Skipped {skippedCount} already-uploaded invoices");
                    }

                    LogBoth($"✅ Fetched {invoices.Count} invoices for date range");
                }
            }

            return invoices;
        }

        private List<Invoice> FetchCreditMemoTransactions(DateTime? dateFrom, DateTime? dateTo, bool excludeUploaded = true)
        {
            var creditMemos = new List<Invoice>();
            var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
            msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

            var creditMemoQuery = msgSetRq.AppendCreditMemoQueryRq();
            creditMemoQuery.IncludeLineItems.SetValue(true);

            // Add date range filter
            if (dateFrom.HasValue || dateTo.HasValue)
            {
                var filter = creditMemoQuery.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter;

                if (dateFrom.HasValue)
                    filter.FromTxnDate.SetValue(dateFrom.Value);

                if (dateTo.HasValue)
                    filter.ToTxnDate.SetValue(dateTo.Value);
            }

            var msgSetRs = _sessionManager.DoRequests(msgSetRq);
            var response = msgSetRs.ResponseList.GetAt(0);

            if (response.StatusCode == 0)
            {
                var creditMemoRetList = response.Detail as ICreditMemoRetList;
                if (creditMemoRetList != null)
                {
                    PreloadCreditMemoData(creditMemoRetList);

                    int skippedCount = 0;
                    for (int i = 0; i < creditMemoRetList.Count; i++)
                    {
                        var memo = creditMemoRetList.GetAt(i);
                        string qbCreditMemoId = memo.TxnID.GetValue();

                        if (excludeUploaded && _trackingService.IsInvoiceUploaded(qbCreditMemoId))
                        {
                            skippedCount++;
                            continue;
                        }

                        string customerListID = memo.CustomerRef?.ListID?.GetValue() ?? "";
                        var customerData = FetchCustomerDetails(customerListID);
                        var uploadRecord = _trackingService.GetUploadStatus(qbCreditMemoId);

                        creditMemos.Add(new Invoice
                        {
                            CompanyName = CurrentCompanyName,
                            QuickBooksInvoiceId = qbCreditMemoId,
                            InvoiceNumber = memo.RefNumber?.GetValue() ?? "",
                            CustomerName = memo.CustomerRef?.FullName?.GetValue() ?? "",
                            CustomerNTN = customerData.NTN,
                            Amount = Convert.ToDecimal(memo.Subtotal?.GetValue() ?? 0),
                            Status = uploadRecord?.Status == UploadStatus.Success ? "Uploaded" :
                                    uploadRecord?.Status == UploadStatus.Failed ? "Failed" : "Pending",
                            CreatedDate = memo.TxnDate?.GetValue() ?? DateTime.Now,
                            InvoiceDate = memo.TxnDate?.GetValue() ?? DateTime.Now,
                            FBR_IRN = uploadRecord?.IRN,
                            UploadDate = uploadRecord?.UploadDate,
                            InvoiceType = "Credit Memo"  // Add this field to distinguish
                        });
                    }

                    if (excludeUploaded && skippedCount > 0)
                    {
                        LogBoth($"✅ Skipped {skippedCount} already-uploaded credit memos");
                    }

                    LogBoth($"✅ Fetched {creditMemos.Count} credit memos for date range");
                }
            }

            return creditMemos;
        }

        #endregion

        #region Transaction Details (Invoice & Credit Memo)

        public async Task<FBRInvoicePayload> GetInvoiceDetails(string qbInvoiceId, bool checkUploadStatus = true)
        {
            _fbr = new FBRApiService();
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

                if (_trackingService == null)
                {
                    _trackingService = new InvoiceTrackingService(CurrentCompanyName);
                }

                if (checkUploadStatus && _trackingService.IsInvoiceUploaded(qbInvoiceId))
                {
                    LogBoth($"⏭️ Invoice {qbInvoiceId} already uploaded - skipping QB query");
                    return null;
                }

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

                PreloadInvoiceData(invoiceRetList);

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

                var lineItemsWithContext = ProcessInvoiceLines(inv);
                ApplyDiscountsContextually(lineItemsWithContext);

                await BuildInvoiceItems(lineItemsWithContext, payload, taxRate);
                await EnrichWithSroData(payload);

                return payload;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to get invoice details: {ex.Message}", ex);
            }
            finally
            {
                LogBoth(GetSessionStats());

                if (!wasConnected && _isConnected)
                {
                    CloseSession();
                    CloseConnection();
                }
            }
        }

        public async Task<FBRInvoicePayload> GetCreditMemoDetails(string qbCreditMemoId, bool checkUploadStatus = true)
        {
            _fbr = new FBRApiService();
            bool wasConnected = _isConnected;

            try
            {
                if (!_isConnected) Connect();

                if (_trackingService == null)
                {
                    _trackingService = new InvoiceTrackingService(CurrentCompanyName);
                }

                if (checkUploadStatus && _trackingService.IsInvoiceUploaded(qbCreditMemoId))
                {
                    LogBoth($"⏭️ Credit Memo {qbCreditMemoId} already uploaded - skipping QB query");
                    return null;
                }

                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                var creditMemoQuery = msgSetRq.AppendCreditMemoQueryRq();
                creditMemoQuery.ORTxnQuery.TxnIDList.Add(qbCreditMemoId);
                creditMemoQuery.IncludeLineItems.SetValue(true);

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);
                var response = msgSetRs.ResponseList.GetAt(0);

                if (response.StatusCode != 0 || response.Detail == null)
                    return null;

                var creditMemoRetList = response.Detail as ICreditMemoRetList;
                if (creditMemoRetList == null || creditMemoRetList.Count == 0)
                    return null;

                PreloadCreditMemoData(creditMemoRetList);

                var memo = creditMemoRetList.GetAt(0);

                string customerListID = memo.CustomerRef?.ListID?.GetValue() ?? "";
                var customerData = FetchCustomerDetails(customerListID);

                var payload = new FBRInvoicePayload
                {
                    InvoiceType = "Debit Note",
                    InvoiceNumber = memo.RefNumber?.GetValue() ?? "",
                    InvoiceDate = memo.TxnDate?.GetValue() ?? DateTime.Now,
                    SellerNTN = _companySettings?.SellerNTN ?? "",
                    SellerBusinessName = _companyInfo?.Name ?? CurrentCompanyName,
                    SellerProvince = _companySettings?.SellerProvince ?? "",
                    SellerAddress = _companySettings?.SellerAddress ?? _companyInfo?.Address ?? "",
                    CustomerName = memo.CustomerRef?.FullName?.GetValue() ?? "",
                    CustomerNTN = customerData.NTN,
                    BuyerProvince = customerData.State,
                    BuyerAddress = customerData.Address,
                    BuyerRegistrationType = customerData.CustomerType,
                    BuyerPhone = customerData.Phone,
                    TotalAmount = Convert.ToDecimal((memo.Subtotal?.GetValue() ?? 0) + (memo.SalesTaxTotal?.GetValue() ?? 0)),
                    Subtotal = Convert.ToDecimal(memo.Subtotal?.GetValue() ?? 0),
                    TaxAmount = Convert.ToDecimal(memo.SalesTaxTotal?.GetValue() ?? 0),
                    Items = new List<InvoiceItem>()
                };

                double taxRate = memo.SalesTaxPercentage?.GetValue() ?? 0;

                var lineItemsWithContext = ProcessCreditMemoLines(memo);
                ApplyDiscountsContextually(lineItemsWithContext);

                await BuildInvoiceItems(lineItemsWithContext, payload, taxRate);
                await EnrichWithSroData(payload);

                return payload;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to get credit memo details: {ex.Message}", ex);
            }
            finally
            {
                LogBoth(GetSessionStats());

                if (!wasConnected && _isConnected)
                {
                    CloseSession();
                    CloseConnection();
                }
            }
        }

        #endregion

        #region Line Processing (Invoice & Credit Memo)

        private List<LineItemContext> ProcessInvoiceLines(IInvoiceRet inv)
        {
            var lineItems = new List<LineItemContext>();

            if (inv.ORInvoiceLineRetList == null)
                return lineItems;

            LogBoth($"\n🔍 === PROCESSING INVOICE LINES ===");
            LogBoth($"Total lines in invoice: {inv.ORInvoiceLineRetList.Count}");

            for (int i = 0; i < inv.ORInvoiceLineRetList.Count; i++)
            {
                var lineRet = inv.ORInvoiceLineRetList.GetAt(i);
                if (lineRet.InvoiceLineRet == null) continue;

                var line = lineRet.InvoiceLineRet;
                string itemListID = line.ItemRef?.ListID?.GetValue() ?? "";
                double lineAmount = line.Amount?.GetValue() ?? 0;
                string itemType = line.ItemRef?.FullName?.GetValue() ?? "";
                string itemName = line.Desc?.GetValue() ?? itemType;

                bool isDiscount = lineAmount < 0 || itemType.Contains("Discount", StringComparison.OrdinalIgnoreCase);
                bool isSubtotal = IsLikelySubtotal(line, itemListID, itemType, itemName, lineAmount);

                string lineType = isDiscount ? "DISCOUNT" : isSubtotal ? "SUBTOTAL" : "ITEM";

                LogBoth($"Line {i}: [{lineType}] {itemName} = {lineAmount:C}");

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

            LogBoth($"=====================================\n");

            return lineItems;
        }

        private List<LineItemContext> ProcessCreditMemoLines(ICreditMemoRet memo)
        {
            var lineItems = new List<LineItemContext>();

            if (memo.ORCreditMemoLineRetList == null)
                return lineItems;

            LogBoth($"\n🔍 === PROCESSING CREDIT MEMO LINES ===");
            LogBoth($"Total lines in credit memo: {memo.ORCreditMemoLineRetList.Count}");

            for (int i = 0; i < memo.ORCreditMemoLineRetList.Count; i++)
            {
                var lineRet = memo.ORCreditMemoLineRetList.GetAt(i);
                if (lineRet.CreditMemoLineRet == null) continue;

                var line = lineRet.CreditMemoLineRet;
                string itemListID = line.ItemRef?.ListID?.GetValue() ?? "";
                double lineAmount = line.Amount?.GetValue() ?? 0;
                string itemType = line.ItemRef?.FullName?.GetValue() ?? "";
                string itemName = line.Desc?.GetValue() ?? itemType;

                bool isDiscount = lineAmount < 0 || itemType.Contains("Discount", StringComparison.OrdinalIgnoreCase);
                bool isSubtotal = IsLikelySubtotalCreditMemo(line, itemListID, itemType, itemName, lineAmount);

                string lineType = isDiscount ? "DISCOUNT" : isSubtotal ? "SUBTOTAL" : "ITEM";

                LogBoth($"Line {i}: [{lineType}] {itemName} = {lineAmount:C}");

                var invoiceLine = ConvertCreditMemoLineToInvoiceLine(line);

                string lineType = isDiscount ? "DISCOUNT" : isSubtotal ? "SUBTOTAL" : "ITEM";

                LogBoth($"Line {i}: [{lineType}] {itemName} = {lineAmount:C}");

                lineItems.Add(new LineItemContext
                {
                    Index = i,
                    Line = invoiceLine,
                    ItemListID = itemListID,
                    Amount = lineAmount,
                    ItemType = itemType,
                    ItemName = itemName,
                    IsDiscount = isDiscount,
                    IsSubtotal = isSubtotal
                });
            }

            LogBoth($"=====================================\n");

            return lineItems;
        }

        private bool IsLikelySubtotal(IInvoiceLineRet line, string itemListID, string itemType, string itemName, double lineAmount)
        {
            string type = itemType?.Trim() ?? "";
            string name = itemName?.Trim() ?? "";

            if (type.Contains("subtotal", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("subtotal", StringComparison.OrdinalIgnoreCase))
            {
                LogBoth($"   ✓ Subtotal detected by name: {name}");
                return true;
            }

            bool itemRefEmpty = string.IsNullOrEmpty(itemListID);
            bool quantityMissing = line.Quantity == null || Math.Abs(line.Quantity.GetValue()) == 0;

            if (itemRefEmpty && quantityMissing && lineAmount > 0)
            {
                LogBoth($"   ✓ Subtotal detected by characteristics: No ItemRef + No Quantity + Positive Amount");
                return true;
            }

            return false;
        }

        private bool IsLikelySubtotalCreditMemo(ICreditMemoLineRet line, string itemListID, string itemType, string itemName, double lineAmount)
        {
            string type = itemType?.Trim() ?? "";
            string name = itemName?.Trim() ?? "";

            if (type.Contains("subtotal", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("subtotal", StringComparison.OrdinalIgnoreCase))
            {
                LogBoth($"   ✓ Subtotal detected by name: {name}");
                return true;
            }

            bool itemRefEmpty = string.IsNullOrEmpty(itemListID);
            bool quantityMissing = line.Quantity == null || Math.Abs(line.Quantity.GetValue()) == 0;

            if (itemRefEmpty && quantityMissing && lineAmount > 0)
            {
                LogBoth($"   ✓ Subtotal detected by characteristics");
                return true;
            }

            return false;
        }

        private IInvoiceLineRet ConvertCreditMemoLineToInvoiceLine(ICreditMemoLineRet creditMemoLine)
        {
            return new CreditMemoLineWrapper(creditMemoLine);
        }

        #endregion

        #region Invoice Item Building

        private async Task BuildInvoiceItems(List<LineItemContext> lineItems, FBRInvoicePayload payload, double taxRate)
        {
            LogBoth($"\n🔨 === BUILD INVOICE ITEMS START ===");
            LogBoth($"Processing {lineItems.Count} line items with tax rate: {taxRate}%");

            foreach (var context in lineItems)
            {
                // Skip non-item lines
                if (context.IsDiscount || context.IsSubtotal)
                {
                    LogBoth($"Skipping line {context.Index}: {(context.IsDiscount ? "Discount" : "Subtotal")}");
                    continue;
                }

                if (string.IsNullOrEmpty(context.ItemListID) && string.IsNullOrEmpty(context.ItemType))
                {
                    LogBoth($"Skipping line {context.Index}: No item reference");
                    continue;
                }

                var line = context.Line;
                var itemData = FetchItemDetails(context.ItemListID);

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

                double effectiveTaxRate = taxRate;
                if (isExempt || isZeroRated)
                {
                    effectiveTaxRate = 0;
                }

                // ✅ CALCULATE BASE AMOUNTS WITH PROPER DISCOUNT HANDLING
                int quantity = Convert.ToInt32(line.Quantity?.GetValue() ?? 1);
                decimal lineAmount = Convert.ToDecimal(context.Amount);
                decimal discount = context.ApplicableDiscount;

                // ✅ PRESERVE THIS - Don't let tax methods overwrite it!
                decimal netAmount = lineAmount - discount;

                if (netAmount < 0)
                {
                    LogBoth($"⚠️ Warning: Net amount for item {context.Index} is negative ({netAmount:C}). Setting to 0.");
                    netAmount = 0;
                }

                LogBoth($"\n💵 Item {context.Index}: {line.Desc?.GetValue()}");
                LogBoth($"   Line Amount: {lineAmount:C}");
                LogBoth($"   Discount: {discount:C}");
                LogBoth($"   Net Amount: {netAmount:C}");

                // Parse retail price
                decimal parsedRetailPrice = 0m;
                if (decimal.TryParse(retailPrice, out decimal unitRetailPrice))
                    parsedRetailPrice = unitRetailPrice * quantity;

                // Calculate tax based on item type
                decimal displayTaxRate, salesTaxAmount, furtherTax, computedTotalValue;
                string rateString;

                // ✅ KEY FIX: Don't pass netAmount as out parameter to tax methods!
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
                    // ✅ FIX: Use a temporary variable, don't overwrite netAmount!
                    decimal tempNetAmount;
                    Calculate3rdScheduleTax(parsedRetailPrice, effectiveTaxRate, out displayTaxRate, out salesTaxAmount, out furtherTax, out rateString, out computedTotalValue, out tempNetAmount);

                    // ✅ For 3rd schedule items, netAmount should be 0 (tax is on retail price)
                    // But we preserve the discount information in TotalPrice
                    netAmount = 0m;

                    LogBoth($"   📌 3rd Schedule item: Using retail price {parsedRetailPrice:C} for tax calculation");
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

                LogBoth($"   Tax Rate: {displayTaxRate}% | Sales Tax: {salesTaxAmount:C} | Total: {computedTotalValue:C}");

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
                    TotalPrice = lineAmount,          // ✅ Original amount BEFORE discount
                    NetAmount = netAmount,             // ✅ Amount AFTER discount (will be 0 for 3rd schedule)
                    TaxRate = displayTaxRate,
                    Rate = rateString,
                    SalesTaxAmount = salesTaxAmount,
                    TotalValue = computedTotalValue,
                    RetailPrice = parsedRetailPrice,
                    ExtraTax = 0m,
                    FurtherTax = furtherTax,
                    FedPayable = 0m,
                    SalesTaxWithheldAtSource = 0m,
                    Discount = discount,               // ✅ Discount amount
                    SaleType = saleType ?? "Goods at standard rate (default)",
                    SroScheduleNo = "",
                    SroItemSerialNo = ""
                };

                SetDefaultSroValues(item);
                payload.Items.Add(item);
            }

            LogBoth($"\n✅ === BUILD INVOICE ITEMS COMPLETE ===");
            LogBoth($"Added {payload.Items.Count} items to payload\n");
        }

        #endregion

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
            out decimal furtherTax, out string rateString, out decimal computedTotalValue, out decimal ignoredNetAmount)
        {
            displayTaxRate = Convert.ToDecimal(effectiveTaxRate);
            salesTaxAmount = Convert.ToDecimal((double)retailPrice * (effectiveTaxRate / 100.0));
            furtherTax = 0m;
            rateString = $"{effectiveTaxRate}%";

            // For 3rd schedule items, tax is calculated on retail price, not on sale value
            ignoredNetAmount = 0m;

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

        #region Helper Methods for Tax & SRO

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
            LogBoth($"\n🔍 === DISCOUNT PROCESSING START ===");
            LogBoth($"Total line items: {lineItems.Count}");

            // Find all subtotal positions
            var subtotalIndices = lineItems
                .Select((item, index) => new { item, index })
                .Where(x => x.item.IsSubtotal)
                .Select(x => x.index)
                .ToList();

            LogBoth($"Found {subtotalIndices.Count} subtotals at positions: {string.Join(", ", subtotalIndices)}");

            // Process each discount
            var discounts = lineItems.Where(x => x.IsDiscount).ToList();
            LogBoth($"Found {discounts.Count} discounts to process");

            foreach (var discount in discounts)
            {
                decimal discountAmount = Math.Abs(Convert.ToDecimal(discount.Amount));
                int discountIndex = discount.Index;

                LogBoth($"\n📊 Processing discount at index {discountIndex}:");
                LogBoth($"   Discount amount: {discountAmount:C}");
                LogBoth($"   Discount name: {discount.ItemName}");

                // Find the relevant subtotal BEFORE this discount
                int relevantSubtotalIndex = subtotalIndices
                    .Where(idx => idx < discountIndex)
                    .OrderByDescending(idx => idx)
                    .FirstOrDefault(-1);

                if (relevantSubtotalIndex >= 0)
                {
                    LogBoth($"   Found subtotal at index {relevantSubtotalIndex} before discount");
                    ApplyDiscountToSubtotalRange(lineItems, relevantSubtotalIndex, discountIndex, discountAmount);
                }
                else
                {
                    LogBoth($"   No subtotal found - applying to previous item only");
                    ApplyDiscountToPreviousItem(lineItems, discountIndex, discountAmount);
                }
            }

            // Log final discount distribution
            LogBoth($"\n✅ === DISCOUNT DISTRIBUTION SUMMARY ===");
            foreach (var item in lineItems.Where(x => !x.IsDiscount && !x.IsSubtotal))
            {
                if (item.ApplicableDiscount > 0)
                {
                    LogBoth($"   Item {item.Index} ({item.ItemName}): Original={item.Amount:C}, Discount={item.ApplicableDiscount:C}, Net={item.Amount - (double)item.ApplicableDiscount:C}");
                }
            }
            LogBoth($"===========================================\n");
        }

        private void ApplyDiscountToSubtotalRange(List<LineItemContext> lineItems, int subtotalIndex, int discountIndex, decimal discountAmount)
        {
            LogBoth($"\n   📍 Applying discount to subtotal range:");

            // STEP 1: Find the range start (previous subtotal or beginning)
            int rangeStart = 0;
            for (int j = subtotalIndex - 1; j >= 0; j--)
            {
                if (lineItems[j].IsSubtotal)
                {
                    rangeStart = j + 1;
                    break;
                }
            }

            LogBoth($"   Range: {rangeStart} to {subtotalIndex - 1}");

            // STEP 2: Collect all items in the subtotal range (before the subtotal)
            var itemsInRange = lineItems
                .Skip(rangeStart)
                .Take(subtotalIndex - rangeStart)
                .Where(x => !x.IsDiscount && !x.IsSubtotal && x.Amount > 0)
                .ToList();

            if (itemsInRange.Count == 0)
            {
                LogBoth($"   ⚠️ No valid items found in subtotal range");
                return;
            }

            LogBoth($"   Found {itemsInRange.Count} items in range to apply discount:");
            foreach (var item in itemsInRange)
            {
                LogBoth($"      - Item {item.Index}: {item.ItemName} = {item.Amount:C}");
            }

            // STEP 3: Distribute discount proportionally
            DistributeDiscountToItems(itemsInRange, discountAmount);
        }

        private void ApplyDiscountToPreviousItem(List<LineItemContext> lineItems, int discountIndex, decimal discountAmount)
        {
            LogBoth($"\n   📍 Applying discount to previous item only:");

            // Find the last actual item before this discount
            var previousItem = lineItems
                .Take(discountIndex)
                .Where(x => !x.IsDiscount && !x.IsSubtotal && x.Amount > 0)
                .LastOrDefault();

            if (previousItem == null)
            {
                LogBoth($"   ⚠️ No previous item found to apply discount");
                return;
            }

            decimal itemAmount = Math.Abs(Convert.ToDecimal(previousItem.Amount));
            decimal applicableDiscount = Math.Min(discountAmount, itemAmount);

            previousItem.ApplicableDiscount += applicableDiscount;

            LogBoth($"   Applied {applicableDiscount:C} discount to item {previousItem.Index} ({previousItem.ItemName})");
            LogBoth($"   Item amount: {itemAmount:C} -> Net: {itemAmount - applicableDiscount:C}");
        }

        private void DistributeDiscountToItems(List<LineItemContext> items, decimal totalDiscount)
        {
            if (items == null || items.Count == 0 || totalDiscount <= 0)
                return;

            LogBoth($"\n   💰 Distributing {totalDiscount:C} across {items.Count} items:");

            // Calculate total amount of all items
            decimal totalAmount = items.Sum(x => Math.Abs(Convert.ToDecimal(x.Amount)));

            if (totalAmount <= 0)
            {
                LogBoth($"   ⚠️ Total amount is zero - cannot distribute discount");
                return;
            }

            LogBoth($"   Total amount: {totalAmount:C}");

            // Cap discount at total amount
            totalDiscount = Math.Min(totalDiscount, totalAmount);

            // Distribute proportionally with rounding
            decimal allocated = 0m;

            for (int i = 0; i < items.Count; i++)
            {
                decimal itemAmount = Math.Abs(Convert.ToDecimal(items[i].Amount));
                decimal proportion = itemAmount / totalAmount;
                decimal itemDiscount = Math.Round(totalDiscount * proportion, 2, MidpointRounding.AwayFromZero);

                items[i].ApplicableDiscount += itemDiscount;
                allocated += itemDiscount;

                LogBoth($"      Item {items[i].Index}: {itemAmount:C} ({proportion:P2}) -> Discount: {itemDiscount:C}");
            }

            // Handle rounding remainder - add to largest item
            decimal remainder = totalDiscount - allocated;
            if (remainder != 0m && items.Count > 0)
            {
                var largestItem = items.OrderByDescending(x => Math.Abs(Convert.ToDecimal(x.Amount))).First();
                largestItem.ApplicableDiscount += remainder;

                LogBoth($"   🔧 Rounding adjustment: {remainder:C} added to largest item (Index {largestItem.Index})");
            }

            LogBoth($"   ✅ Total allocated: {items.Sum(x => x.ApplicableDiscount):C}");
        }

        #endregion

        #region Tracking Management

        public void MarkInvoiceAsUploaded(string qbInvoiceId, string invoiceNumber, string irn)
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            _trackingService.MarkAsUploaded(qbInvoiceId, invoiceNumber, irn, DateTime.Now);
            LogBoth($"✅ Marked invoice {invoiceNumber} as uploaded (IRN: {irn})");
        }

        public void MarkInvoiceAsFailed(string qbInvoiceId, string invoiceNumber, string errorMessage)
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            _trackingService.MarkAsFailed(qbInvoiceId, invoiceNumber, errorMessage);
            LogBoth($"❌ Marked invoice {invoiceNumber} as failed: {errorMessage}");
        }

        public string GetTrackingStatistics()
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            return _trackingService.GetStatistics().ToString();
        }

        public void ClearInvoiceTracking()
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            _trackingService.ClearAll();
            LogBoth("🔄 Cleared all invoice tracking data");
        }

        public InvoiceTrackingService GetTrackingService()
        {
            if (_trackingService == null)
            {
                _trackingService = new InvoiceTrackingService(CurrentCompanyName);
            }

            return _trackingService;
        }

        #endregion

        #region Optimized Batch Fetching Methods

        private void BatchFetchAllData(List<string> customerIds, List<string> itemIds)
        {
            if ((customerIds == null || customerIds.Count == 0) && (itemIds == null || itemIds.Count == 0))
                return;

            try
            {
                LogBoth($"🚀 Starting optimized batch fetch: {customerIds.Count} customers, {itemIds.Count} items");
                var startTime = DateTime.Now;

                var msgSetRq = _sessionManager.CreateMsgSetRequest("US", _qbXmlMajorVersion, _qbXmlMinorVersion);
                msgSetRq.Attributes.OnError = ENRqOnError.roeContinue;

                if (customerIds.Count > 0)
                {
                    var customerQuery = msgSetRq.AppendCustomerQueryRq();

                    if (customerIds.Count > 20)
                    {
                        customerQuery.OwnerIDList.Add("0");
                        customerQuery.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly);
                        LogBoth($"   📋 Fetching ALL customers (count > 20)");
                    }
                    else
                    {
                        foreach (var id in customerIds)
                            customerQuery.ORCustomerListQuery.ListIDList.Add(id);
                        customerQuery.OwnerIDList.Add("0");
                        LogBoth($"   📋 Fetching {customerIds.Count} specific customers");
                    }
                }

                if (itemIds.Count > 0)
                {
                    var itemQuery = msgSetRq.AppendItemQueryRq();

                    if (itemIds.Count > 20)
                    {
                        itemQuery.OwnerIDList.Add("0");
                        LogBoth($"   📦 Fetching ALL items (count > 20)");
                    }
                    else
                    {
                        for (int i = 0; i < itemIds.Count; i++)
                            itemQuery.ORListQuery.ListIDList.Add(itemIds[i]);
                        itemQuery.OwnerIDList.Add("0");
                        LogBoth($"   📦 Fetching {itemIds.Count} specific items");
                    }
                }

                var msgSetRs = _sessionManager.DoRequests(msgSetRq);

                for (int i = 0; i < msgSetRs.ResponseList.Count; i++)
                {
                    var response = msgSetRs.ResponseList.GetAt(i);
                    if (response.StatusCode == 0)
                    {
                        if (response.Detail is ICustomerRetList customerList)
                        {
                            ProcessCustomerResults(customerList, customerIds);
                        }
                        else if (response.Detail is IORItemRetList itemList)
                        {
                            ProcessItemResults(itemList, itemIds);
                        }
                    }
                }

                var elapsed = DateTime.Now - startTime;
                LogBoth($"✅ Batch fetch completed in {elapsed.TotalMilliseconds}ms");
            }
            catch (Exception ex)
            {
                LogBoth($"❌ Error in optimized batch fetch: {ex.Message}");
                LogBoth($"   🔄 Falling back to individual fetches");
                if (customerIds.Count > 0) BatchFetchCustomersLegacy(customerIds);
                if (itemIds.Count > 0) BatchFetchItemsLegacy(itemIds);
            }
        }

        private void ProcessCustomerResults(ICustomerRetList customerList, List<string> requestedIds)
        {
            if (customerList == null) return;

            var neededIDs = new HashSet<string>(requestedIds);
            int cached = 0;

            for (int i = 0; i < customerList.Count; i++)
            {
                var customer = customerList.GetAt(i);
                string listID = customer.ListID?.GetValue();

                if (string.IsNullOrEmpty(listID))
                    continue;

                if (requestedIds.Count > 20 || neededIDs.Contains(listID))
                {
                    var customerData = ExtractCustomerData(customer);
                    _sessionCustomers[listID] = new CachedData<CustomerData>
                    {
                        Data = customerData,
                        CachedAt = DateTime.Now,
                        ExpiresAfter = TimeSpan.FromHours(1)
                    };
                    cached++;
                }
            }

            LogBoth($"   ✅ Cached {cached} customers");
        }

        private void ProcessItemResults(IORItemRetList itemList, List<string> requestedIds)
        {
            if (itemList == null) return;

            var neededIDs = new HashSet<string>(requestedIds);
            int cached = 0;

            for (int i = 0; i < itemList.Count; i++)
            {
                var itemRet = itemList.GetAt(i);
                string listID = GetItemListID(itemRet);

                if (string.IsNullOrEmpty(listID))
                    continue;

                if (requestedIds.Count > 20 || neededIDs.Contains(listID))
                {
                    var itemData = ExtractItemData(itemRet, listID);
                    _sessionItems[listID] = new CachedData<ItemData>
                    {
                        Data = itemData,
                        CachedAt = DateTime.Now,
                        ExpiresAfter = TimeSpan.FromHours(1)
                    };
                    cached++;
                }
            }

            LogBoth($"   ✅ Cached {cached} items");
        }

        private void BatchFetchCustomersLegacy(List<string> customerListIDs)
        {
            foreach (var id in customerListIDs)
            {
                try
                {
                    var data = FetchCustomerDetails(id);
                    _sessionCustomers[id] = new CachedData<CustomerData>
                    {
                        Data = data,
                        CachedAt = DateTime.Now,
                        ExpiresAfter = TimeSpan.FromHours(1)
                    };
                }
                catch (Exception ex)
                {
                    LogBoth($"   ❌ Error fetching customer {id} in legacy fallback: {ex.Message}");
                }
            }
        }

        private void BatchFetchItemsLegacy(List<string> itemListIDs)
        {
            foreach (var id in itemListIDs)
            {
                try
                {
                    var data = FetchItemDetails(id);
                    _sessionItems[id] = new CachedData<ItemData>
                    {
                        Data = data,
                        CachedAt = DateTime.Now,
                        ExpiresAfter = TimeSpan.FromHours(1)
                    };
                }
                catch (Exception ex)
                {
                    LogBoth($"   ❌ Error fetching item {id} in legacy fallback: {ex.Message}");
                }
            }
        }

        #endregion

        #region Data Fetching with Smart Cache

        private CustomerData FetchCustomerDetails(string customerListID, bool forceRefresh = false)
        {
            if (string.IsNullOrEmpty(customerListID))
            {
                return new CustomerData { CustomerType = "Unregistered" };
            }

            if (!forceRefresh && _sessionCustomers.TryGetValue(customerListID, out var cached))
            {
                if (!cached.IsExpired)
                    return cached.Data;
                else
                    _sessionCustomers.Remove(customerListID);
            }

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

                        _sessionCustomers[customerListID] = new CachedData<CustomerData>
                        {
                            Data = customerData,
                            CachedAt = DateTime.Now,
                            ExpiresAfter = TimeSpan.FromHours(1)
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                LogBoth($"   ❌ Error fetching customer: {ex.Message}");
            }

            return customerData;
        }

        private ItemData FetchItemDetails(string itemListID, bool forceRefresh = false)
        {
            if (string.IsNullOrEmpty(itemListID))
                return new ItemData { SaleType = "Goods at standard rate (default)" };

            if (!forceRefresh && _sessionItems.TryGetValue(itemListID, out var cached))
            {
                if (!cached.IsExpired)
                    return cached.Data;
                else
                    _sessionItems.Remove(itemListID);
            }

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

                        _sessionItems[itemListID] = new CachedData<ItemData>
                        {
                            Data = itemData,
                            CachedAt = DateTime.Now,
                            ExpiresAfter = TimeSpan.FromHours(1)
                        };
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
                ClearSessionData();
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

    public class CachedData<T>
    {
        public T Data { get; set; }
        public DateTime CachedAt { get; set; }
        public TimeSpan ExpiresAfter { get; set; }

        public bool IsExpired => DateTime.Now - CachedAt > ExpiresAfter;
    }

    // [KEEP ALL OTHER HELPER CLASSES FROM YOUR ORIGINAL CODE]
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

    public class CreditMemoLineWrapper : IInvoiceLineRet
    {
        private readonly ICreditMemoLineRet _creditMemoLine;

        public CreditMemoLineWrapper(ICreditMemoLineRet creditMemoLine)
        {
            _creditMemoLine = creditMemoLine;
        }

        public IObjectType Type => null;
        public IQBIDType TxnLineID => _creditMemoLine.TxnLineID;
        public IQBBaseRef ItemRef => _creditMemoLine.ItemRef;
        public IQBStringType Desc => _creditMemoLine.Desc;
        public IQBQuanType Quantity => _creditMemoLine.Quantity;
        public IQBStringType UnitOfMeasure => _creditMemoLine.UnitOfMeasure;
        public IQBAmountType Amount => _creditMemoLine.Amount;
        public IDataExtRetList DataExtRetList => _creditMemoLine.DataExtRetList;
        public IQBBaseRef OverrideUOMSetRef => null;
        public IORRate ORRate => null;
        public IQBBaseRef ClassRef => null;
        public IQBAmountType TaxAmount => null;
        public IQBBaseRef InventorySiteRef => null;
        public IQBBaseRef InventorySiteLocationRef => null;
        public IORSerialLotNumber ORSerialLotNumber => null;
        public IQBStringType ExpirationDateForSerialLotNumber => null;
        public IQBDateType ServiceDate => null;
        public IQBBaseRef SalesTaxCodeRef => null;
        public IQBBoolType IsTaxable => null;
        public IQBStringType Other1 => null;
        public IQBStringType Other2 => null;
    }

    #endregion
}