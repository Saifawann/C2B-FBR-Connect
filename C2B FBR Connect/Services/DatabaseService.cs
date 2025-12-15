using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Threading;

namespace C2B_FBR_Connect.Services
{
    public class DatabaseService
    {
        private readonly string _connectionString;
        private readonly string _dbPath;
        private static readonly object _lockObject = new object();
        private const int MaxRetries = 3;
        private const int RetryDelayMs = 100;

        public string DatabasePath => _dbPath;

        public DatabaseService()
        {
            _dbPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "C2b Smart App",
                "c2b_smart_app.db"
            );

            var dir = Path.GetDirectoryName(_dbPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            _connectionString = $"Data Source={_dbPath};Version=3;Journal Mode=WAL;Busy Timeout=5000;Pooling=True;Max Pool Size=100;";

            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            lock (_lockObject)
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                ExecuteNonQuery(conn, "PRAGMA journal_mode=WAL;");
                ExecuteNonQuery(conn, "PRAGMA synchronous=NORMAL;");
                ExecuteNonQuery(conn, "PRAGMA cache_size=10000;");
                ExecuteNonQuery(conn, "PRAGMA temp_store=MEMORY;");

                string createCompaniesTable = @"
                    CREATE TABLE IF NOT EXISTS Companies (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        CompanyName TEXT UNIQUE NOT NULL,
                        FBRToken TEXT NOT NULL,
                        SellerNTN TEXT,
                        SellerAddress TEXT,
                        SellerProvince TEXT,
                        SellerPhone TEXT,
                        SellerEmail TEXT,
                        LogoImage BLOB,
                        Environment TEXT DEFAULT 'Sandbox',
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        ModifiedDate DATETIME
                    );";

                // ✅ Changed TransactionType to InvoiceType
                string createInvoicesTable = @"
                    CREATE TABLE IF NOT EXISTS Invoices (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        CompanyName TEXT NOT NULL,
                        QuickBooksInvoiceId TEXT NOT NULL,
                        InvoiceNumber TEXT NOT NULL,
                        CustomerName TEXT NOT NULL,
                        CustomerNTN TEXT,
                        CustomerAddress TEXT,
                        CustomerPhone TEXT,
                        CustomerEmail TEXT,
                        CustomerType TEXT DEFAULT 'Unregistered',
                        Amount DECIMAL NOT NULL,
                        TotalAmount DECIMAL DEFAULT 0,
                        TaxAmount DECIMAL DEFAULT 0,
                        DiscountAmount DECIMAL DEFAULT 0,
                        InvoiceDate DATETIME,
                        PaymentMode TEXT DEFAULT 'Cash',
                        Status TEXT DEFAULT 'Pending',
                        FBR_IRN TEXT,
                        FBR_QRCode TEXT,
                        UploadDate DATETIME,
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        ModifiedDate DATETIME,
                        ErrorMessage TEXT,
                        InvoiceType TEXT DEFAULT 'Invoice',
                        UNIQUE(CompanyName, QuickBooksInvoiceId)
                    );";

                string createInvoiceItemsTable = @"
                    CREATE TABLE IF NOT EXISTS InvoiceItems (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        InvoiceId INTEGER NOT NULL,
                        ItemName TEXT NOT NULL,
                        ItemDescription TEXT,
                        Quantity INTEGER NOT NULL,
                        UnitPrice DECIMAL NOT NULL,
                        TotalPrice DECIMAL NOT NULL,
                        NetAmount DECIMAL DEFAULT 0,
                        TaxRate DECIMAL DEFAULT 0,
                        SalesTaxAmount DECIMAL DEFAULT 0,
                        TotalValue DECIMAL DEFAULT 0,
                        HSCode TEXT,
                        UnitOfMeasure TEXT,
                        RetailPrice DECIMAL DEFAULT 0,
                        ExtraTax DECIMAL DEFAULT 0,
                        FurtherTax DECIMAL DEFAULT 0,
                        FedPayable DECIMAL DEFAULT 0,
                        SalesTaxWithheldAtSource DECIMAL DEFAULT 0,
                        Discount DECIMAL DEFAULT 0,
                        SaleType TEXT,
                        SroScheduleNo TEXT,
                        SroItemSerialNo TEXT,
                        FOREIGN KEY (InvoiceId) REFERENCES Invoices(Id) ON DELETE CASCADE
                    );";

                string createTransactionTypesTable = @"
                    CREATE TABLE IF NOT EXISTS TransactionTypes (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        TransactionTypeId INTEGER NOT NULL UNIQUE,
                        TransactionDesc TEXT NOT NULL,
                        LastUpdated TEXT NOT NULL
                    );
                    CREATE INDEX IF NOT EXISTS idx_transaction_type_id ON TransactionTypes(TransactionTypeId);
                    CREATE INDEX IF NOT EXISTS idx_invoiceitems_invoiceid ON InvoiceItems(InvoiceId);
                ";

                ExecuteNonQuery(conn, createCompaniesTable);
                ExecuteNonQuery(conn, createInvoicesTable);
                ExecuteNonQuery(conn, createInvoiceItemsTable);
                ExecuteNonQuery(conn, createTransactionTypesTable);

                Console.WriteLine("✅ Database initialized successfully");
            }
            MigrateDatabase();
        }

        private void ExecuteNonQuery(SQLiteConnection conn, string sql)
        {
            using var cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        private T ExecuteWithRetry<T>(Func<T> operation, string operationName)
        {
            int retryCount = 0;
            Exception lastException = null;

            while (retryCount < MaxRetries)
            {
                try
                {
                    lock (_lockObject)
                    {
                        return operation();
                    }
                }
                catch (SQLiteException ex) when (ex.Message.Contains("database is locked"))
                {
                    lastException = ex;
                    retryCount++;

                    if (retryCount < MaxRetries)
                    {
                        Console.WriteLine($"Database locked during {operationName}, retry {retryCount}/{MaxRetries}...");
                        Thread.Sleep(RetryDelayMs * retryCount);
                    }
                }
            }

            throw new Exception($"Failed to complete {operationName} after {MaxRetries} retries", lastException);
        }

        #region Companies CRUD

        public Company GetCompanyByName(string companyName)
        {
            return ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand("SELECT * FROM Companies WHERE CompanyName = @name", conn);
                cmd.Parameters.AddWithValue("@name", companyName);

                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    return new Company
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        CompanyName = reader["CompanyName"].ToString(),
                        FBRToken = reader["FBRToken"].ToString(),
                        SellerNTN = reader["SellerNTN"]?.ToString(),
                        SellerAddress = reader["SellerAddress"]?.ToString(),
                        SellerProvince = reader["SellerProvince"]?.ToString(),
                        SellerPhone = reader["SellerPhone"]?.ToString(),
                        SellerEmail = reader["SellerEmail"]?.ToString(),
                        LogoImage = reader["LogoImage"] != DBNull.Value ? (byte[])reader["LogoImage"] : null,
                        Environment = reader["Environment"] != DBNull.Value ? reader["Environment"].ToString() : "Sandbox", // ✅ ADD THIS
                        CreatedDate = Convert.ToDateTime(reader["CreatedDate"]),
                        ModifiedDate = reader["ModifiedDate"] != DBNull.Value ? Convert.ToDateTime(reader["ModifiedDate"]) : (DateTime?)null
                    };
                }

                return null;
            }, "GetCompanyByName");
        }

        public List<Company> GetAllCompanies()
        {
            return ExecuteWithRetry(() =>
            {
                var companies = new List<Company>();

                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand("SELECT * FROM Companies ORDER BY CompanyName", conn);
                using var reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    companies.Add(new Company
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        CompanyName = reader["CompanyName"].ToString(),
                        FBRToken = reader["FBRToken"].ToString(),
                        SellerNTN = reader["SellerNTN"]?.ToString(),
                        SellerAddress = reader["SellerAddress"]?.ToString(),
                        SellerProvince = reader["SellerProvince"]?.ToString(),
                        SellerPhone = reader["SellerPhone"]?.ToString(),
                        SellerEmail = reader["SellerEmail"]?.ToString(),
                        Environment = reader["Environment"] != DBNull.Value ? reader["Environment"].ToString() : "Sandbox", // ✅ ADD THIS
                        CreatedDate = Convert.ToDateTime(reader["CreatedDate"]),
                        ModifiedDate = reader["ModifiedDate"] != DBNull.Value ? Convert.ToDateTime(reader["ModifiedDate"]) : (DateTime?)null
                    });
                }

                return companies;
            }, "GetAllCompanies");
        }

        public void SaveCompany(Company company)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var checkCmd = new SQLiteCommand("SELECT Id FROM Companies WHERE CompanyName = @name", conn);
                checkCmd.Parameters.AddWithValue("@name", company.CompanyName);
                var existingId = checkCmd.ExecuteScalar();

                SQLiteCommand cmd;

                if (existingId != null)
                {
                    cmd = new SQLiteCommand(@"
                        UPDATE Companies 
                        SET FBRToken = @token, 
                            SellerNTN = @ntn,
                            SellerAddress = @address,
                            SellerProvince = @province,
                            SellerPhone = @phone,
                            SellerEmail = @email,
                            LogoImage = @logoImage,
                            Environment = @environment,
                            ModifiedDate = @modified
                        WHERE CompanyName = @name", conn);

                    cmd.Parameters.AddWithValue("@modified", DateTime.Now);
                }
                else
                {
                    cmd = new SQLiteCommand(@"
                        INSERT INTO Companies (CompanyName, FBRToken, SellerNTN, SellerAddress, SellerProvince, SellerPhone, SellerEmail, LogoImage, Environment, CreatedDate)
                        VALUES (@name, @token, @ntn, @address, @province, @phone, @email, @logoImage, @environment, @created)", conn);

                    cmd.Parameters.AddWithValue("@created", DateTime.Now);
                }

                cmd.Parameters.AddWithValue("@name", company.CompanyName);
                cmd.Parameters.AddWithValue("@token", company.FBRToken);
                cmd.Parameters.AddWithValue("@environment", company.Environment ?? "Sandbox");
                cmd.Parameters.AddWithValue("@ntn", company.SellerNTN ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@address", company.SellerAddress ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@province", company.SellerProvince ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@phone", company.SellerPhone ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@email", company.SellerEmail ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@logoImage", company.LogoImage ?? (object)DBNull.Value);

                cmd.ExecuteNonQuery();
                return true;
            }, "SaveCompany");
        }

        public void DeleteCompany(string companyName)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand("DELETE FROM Companies WHERE CompanyName = @name", conn);
                cmd.Parameters.AddWithValue("@name", companyName);
                cmd.ExecuteNonQuery();

                return true;
            }, "DeleteCompany");
        }



        #endregion

        #region Invoices CRUD

        public void SaveInvoice(Invoice invoice)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var checkCmd = new SQLiteCommand(@"
                    SELECT Id FROM Invoices 
                    WHERE QuickBooksInvoiceId = @qbId 
                    AND CompanyName = @company", conn);

                checkCmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                checkCmd.Parameters.AddWithValue("@company", invoice.CompanyName);

                var existingId = checkCmd.ExecuteScalar();

                if (existingId == null)
                {
                    var cmd = new SQLiteCommand(@"
                        INSERT INTO Invoices (
                            CompanyName, QuickBooksInvoiceId, InvoiceNumber, 
                            CustomerName, CustomerNTN, CustomerAddress, CustomerPhone, CustomerEmail, CustomerType,
                            Amount, TotalAmount, TaxAmount, DiscountAmount,
                            InvoiceDate, PaymentMode, Status, 
                            FBR_IRN, FBR_QRCode, UploadDate, ErrorMessage, InvoiceType, CreatedDate
                        ) VALUES (
                            @company, @qbId, @invNum, 
                            @custName, @custNTN, @custAddr, @custPhone, @custEmail, @custType,
                            @amount, @totalAmount, @taxAmount, @discountAmount,
                            @invDate, @paymentMode, @status, 
                            @irn, @qr, @uploadDate, @error, @invoiceType, @created
                        )", conn);

                    cmd.Parameters.AddWithValue("@company", invoice.CompanyName);
                    cmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                    cmd.Parameters.AddWithValue("@invNum", invoice.InvoiceNumber);
                    cmd.Parameters.AddWithValue("@custName", invoice.CustomerName);
                    cmd.Parameters.AddWithValue("@custNTN", invoice.CustomerNTN ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@custAddr", invoice.CustomerAddress ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@custPhone", invoice.CustomerPhone ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@custEmail", invoice.CustomerEmail ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@custType", invoice.CustomerType ?? "Unregistered");
                    cmd.Parameters.AddWithValue("@amount", invoice.Amount);
                    cmd.Parameters.AddWithValue("@totalAmount", invoice.TotalAmount);
                    cmd.Parameters.AddWithValue("@taxAmount", invoice.TaxAmount);
                    cmd.Parameters.AddWithValue("@discountAmount", invoice.DiscountAmount);
                    cmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@paymentMode", invoice.PaymentMode ?? "Cash");
                    cmd.Parameters.AddWithValue("@status", invoice.Status);
                    cmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@uploadDate", invoice.UploadDate ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@invoiceType", invoice.InvoiceType ?? "Invoice");
                    cmd.Parameters.AddWithValue("@created", invoice.CreatedDate);

                    cmd.ExecuteNonQuery();
                }

                return true;
            }, "SaveInvoice");
        }

        public void SaveInvoiceWithDetails(Invoice invoice)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                using var transaction = conn.BeginTransaction();
                try
                {
                    var checkCmd = new SQLiteCommand(@"
                        SELECT Id, Status, FBR_IRN, FBR_QRCode, UploadDate, ErrorMessage 
                        FROM Invoices 
                        WHERE QuickBooksInvoiceId = @qbId 
                        AND CompanyName = @company", conn, transaction);

                    checkCmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                    checkCmd.Parameters.AddWithValue("@company", invoice.CompanyName);

                    int invoiceId;
                    bool isExisting = false;

                    using (var reader = checkCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            invoiceId = Convert.ToInt32(reader["Id"]);
                            invoice.Id = invoiceId;
                            isExisting = true;
                        }
                        else
                        {
                            invoiceId = -1;
                        }
                    }

                    if (isExisting && invoiceId > 0)
                    {
                        var updateCmd = new SQLiteCommand(@"
                            UPDATE Invoices SET
                                InvoiceNumber = @invNum,
                                CustomerName = @custName,
                                CustomerNTN = @custNTN,
                                CustomerAddress = @custAddr,
                                CustomerPhone = @custPhone,
                                CustomerEmail = @custEmail,
                                CustomerType = @custType,
                                Amount = @amount,
                                TotalAmount = @totalAmount,
                                TaxAmount = @taxAmount,
                                DiscountAmount = @discountAmount,
                                InvoiceDate = @invDate,
                                PaymentMode = @paymentMode,
                                Status = @status,
                                FBR_IRN = @irn,
                                FBR_QRCode = @qr,
                                UploadDate = @uploadDate,
                                ErrorMessage = @error,
                                InvoiceType = @invoiceType,
                                ModifiedDate = @modified
                            WHERE Id = @id", conn, transaction);

                        updateCmd.Parameters.AddWithValue("@id", invoiceId);
                        updateCmd.Parameters.AddWithValue("@invNum", invoice.InvoiceNumber);
                        updateCmd.Parameters.AddWithValue("@custName", invoice.CustomerName);
                        updateCmd.Parameters.AddWithValue("@custNTN", invoice.CustomerNTN ?? (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@custAddr", invoice.CustomerAddress ?? (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@custPhone", invoice.CustomerPhone ?? (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@custEmail", invoice.CustomerEmail ?? (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@custType", invoice.CustomerType ?? "Unregistered");
                        updateCmd.Parameters.AddWithValue("@amount", invoice.Amount);
                        updateCmd.Parameters.AddWithValue("@totalAmount", invoice.TotalAmount);
                        updateCmd.Parameters.AddWithValue("@taxAmount", invoice.TaxAmount);
                        updateCmd.Parameters.AddWithValue("@discountAmount", invoice.DiscountAmount);
                        updateCmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@paymentMode", invoice.PaymentMode ?? "Cash");
                        updateCmd.Parameters.AddWithValue("@status", invoice.Status ?? "Pending");
                        updateCmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@uploadDate", invoice.UploadDate ?? (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@invoiceType", invoice.InvoiceType ?? "Invoice");
                        updateCmd.Parameters.AddWithValue("@modified", DateTime.Now);

                        updateCmd.ExecuteNonQuery();

                        System.Diagnostics.Debug.WriteLine($"✅ Updated invoice {invoice.InvoiceNumber} (ID: {invoiceId})");
                    }
                    else
                    {
                        var insertCmd = new SQLiteCommand(@"
                            INSERT INTO Invoices (
                                CompanyName, QuickBooksInvoiceId, InvoiceNumber,
                                CustomerName, CustomerNTN, CustomerAddress, CustomerPhone, CustomerEmail, CustomerType,
                                Amount, TotalAmount, TaxAmount, DiscountAmount,
                                InvoiceDate, PaymentMode, Status,
                                FBR_IRN, FBR_QRCode, UploadDate, ErrorMessage, InvoiceType, CreatedDate
                            ) VALUES (
                                @company, @qbId, @invNum,
                                @custName, @custNTN, @custAddr, @custPhone, @custEmail, @custType,
                                @amount, @totalAmount, @taxAmount, @discountAmount,
                                @invDate, @paymentMode, @status,
                                @irn, @qr, @uploadDate, @error, @invoiceType, @created
                            );
                            SELECT last_insert_rowid();", conn, transaction);

                        insertCmd.Parameters.AddWithValue("@company", invoice.CompanyName);
                        insertCmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                        insertCmd.Parameters.AddWithValue("@invNum", invoice.InvoiceNumber);
                        insertCmd.Parameters.AddWithValue("@custName", invoice.CustomerName);
                        insertCmd.Parameters.AddWithValue("@custNTN", invoice.CustomerNTN ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@custAddr", invoice.CustomerAddress ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@custPhone", invoice.CustomerPhone ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@custEmail", invoice.CustomerEmail ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@custType", invoice.CustomerType ?? "Unregistered");
                        insertCmd.Parameters.AddWithValue("@amount", invoice.Amount);
                        insertCmd.Parameters.AddWithValue("@totalAmount", invoice.TotalAmount);
                        insertCmd.Parameters.AddWithValue("@taxAmount", invoice.TaxAmount);
                        insertCmd.Parameters.AddWithValue("@discountAmount", invoice.DiscountAmount);
                        insertCmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@paymentMode", invoice.PaymentMode ?? "Cash");
                        insertCmd.Parameters.AddWithValue("@status", invoice.Status ?? "Pending");
                        insertCmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@uploadDate", invoice.UploadDate ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@invoiceType", invoice.InvoiceType ?? "Invoice");
                        insertCmd.Parameters.AddWithValue("@created", DateTime.Now);

                        invoiceId = Convert.ToInt32(insertCmd.ExecuteScalar());
                        invoice.Id = invoiceId;

                        System.Diagnostics.Debug.WriteLine($"✅ Inserted new invoice {invoice.InvoiceNumber} (ID: {invoiceId})");
                    }

                    // Delete existing invoice items
                    var deleteItemsCmd = new SQLiteCommand(
                        "DELETE FROM InvoiceItems WHERE InvoiceId = @invoiceId",
                        conn, transaction);
                    deleteItemsCmd.Parameters.AddWithValue("@invoiceId", invoiceId);
                    int deletedCount = deleteItemsCmd.ExecuteNonQuery();

                    System.Diagnostics.Debug.WriteLine($"🗑️ Deleted {deletedCount} old items for invoice {invoice.InvoiceNumber}");

                    // Insert fresh invoice items
                    if (invoice.Items != null && invoice.Items.Count > 0)
                    {
                        foreach (var item in invoice.Items)
                        {
                            var itemCmd = new SQLiteCommand(@"
                                INSERT INTO InvoiceItems (
                                    InvoiceId, ItemName, ItemDescription, Quantity,
                                    UnitPrice, TotalPrice, NetAmount,
                                    TaxRate, SalesTaxAmount, TotalValue,
                                    HSCode, UnitOfMeasure, RetailPrice,
                                    ExtraTax, FurtherTax, FedPayable,
                                    SalesTaxWithheldAtSource, Discount,
                                    SaleType, SroScheduleNo, SroItemSerialNo
                                ) VALUES (
                                    @invoiceId, @itemName, @itemDesc, @qty,
                                    @unitPrice, @totalPrice, @netAmount,
                                    @taxRate, @salesTax, @totalValue,
                                    @hsCode, @uom, @retailPrice,
                                    @extraTax, @furtherTax, @fedPayable,
                                    @salesTaxWithheld, @discount,
                                    @saleType, @sroSchedule, @sroItem
                                )", conn, transaction);

                            itemCmd.Parameters.AddWithValue("@invoiceId", invoiceId);
                            itemCmd.Parameters.AddWithValue("@itemName", item.ItemName ?? "");
                            itemCmd.Parameters.AddWithValue("@itemDesc", item.ItemDescription ?? "");
                            itemCmd.Parameters.AddWithValue("@qty", item.Quantity);
                            itemCmd.Parameters.AddWithValue("@unitPrice", item.UnitPrice);
                            itemCmd.Parameters.AddWithValue("@totalPrice", item.TotalPrice);
                            itemCmd.Parameters.AddWithValue("@netAmount", item.NetAmount);
                            itemCmd.Parameters.AddWithValue("@taxRate", item.TaxRate);
                            itemCmd.Parameters.AddWithValue("@salesTax", item.SalesTaxAmount);
                            itemCmd.Parameters.AddWithValue("@totalValue", item.TotalValue);
                            itemCmd.Parameters.AddWithValue("@hsCode", item.HSCode ?? "");
                            itemCmd.Parameters.AddWithValue("@uom", item.UnitOfMeasure ?? "");
                            itemCmd.Parameters.AddWithValue("@retailPrice", item.RetailPrice);
                            itemCmd.Parameters.AddWithValue("@extraTax", item.ExtraTax);
                            itemCmd.Parameters.AddWithValue("@furtherTax", item.FurtherTax);
                            itemCmd.Parameters.AddWithValue("@fedPayable", item.FedPayable);
                            itemCmd.Parameters.AddWithValue("@salesTaxWithheld", item.SalesTaxWithheldAtSource);
                            itemCmd.Parameters.AddWithValue("@discount", item.Discount);
                            itemCmd.Parameters.AddWithValue("@saleType", item.SaleType ?? "");
                            itemCmd.Parameters.AddWithValue("@sroSchedule", item.SroScheduleNo ?? "");
                            itemCmd.Parameters.AddWithValue("@sroItem", item.SroItemSerialNo ?? "");

                            itemCmd.ExecuteNonQuery();
                        }

                        System.Diagnostics.Debug.WriteLine($"✅ Inserted {invoice.Items.Count} fresh items for invoice {invoice.InvoiceNumber}");
                    }

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    throw new Exception($"Error saving invoice with details: {ex.Message}", ex);
                }

                return true;
            }, "SaveInvoiceWithDetails");
        }

        #region Optimized Batch Operations

        public void SaveInvoicesBatch(List<Invoice> invoices)
        {
            if (invoices == null || invoices.Count == 0)
                return;

            ExecuteWithRetry(() =>
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();

                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            foreach (var invoice in invoices)
                            {
                                SaveInvoiceWithDetailsInternal(invoice, connection, transaction);
                            }

                            transaction.Commit();
                            Console.WriteLine($"✅ Batch saved {invoices.Count} invoices in single transaction");
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            Console.WriteLine($"❌ Batch save failed: {ex.Message}");
                            throw;
                        }
                    }
                }
                return true;
            }, "SaveInvoicesBatch");
        }

        public List<Invoice> GetInvoicesWithDetails(string companyName)
        {
            return ExecuteWithRetry(() =>
            {
                var invoiceDict = new Dictionary<int, Invoice>();

                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();

                    string query = @"
                        SELECT 
                            i.Id, i.CompanyName, i.QuickBooksInvoiceId, i.InvoiceNumber,
                            i.CustomerName, i.CustomerNTN, i.CustomerAddress, i.CustomerPhone,
                            i.CustomerEmail, i.CustomerType, i.Amount, i.TotalAmount,
                            i.TaxAmount, i.DiscountAmount, i.InvoiceDate, i.PaymentMode,
                            i.Status, i.FBR_IRN, i.FBR_QRCode, i.UploadDate,
                            i.CreatedDate, i.ModifiedDate, i.ErrorMessage, i.InvoiceType,
                            ii.Id as ItemId, ii.InvoiceId as ItemInvoiceId,
                            ii.ItemName, ii.ItemDescription, ii.Quantity, ii.UnitPrice,
                            ii.TotalPrice, ii.NetAmount, ii.TaxRate, ii.SalesTaxAmount,
                            ii.TotalValue, ii.HSCode, ii.UnitOfMeasure, ii.RetailPrice,
                            ii.ExtraTax, ii.FurtherTax, ii.FedPayable,
                            ii.SalesTaxWithheldAtSource, ii.Discount, ii.SaleType,
                            ii.SroScheduleNo, ii.SroItemSerialNo
                        FROM Invoices i
                        LEFT JOIN InvoiceItems ii ON i.Id = ii.InvoiceId
                        WHERE i.CompanyName = @CompanyName
                        ORDER BY i.CreatedDate DESC, ii.Id";

                    using (var cmd = new SQLiteCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@CompanyName", companyName);

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int invoiceId = Convert.ToInt32(reader["Id"]);

                                if (!invoiceDict.ContainsKey(invoiceId))
                                {
                                    invoiceDict[invoiceId] = new Invoice
                                    {
                                        Id = invoiceId,
                                        CompanyName = reader["CompanyName"].ToString(),
                                        QuickBooksInvoiceId = reader["QuickBooksInvoiceId"].ToString(),
                                        InvoiceNumber = reader["InvoiceNumber"].ToString(),
                                        CustomerName = reader["CustomerName"].ToString(),
                                        CustomerNTN = reader["CustomerNTN"]?.ToString(),
                                        CustomerAddress = reader["CustomerAddress"]?.ToString(),
                                        CustomerPhone = reader["CustomerPhone"]?.ToString(),
                                        CustomerEmail = reader["CustomerEmail"]?.ToString(),
                                        CustomerType = reader["CustomerType"] != DBNull.Value ? reader["CustomerType"].ToString() : "Unregistered",
                                        Amount = Convert.ToDecimal(reader["Amount"]),
                                        TotalAmount = reader["TotalAmount"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAmount"]) : 0,
                                        TaxAmount = reader["TaxAmount"] != DBNull.Value ? Convert.ToDecimal(reader["TaxAmount"]) : 0,
                                        DiscountAmount = reader["DiscountAmount"] != DBNull.Value ? Convert.ToDecimal(reader["DiscountAmount"]) : 0,
                                        InvoiceDate = reader["InvoiceDate"] != DBNull.Value ? Convert.ToDateTime(reader["InvoiceDate"]) : DateTime.MinValue,
                                        PaymentMode = reader["PaymentMode"]?.ToString() ?? "Cash",
                                        Status = reader["Status"].ToString(),
                                        FBR_IRN = reader["FBR_IRN"]?.ToString(),
                                        FBR_QRCode = reader["FBR_QRCode"]?.ToString(),
                                        UploadDate = reader["UploadDate"] != DBNull.Value ? (DateTime?)Convert.ToDateTime(reader["UploadDate"]) : null,
                                        ErrorMessage = reader["ErrorMessage"]?.ToString(),
                                        CreatedDate = Convert.ToDateTime(reader["CreatedDate"]),
                                        ModifiedDate = reader["ModifiedDate"] != DBNull.Value ? (DateTime?)Convert.ToDateTime(reader["ModifiedDate"]) : null,
                                        InvoiceType = reader["InvoiceType"] != DBNull.Value ? reader["InvoiceType"].ToString() : "Invoice",
                                        Items = new List<InvoiceItem>()
                                    };
                                }

                                if (reader["ItemId"] != DBNull.Value)
                                {
                                    invoiceDict[invoiceId].Items.Add(new InvoiceItem
                                    {
                                        Id = Convert.ToInt32(reader["ItemId"]),
                                        InvoiceId = Convert.ToInt32(reader["ItemInvoiceId"]),
                                        ItemName = reader["ItemName"].ToString(),
                                        ItemDescription = reader["ItemDescription"]?.ToString() ?? "",
                                        Quantity = Convert.ToInt32(reader["Quantity"]),
                                        UnitPrice = Convert.ToDecimal(reader["UnitPrice"]),
                                        TotalPrice = Convert.ToDecimal(reader["TotalPrice"]),
                                        NetAmount = reader["NetAmount"] != DBNull.Value ? Convert.ToDecimal(reader["NetAmount"]) : 0,
                                        TaxRate = Convert.ToDecimal(reader["TaxRate"]),
                                        SalesTaxAmount = Convert.ToDecimal(reader["SalesTaxAmount"]),
                                        TotalValue = Convert.ToDecimal(reader["TotalValue"]),
                                        HSCode = reader["HSCode"]?.ToString() ?? "",
                                        UnitOfMeasure = reader["UnitOfMeasure"]?.ToString() ?? "",
                                        RetailPrice = reader["RetailPrice"] != DBNull.Value ? Convert.ToDecimal(reader["RetailPrice"]) : 0,
                                        ExtraTax = reader["ExtraTax"] != DBNull.Value ? Convert.ToDecimal(reader["ExtraTax"]) : 0,
                                        FurtherTax = reader["FurtherTax"] != DBNull.Value ? Convert.ToDecimal(reader["FurtherTax"]) : 0,
                                        FedPayable = reader["FedPayable"] != DBNull.Value ? Convert.ToDecimal(reader["FedPayable"]) : 0,
                                        SalesTaxWithheldAtSource = reader["SalesTaxWithheldAtSource"] != DBNull.Value ? Convert.ToDecimal(reader["SalesTaxWithheldAtSource"]) : 0,
                                        Discount = reader["Discount"] != DBNull.Value ? Convert.ToDecimal(reader["Discount"]) : 0,
                                        SaleType = reader["SaleType"]?.ToString() ?? "",
                                        SroScheduleNo = reader["SroScheduleNo"]?.ToString() ?? "",
                                        SroItemSerialNo = reader["SroItemSerialNo"]?.ToString() ?? ""
                                    });
                                }
                            }
                        }
                    }
                }

                return invoiceDict.Values.ToList();
            }, "GetInvoicesWithDetails");
        }

        private void SaveInvoiceWithDetailsInternal(Invoice invoice, SQLiteConnection connection, SQLiteTransaction transaction)
        {
            string checkQuery = @"
                SELECT Id, Status, FBR_IRN, FBR_QRCode, UploadDate 
                FROM Invoices 
                WHERE QuickBooksInvoiceId = @QuickBooksInvoiceId 
                AND CompanyName = @CompanyName";

            int invoiceId;
            bool isExisting = false;

            using (var checkCmd = new SQLiteCommand(checkQuery, connection, transaction))
            {
                checkCmd.Parameters.AddWithValue("@QuickBooksInvoiceId", invoice.QuickBooksInvoiceId);
                checkCmd.Parameters.AddWithValue("@CompanyName", invoice.CompanyName);

                using (var reader = checkCmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        invoiceId = Convert.ToInt32(reader["Id"]);
                        invoice.Id = invoiceId;
                        isExisting = true;
                    }
                    else
                    {
                        invoiceId = -1;
                    }
                }
            }

            if (isExisting && invoiceId > 0)
            {
                string updateQuery = @"
                    UPDATE Invoices SET
                        InvoiceNumber = @invNum,
                        CustomerName = @custName,
                        CustomerNTN = @custNTN,
                        CustomerAddress = @custAddr,
                        CustomerPhone = @custPhone,
                        CustomerEmail = @custEmail,
                        CustomerType = @custType,
                        Amount = @amount,
                        TotalAmount = @totalAmount,
                        TaxAmount = @taxAmount,
                        DiscountAmount = @discountAmount,
                        InvoiceDate = @invDate,
                        PaymentMode = @paymentMode,
                        Status = @status,
                        FBR_IRN = @irn,
                        FBR_QRCode = @qr,
                        UploadDate = @uploadDate,
                        ErrorMessage = @error,
                        InvoiceType = @invoiceType,
                        ModifiedDate = @modified
                    WHERE Id = @id";

                using (var updateCmd = new SQLiteCommand(updateQuery, connection, transaction))
                {
                    updateCmd.Parameters.AddWithValue("@id", invoiceId);
                    updateCmd.Parameters.AddWithValue("@invNum", invoice.InvoiceNumber);
                    updateCmd.Parameters.AddWithValue("@custName", invoice.CustomerName);
                    updateCmd.Parameters.AddWithValue("@custNTN", invoice.CustomerNTN ?? (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@custAddr", invoice.CustomerAddress ?? (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@custPhone", invoice.CustomerPhone ?? (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@custEmail", invoice.CustomerEmail ?? (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@custType", invoice.CustomerType ?? "Unregistered");
                    updateCmd.Parameters.AddWithValue("@amount", invoice.Amount);
                    updateCmd.Parameters.AddWithValue("@totalAmount", invoice.TotalAmount);
                    updateCmd.Parameters.AddWithValue("@taxAmount", invoice.TaxAmount);
                    updateCmd.Parameters.AddWithValue("@discountAmount", invoice.DiscountAmount);
                    updateCmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@paymentMode", invoice.PaymentMode ?? "Cash");
                    updateCmd.Parameters.AddWithValue("@status", invoice.Status ?? "Pending");
                    updateCmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@uploadDate", invoice.UploadDate ?? (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
                    updateCmd.Parameters.AddWithValue("@invoiceType", invoice.InvoiceType ?? "Invoice");
                    updateCmd.Parameters.AddWithValue("@modified", DateTime.Now);

                    updateCmd.ExecuteNonQuery();
                }
            }
            else
            {
                string insertQuery = @"
                    INSERT INTO Invoices (
                        CompanyName, QuickBooksInvoiceId, InvoiceNumber,
                        CustomerName, CustomerNTN, CustomerAddress, CustomerPhone, CustomerEmail, CustomerType,
                        Amount, TotalAmount, TaxAmount, DiscountAmount,
                        InvoiceDate, PaymentMode, Status,
                        FBR_IRN, FBR_QRCode, UploadDate, ErrorMessage, InvoiceType, CreatedDate
                    ) VALUES (
                        @company, @qbId, @invNum,
                        @custName, @custNTN, @custAddr, @custPhone, @custEmail, @custType,
                        @amount, @totalAmount, @taxAmount, @discountAmount,
                        @invDate, @paymentMode, @status,
                        @irn, @qr, @uploadDate, @error, @invoiceType, @created
                    );
                    SELECT last_insert_rowid();";

                using (var insertCmd = new SQLiteCommand(insertQuery, connection, transaction))
                {
                    insertCmd.Parameters.AddWithValue("@company", invoice.CompanyName);
                    insertCmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                    insertCmd.Parameters.AddWithValue("@invNum", invoice.InvoiceNumber);
                    insertCmd.Parameters.AddWithValue("@custName", invoice.CustomerName);
                    insertCmd.Parameters.AddWithValue("@custNTN", invoice.CustomerNTN ?? (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@custAddr", invoice.CustomerAddress ?? (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@custPhone", invoice.CustomerPhone ?? (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@custEmail", invoice.CustomerEmail ?? (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@custType", invoice.CustomerType ?? "Unregistered");
                    insertCmd.Parameters.AddWithValue("@amount", invoice.Amount);
                    insertCmd.Parameters.AddWithValue("@totalAmount", invoice.TotalAmount);
                    insertCmd.Parameters.AddWithValue("@taxAmount", invoice.TaxAmount);
                    insertCmd.Parameters.AddWithValue("@discountAmount", invoice.DiscountAmount);
                    insertCmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@paymentMode", invoice.PaymentMode ?? "Cash");
                    insertCmd.Parameters.AddWithValue("@status", invoice.Status ?? "Pending");
                    insertCmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@uploadDate", invoice.UploadDate ?? (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@invoiceType", invoice.InvoiceType ?? "Invoice");
                    insertCmd.Parameters.AddWithValue("@created", DateTime.Now);

                    invoiceId = Convert.ToInt32(insertCmd.ExecuteScalar());
                    invoice.Id = invoiceId;
                }
            }

            // Delete existing items
            string deleteItemsQuery = "DELETE FROM InvoiceItems WHERE InvoiceId = @invoiceId";
            using (var deleteCmd = new SQLiteCommand(deleteItemsQuery, connection, transaction))
            {
                deleteCmd.Parameters.AddWithValue("@invoiceId", invoiceId);
                deleteCmd.ExecuteNonQuery();
            }

            // Insert fresh items
            if (invoice.Items != null && invoice.Items.Count > 0)
            {
                string insertItemQuery = @"
                    INSERT INTO InvoiceItems (
                        InvoiceId, ItemName, ItemDescription, Quantity,
                        UnitPrice, TotalPrice, NetAmount,
                        TaxRate, SalesTaxAmount, TotalValue,
                        HSCode, UnitOfMeasure, RetailPrice,
                        ExtraTax, FurtherTax, FedPayable,
                        SalesTaxWithheldAtSource, Discount,
                        SaleType, SroScheduleNo, SroItemSerialNo
                    ) VALUES (
                        @invoiceId, @itemName, @itemDesc, @qty,
                        @unitPrice, @totalPrice, @netAmount,
                        @taxRate, @salesTax, @totalValue,
                        @hsCode, @uom, @retailPrice,
                        @extraTax, @furtherTax, @fedPayable,
                        @salesTaxWithheld, @discount,
                        @saleType, @sroSchedule, @sroItem
                    )";

                foreach (var item in invoice.Items)
                {
                    using (var itemCmd = new SQLiteCommand(insertItemQuery, connection, transaction))
                    {
                        itemCmd.Parameters.AddWithValue("@invoiceId", invoiceId);
                        itemCmd.Parameters.AddWithValue("@itemName", item.ItemName ?? "");
                        itemCmd.Parameters.AddWithValue("@itemDesc", item.ItemDescription ?? "");
                        itemCmd.Parameters.AddWithValue("@qty", item.Quantity);
                        itemCmd.Parameters.AddWithValue("@unitPrice", item.UnitPrice);
                        itemCmd.Parameters.AddWithValue("@totalPrice", item.TotalPrice);
                        itemCmd.Parameters.AddWithValue("@netAmount", item.NetAmount);
                        itemCmd.Parameters.AddWithValue("@taxRate", item.TaxRate);
                        itemCmd.Parameters.AddWithValue("@salesTax", item.SalesTaxAmount);
                        itemCmd.Parameters.AddWithValue("@totalValue", item.TotalValue);
                        itemCmd.Parameters.AddWithValue("@hsCode", item.HSCode ?? "");
                        itemCmd.Parameters.AddWithValue("@uom", item.UnitOfMeasure ?? "");
                        itemCmd.Parameters.AddWithValue("@retailPrice", item.RetailPrice);
                        itemCmd.Parameters.AddWithValue("@extraTax", item.ExtraTax);
                        itemCmd.Parameters.AddWithValue("@furtherTax", item.FurtherTax);
                        itemCmd.Parameters.AddWithValue("@fedPayable", item.FedPayable);
                        itemCmd.Parameters.AddWithValue("@salesTaxWithheld", item.SalesTaxWithheldAtSource);
                        itemCmd.Parameters.AddWithValue("@discount", item.Discount);
                        itemCmd.Parameters.AddWithValue("@saleType", item.SaleType ?? "");
                        itemCmd.Parameters.AddWithValue("@sroSchedule", item.SroScheduleNo ?? "");
                        itemCmd.Parameters.AddWithValue("@sroItem", item.SroItemSerialNo ?? "");

                        itemCmd.ExecuteNonQuery();
                    }
                }
            }
        }

        #endregion

        public List<Invoice> GetInvoices(string companyName)
        {
            return ExecuteWithRetry(() =>
            {
                var invoices = new List<Invoice>();

                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    SELECT * FROM Invoices 
                    WHERE CompanyName = @company 
                    ORDER BY CreatedDate DESC", conn);

                cmd.Parameters.AddWithValue("@company", companyName);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    invoices.Add(MapInvoiceFromReader(reader));
                }

                return invoices;
            }, "GetInvoices");
        }

        public Invoice GetInvoiceWithDetails(string qbInvoiceId, string companyName)
        {
            return ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    SELECT * FROM Invoices 
                    WHERE QuickBooksInvoiceId = @qbId 
                    AND CompanyName = @company", conn);

                cmd.Parameters.AddWithValue("@qbId", qbInvoiceId);
                cmd.Parameters.AddWithValue("@company", companyName);

                Invoice invoice = null;

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        invoice = MapInvoiceFromReader(reader);
                    }
                }

                if (invoice == null)
                    return null;

                var itemsCmd = new SQLiteCommand(
                    "SELECT * FROM InvoiceItems WHERE InvoiceId = @invoiceId",
                    conn);
                itemsCmd.Parameters.AddWithValue("@invoiceId", invoice.Id);

                invoice.Items = new List<InvoiceItem>();

                using (var itemsReader = itemsCmd.ExecuteReader())
                {
                    while (itemsReader.Read())
                    {
                        invoice.Items.Add(MapInvoiceItemFromReader(itemsReader));
                    }
                }

                return invoice;
            }, "GetInvoiceWithDetails");
        }

        public void UpdateInvoiceStatus(string qbInvoiceId, string status, string irn, string error)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    UPDATE Invoices 
                    SET Status = @status, 
                        FBR_IRN = COALESCE(@irn, FBR_IRN),
                        FBR_QRCode = COALESCE(@irn, FBR_QRCode),
                        UploadDate = @date, 
                        ErrorMessage = @error,
                        ModifiedDate = @modified
                    WHERE QuickBooksInvoiceId = @id", conn);

                cmd.Parameters.AddWithValue("@status", status);
                cmd.Parameters.AddWithValue("@irn", irn ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@date", status == "Uploaded" ? DateTime.Now : (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@error", error ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@modified", DateTime.Now);
                cmd.Parameters.AddWithValue("@id", qbInvoiceId);

                cmd.ExecuteNonQuery();
                return true;
            }, "UpdateInvoiceStatus");
        }

        #endregion

        #region Helper Methods

        private Invoice MapInvoiceFromReader(SQLiteDataReader reader)
        {
            return new Invoice
            {
                Id = Convert.ToInt32(reader["Id"]),
                CompanyName = reader["CompanyName"].ToString(),
                QuickBooksInvoiceId = reader["QuickBooksInvoiceId"].ToString(),
                InvoiceNumber = reader["InvoiceNumber"].ToString(),
                CustomerName = reader["CustomerName"].ToString(),
                CustomerNTN = reader["CustomerNTN"]?.ToString(),
                CustomerAddress = reader["CustomerAddress"]?.ToString(),
                CustomerPhone = reader["CustomerPhone"]?.ToString(),
                CustomerEmail = reader["CustomerEmail"]?.ToString(),
                CustomerType = reader["CustomerType"] != DBNull.Value ? reader["CustomerType"].ToString() : "Unregistered",
                Amount = Convert.ToDecimal(reader["Amount"]),
                TotalAmount = reader["TotalAmount"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAmount"]) : 0,
                TaxAmount = reader["TaxAmount"] != DBNull.Value ? Convert.ToDecimal(reader["TaxAmount"]) : 0,
                DiscountAmount = reader["DiscountAmount"] != DBNull.Value ? Convert.ToDecimal(reader["DiscountAmount"]) : 0,
                InvoiceDate = reader["InvoiceDate"] != DBNull.Value ? Convert.ToDateTime(reader["InvoiceDate"]) : DateTime.MinValue,
                PaymentMode = reader["PaymentMode"]?.ToString() ?? "Cash",
                Status = reader["Status"].ToString(),
                FBR_IRN = reader["FBR_IRN"]?.ToString(),
                FBR_QRCode = reader["FBR_QRCode"]?.ToString(),
                UploadDate = reader["UploadDate"] != DBNull.Value ? Convert.ToDateTime(reader["UploadDate"]) : (DateTime?)null,
                CreatedDate = Convert.ToDateTime(reader["CreatedDate"]),
                ModifiedDate = reader["ModifiedDate"] != DBNull.Value ? Convert.ToDateTime(reader["ModifiedDate"]) : (DateTime?)null,
                ErrorMessage = reader["ErrorMessage"]?.ToString(),
                InvoiceType = reader["InvoiceType"] != DBNull.Value ? reader["InvoiceType"].ToString() : "Invoice",
                Items = new List<InvoiceItem>()
            };
        }

        private InvoiceItem MapInvoiceItemFromReader(SQLiteDataReader reader)
        {
            return new InvoiceItem
            {
                Id = Convert.ToInt32(reader["Id"]),
                InvoiceId = Convert.ToInt32(reader["InvoiceId"]),
                ItemName = reader["ItemName"].ToString(),
                ItemDescription = reader["ItemDescription"]?.ToString() ?? "",
                Quantity = Convert.ToInt32(reader["Quantity"]),
                UnitPrice = Convert.ToDecimal(reader["UnitPrice"]),
                TotalPrice = Convert.ToDecimal(reader["TotalPrice"]),
                NetAmount = reader["NetAmount"] != DBNull.Value ? Convert.ToDecimal(reader["NetAmount"]) : 0,
                TaxRate = Convert.ToDecimal(reader["TaxRate"]),
                SalesTaxAmount = Convert.ToDecimal(reader["SalesTaxAmount"]),
                TotalValue = Convert.ToDecimal(reader["TotalValue"]),
                HSCode = reader["HSCode"]?.ToString() ?? "",
                UnitOfMeasure = reader["UnitOfMeasure"]?.ToString() ?? "",
                RetailPrice = reader["RetailPrice"] != DBNull.Value ? Convert.ToDecimal(reader["RetailPrice"]) : 0,
                ExtraTax = reader["ExtraTax"] != DBNull.Value ? Convert.ToDecimal(reader["ExtraTax"]) : 0,
                FurtherTax = reader["FurtherTax"] != DBNull.Value ? Convert.ToDecimal(reader["FurtherTax"]) : 0,
                FedPayable = reader["FedPayable"] != DBNull.Value ? Convert.ToDecimal(reader["FedPayable"]) : 0,
                SalesTaxWithheldAtSource = reader["SalesTaxWithheldAtSource"] != DBNull.Value ? Convert.ToDecimal(reader["SalesTaxWithheldAtSource"]) : 0,
                Discount = reader["Discount"] != DBNull.Value ? Convert.ToDecimal(reader["Discount"]) : 0,
                SaleType = reader["SaleType"]?.ToString() ?? "",
                SroScheduleNo = reader["SroScheduleNo"]?.ToString() ?? "",
                SroItemSerialNo = reader["SroItemSerialNo"]?.ToString() ?? ""
            };
        }

        #endregion

        #region Transaction Types Methods

        public void SaveTransactionTypes(List<TransactionType> transactionTypes)
        {
            ExecuteWithRetry(() =>
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();

                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            using (var deleteCmd = new SQLiteCommand("DELETE FROM TransactionTypes", connection, transaction))
                            {
                                deleteCmd.ExecuteNonQuery();
                            }

                            foreach (var transType in transactionTypes)
                            {
                                string sql = @"
                                    INSERT INTO TransactionTypes (TransactionTypeId, TransactionDesc, LastUpdated)
                                    VALUES (@TransactionTypeId, @TransactionDesc, @LastUpdated)";

                                using (var cmd = new SQLiteCommand(sql, connection, transaction))
                                {
                                    cmd.Parameters.AddWithValue("@TransactionTypeId", transType.TransactionTypeId);
                                    cmd.Parameters.AddWithValue("@TransactionDesc", transType.TransactionDesc);
                                    cmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                    cmd.ExecuteNonQuery();
                                }
                            }

                            transaction.Commit();
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
                return true;
            }, "SaveTransactionTypes");
        }

        public List<TransactionType> GetTransactionTypes()
        {
            return ExecuteWithRetry(() =>
            {
                var transactionTypes = new List<TransactionType>();

                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();

                    string sql = "SELECT Id, TransactionTypeId, TransactionDesc, LastUpdated FROM TransactionTypes ORDER BY TransactionDesc";

                    using (var command = new SQLiteCommand(sql, connection))
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            transactionTypes.Add(new TransactionType
                            {
                                Id = reader.GetInt32(0),
                                TransactionTypeId = reader.GetInt32(1),
                                TransactionDesc = reader.GetString(2),
                                LastUpdated = DateTime.Parse(reader.GetString(3))
                            });
                        }
                    }
                }

                return transactionTypes;
            }, "GetTransactionTypes");
        }

        public TransactionType GetTransactionTypeById(int transactionTypeId)
        {
            return ExecuteWithRetry(() =>
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();

                    string sql = "SELECT Id, TransactionTypeId, TransactionDesc, LastUpdated FROM TransactionTypes WHERE TransactionTypeId = @TransactionTypeId";

                    using (var command = new SQLiteCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@TransactionTypeId", transactionTypeId);

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new TransactionType
                                {
                                    Id = reader.GetInt32(0),
                                    TransactionTypeId = reader.GetInt32(1),
                                    TransactionDesc = reader.GetString(2),
                                    LastUpdated = DateTime.Parse(reader.GetString(3))
                                };
                            }
                        }
                    }
                }

                return null;
            }, "GetTransactionTypeById");
        }

        #endregion

        # region Migrations
        private void MigrateDatabase()
        {
            lock (_lockObject)
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                // Check if Environment column exists
                var checkColumnCmd = new SQLiteCommand(
                    "PRAGMA table_info(Companies);", conn);

                bool hasEnvironmentColumn = false;
                using (var reader = checkColumnCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string columnName = reader["name"].ToString();
                        if (columnName == "Environment")
                        {
                            hasEnvironmentColumn = true;
                            break;
                        }
                    }
                }

                // Add Environment column if it doesn't exist
                if (!hasEnvironmentColumn)
                {
                    Console.WriteLine("⚙️ Adding Environment column to Companies table...");
                    var alterCmd = new SQLiteCommand(
                        "ALTER TABLE Companies ADD COLUMN Environment TEXT DEFAULT 'Sandbox';", conn);
                    alterCmd.ExecuteNonQuery();
                    Console.WriteLine("✅ Environment column added successfully");
                }
            }
        }
        #endregion


    }
}