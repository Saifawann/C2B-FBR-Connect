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

        public DatabaseService()
        {
            _dbPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "FBRInvoicing",
                "fbr_invoicing.db"
            );

            var dir = Path.GetDirectoryName(_dbPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            // SOLUTION 1: Enhanced connection string with timeout and journal mode
            _connectionString = $"Data Source={_dbPath};Version=3;Journal Mode=WAL;Busy Timeout=5000;Pooling=True;Max Pool Size=100;";

            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            lock (_lockObject)
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                // SOLUTION 2: Enable WAL mode for better concurrency
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
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        ModifiedDate DATETIME
                    );";

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
                        UNIQUE(CompanyName, QuickBooksInvoiceId)
                    );";

                string createInvoiceItemsTable = @"
                    CREATE TABLE IF NOT EXISTS InvoiceItems (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        InvoiceId INTEGER NOT NULL,
                        ItemName TEXT NOT NULL,
                        Quantity INTEGER NOT NULL,
                        UnitPrice DECIMAL NOT NULL,
                        TotalPrice DECIMAL NOT NULL,
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
                        FOREIGN KEY (InvoiceId) REFERENCES Invoices(Id) ON DELETE CASCADE
                    );";

                ExecuteNonQuery(conn, createCompaniesTable);
                ExecuteNonQuery(conn, createInvoicesTable);
                ExecuteNonQuery(conn, createInvoiceItemsTable);
            }
        }

        private void ExecuteNonQuery(SQLiteConnection conn, string sql)
        {
            using var cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        // SOLUTION 3: Retry mechanism wrapper
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
                        Thread.Sleep(RetryDelayMs * retryCount); // Exponential backoff
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
                        CreatedDate = Convert.ToDateTime(reader["CreatedDate"]),
                        ModifiedDate = reader["ModifiedDate"] != DBNull.Value
                            ? Convert.ToDateTime(reader["ModifiedDate"])
                            : (DateTime?)null
                    };
                }
                return null;
            }, "GetCompanyByName");
        }

        public void SaveCompany(Company company)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    INSERT OR REPLACE INTO Companies 
                    (CompanyName, FBRToken, SellerNTN, SellerAddress, SellerProvince, 
                     SellerPhone, SellerEmail, ModifiedDate) 
                    VALUES (@name, @token, @ntn, @address, @province, @phone, @email, @modified)", conn);

                cmd.Parameters.AddWithValue("@name", company.CompanyName);
                cmd.Parameters.AddWithValue("@token", company.FBRToken);
                cmd.Parameters.AddWithValue("@ntn", company.SellerNTN ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@address", company.SellerAddress ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@province", company.SellerProvince ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@phone", company.SellerPhone ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@email", company.SellerEmail ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@modified", DateTime.Now);

                cmd.ExecuteNonQuery();
                return true;
            }, "SaveCompany");
        }

        #endregion

        #region Invoices CRUD

        public List<Invoice> GetInvoices(string companyName)
        {
            return ExecuteWithRetry(() =>
            {
                var invoices = new List<Invoice>();
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(
                    "SELECT * FROM Invoices WHERE CompanyName = @name ORDER BY CreatedDate DESC", conn);
                cmd.Parameters.AddWithValue("@name", companyName);

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
                    WHERE QuickBooksInvoiceId = @qbId AND CompanyName = @company", conn);
                cmd.Parameters.AddWithValue("@qbId", qbInvoiceId);
                cmd.Parameters.AddWithValue("@company", companyName);

                using var reader = cmd.ExecuteReader();
                if (!reader.Read()) return null;

                var invoice = MapInvoiceFromReader(reader);
                var invoiceId = invoice.Id;
                reader.Close();

                // Load invoice items
                var itemCmd = new SQLiteCommand(@"
                    SELECT * FROM InvoiceItems WHERE InvoiceId = @id", conn);
                itemCmd.Parameters.AddWithValue("@id", invoiceId);

                using var itemReader = itemCmd.ExecuteReader();
                while (itemReader.Read())
                {
                    invoice.Items.Add(new InvoiceItem
                    {
                        Id = Convert.ToInt32(itemReader["Id"]),
                        InvoiceId = Convert.ToInt32(itemReader["InvoiceId"]),
                        ItemName = itemReader["ItemName"].ToString(),
                        Quantity = Convert.ToInt32(itemReader["Quantity"]),
                        UnitPrice = Convert.ToDecimal(itemReader["UnitPrice"]),
                        TotalPrice = Convert.ToDecimal(itemReader["TotalPrice"]),
                        TaxRate = Convert.ToDecimal(itemReader["TaxRate"]),
                        SalesTaxAmount = Convert.ToDecimal(itemReader["SalesTaxAmount"]),
                        TotalValue = Convert.ToDecimal(itemReader["TotalValue"]),
                        HSCode = itemReader["HSCode"]?.ToString(),
                        UnitOfMeasure = itemReader["UnitOfMeasure"]?.ToString(),
                        RetailPrice = Convert.ToDecimal(itemReader["RetailPrice"]),
                        ExtraTax = Convert.ToDecimal(itemReader["ExtraTax"]),
                        FurtherTax = Convert.ToDecimal(itemReader["FurtherTax"]),
                        FedPayable = Convert.ToDecimal(itemReader["FedPayable"]),
                        SalesTaxWithheldAtSource = Convert.ToDecimal(itemReader["SalesTaxWithheldAtSource"]),
                        Discount = Convert.ToDecimal(itemReader["Discount"]),
                        SaleType = itemReader["SaleType"]?.ToString()
                    });
                }

                return invoice;
            }, "GetInvoiceWithDetails");
        }

        public void SaveInvoice(Invoice invoice)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    INSERT INTO Invoices 
                    (CompanyName, QuickBooksInvoiceId, InvoiceNumber, CustomerName, CustomerNTN,
                     CustomerAddress, CustomerPhone, CustomerEmail, Amount, TotalAmount, TaxAmount, 
                     DiscountAmount, InvoiceDate, PaymentMode, Status, FBR_IRN, FBR_QRCode, 
                     UploadDate, ErrorMessage, ModifiedDate) 
                    VALUES 
                    (@company, @qbId, @invNo, @customer, @ntn, @address, @phone, @email,
                     @amount, @total, @tax, @discount, @invDate, @payment, @status, 
                     @irn, @qr, @upload, @error, @modified)
                    ON CONFLICT(CompanyName, QuickBooksInvoiceId) 
                    DO UPDATE SET
                        InvoiceNumber = @invNo,
                        CustomerName = @customer,
                        CustomerNTN = @ntn,
                        CustomerAddress = @address,
                        CustomerPhone = @phone,
                        CustomerEmail = @email,
                        Amount = @amount,
                        TotalAmount = @total,
                        TaxAmount = @tax,
                        DiscountAmount = @discount,
                        InvoiceDate = @invDate,
                        PaymentMode = @payment,
                        Status = CASE 
                            WHEN Status IN ('Uploaded', 'Failed') THEN Status 
                            ELSE @status 
                        END,
                        FBR_IRN = CASE WHEN FBR_IRN IS NOT NULL THEN FBR_IRN ELSE @irn END,
                        FBR_QRCode = CASE WHEN FBR_QRCode IS NOT NULL THEN FBR_QRCode ELSE @qr END,
                        UploadDate = CASE WHEN UploadDate IS NOT NULL THEN UploadDate ELSE @upload END,
                        ErrorMessage = CASE 
                            WHEN Status = 'Failed' THEN ErrorMessage 
                            ELSE @error 
                        END,
                        ModifiedDate = @modified", conn);

                cmd.Parameters.AddWithValue("@company", invoice.CompanyName);
                cmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                cmd.Parameters.AddWithValue("@invNo", invoice.InvoiceNumber);
                cmd.Parameters.AddWithValue("@customer", invoice.CustomerName);
                cmd.Parameters.AddWithValue("@ntn", invoice.CustomerNTN ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@address", invoice.CustomerAddress ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@phone", invoice.CustomerPhone ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@email", invoice.CustomerEmail ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@amount", invoice.Amount);
                cmd.Parameters.AddWithValue("@total", invoice.TotalAmount);
                cmd.Parameters.AddWithValue("@tax", invoice.TaxAmount);
                cmd.Parameters.AddWithValue("@discount", invoice.DiscountAmount);
                cmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@payment", invoice.PaymentMode ?? "Cash");
                cmd.Parameters.AddWithValue("@status", invoice.Status);
                cmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@upload", invoice.UploadDate ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@modified", DateTime.Now);

                cmd.ExecuteNonQuery();
                return true;
            }, "SaveInvoice");
        }

        public void SaveInvoiceWithDetails(Invoice invoice)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                // SOLUTION 4: Use IMMEDIATE transaction for write operations
                using var cmd = new SQLiteCommand("BEGIN IMMEDIATE", conn);
                cmd.ExecuteNonQuery();

                try
                {
                    // Save main invoice
                    SaveInvoiceInternal(conn, invoice);

                    // Get invoice ID
                    var getIdCmd = new SQLiteCommand(
                        "SELECT Id FROM Invoices WHERE QuickBooksInvoiceId = @qbId AND CompanyName = @company",
                        conn);
                    getIdCmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                    getIdCmd.Parameters.AddWithValue("@company", invoice.CompanyName);
                    var invoiceId = Convert.ToInt32(getIdCmd.ExecuteScalar());

                    // Delete existing items
                    var deleteCmd = new SQLiteCommand(
                        "DELETE FROM InvoiceItems WHERE InvoiceId = @id", conn);
                    deleteCmd.Parameters.AddWithValue("@id", invoiceId);
                    deleteCmd.ExecuteNonQuery();

                    // Insert new items
                    if (invoice.Items != null && invoice.Items.Count > 0)
                    {
                        foreach (var item in invoice.Items)
                        {
                            SaveInvoiceItemInternal(conn, invoiceId, item);
                        }
                    }

                    new SQLiteCommand("COMMIT", conn).ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    new SQLiteCommand("ROLLBACK", conn).ExecuteNonQuery();
                    throw;
                }
            }, "SaveInvoiceWithDetails");
        }

        private void SaveInvoiceInternal(SQLiteConnection conn, Invoice invoice)
        {
            var cmd = new SQLiteCommand(@"
                INSERT OR REPLACE INTO Invoices 
                (CompanyName, QuickBooksInvoiceId, InvoiceNumber, CustomerName, CustomerNTN,
                 CustomerAddress, CustomerPhone, CustomerEmail, Amount, TotalAmount, TaxAmount, 
                 DiscountAmount, InvoiceDate, PaymentMode, Status, FBR_IRN, FBR_QRCode, 
                 UploadDate, ErrorMessage, ModifiedDate) 
                VALUES 
                (@company, @qbId, @invNo, @customer, @ntn, @address, @phone, @email,
                 @amount, @total, @tax, @discount, @invDate, @payment, @status, 
                 @irn, @qr, @upload, @error, @modified)", conn);

            cmd.Parameters.AddWithValue("@company", invoice.CompanyName);
            cmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
            cmd.Parameters.AddWithValue("@invNo", invoice.InvoiceNumber);
            cmd.Parameters.AddWithValue("@customer", invoice.CustomerName);
            cmd.Parameters.AddWithValue("@ntn", invoice.CustomerNTN ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@address", invoice.CustomerAddress ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@phone", invoice.CustomerPhone ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@email", invoice.CustomerEmail ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@amount", invoice.Amount);
            cmd.Parameters.AddWithValue("@total", invoice.TotalAmount);
            cmd.Parameters.AddWithValue("@tax", invoice.TaxAmount);
            cmd.Parameters.AddWithValue("@discount", invoice.DiscountAmount);
            cmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@payment", invoice.PaymentMode ?? "Cash");
            cmd.Parameters.AddWithValue("@status", invoice.Status);
            cmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@upload", invoice.UploadDate ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@modified", DateTime.Now);

            cmd.ExecuteNonQuery();
        }

        private void SaveInvoiceItemInternal(SQLiteConnection conn, int invoiceId, InvoiceItem item)
        {
            var itemCmd = new SQLiteCommand(@"
                INSERT INTO InvoiceItems 
                (InvoiceId, ItemName, Quantity, UnitPrice, TotalPrice, TaxRate, 
                 SalesTaxAmount, TotalValue, HSCode, UnitOfMeasure, RetailPrice, 
                 ExtraTax, FurtherTax, FedPayable, SalesTaxWithheldAtSource, 
                 Discount, SaleType)
                VALUES (@invId, @name, @qty, @unit, @total, @taxRate, 
                        @salesTax, @totalVal, @hs, @uom, @retail, 
                        @extra, @further, @fed, @withheld, @disc, @saleType)", conn);

            itemCmd.Parameters.AddWithValue("@invId", invoiceId);
            itemCmd.Parameters.AddWithValue("@name", item.ItemName);
            itemCmd.Parameters.AddWithValue("@qty", item.Quantity);
            itemCmd.Parameters.AddWithValue("@unit", item.UnitPrice);
            itemCmd.Parameters.AddWithValue("@total", item.TotalPrice);
            itemCmd.Parameters.AddWithValue("@taxRate", item.TaxRate);
            itemCmd.Parameters.AddWithValue("@salesTax", item.SalesTaxAmount);
            itemCmd.Parameters.AddWithValue("@totalVal", item.TotalValue);
            itemCmd.Parameters.AddWithValue("@hs", item.HSCode ?? (object)DBNull.Value);
            itemCmd.Parameters.AddWithValue("@uom", item.UnitOfMeasure ?? (object)DBNull.Value);
            itemCmd.Parameters.AddWithValue("@retail", item.RetailPrice);
            itemCmd.Parameters.AddWithValue("@extra", item.ExtraTax);
            itemCmd.Parameters.AddWithValue("@further", item.FurtherTax);
            itemCmd.Parameters.AddWithValue("@fed", item.FedPayable);
            itemCmd.Parameters.AddWithValue("@withheld", item.SalesTaxWithheldAtSource);
            itemCmd.Parameters.AddWithValue("@disc", item.Discount);
            itemCmd.Parameters.AddWithValue("@saleType", item.SaleType ?? (object)DBNull.Value);

            itemCmd.ExecuteNonQuery();
        }

        public void UpdateInvoiceStatus(string qbInvoiceId, string status, string irn, string error = null)
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
                Items = new List<InvoiceItem>()
            };
        }

        #endregion
    }
}