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
                        StrNo TEXT,
                        SellerAddress TEXT,
                        SellerProvince TEXT,
                        SellerPhone TEXT,
                        SellerEmail TEXT,
                        LogoImage BLOB,
                        Environment TEXT DEFAULT 'Sandbox',
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        ModifiedDate DATETIME
                    );";

                string createCustomersTable = @"
                    CREATE TABLE IF NOT EXISTS Customers (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        CompanyName TEXT NOT NULL,
                        QuickBooksCustomerId TEXT NOT NULL,
                        CustomerName TEXT NOT NULL,
                        CustomerNTN TEXT,
                        CustomerStrNo TEXT,
                        CustomerAddress TEXT,
                        CustomerPhone TEXT,
                        CustomerEmail TEXT,
                        CustomerType TEXT DEFAULT 'Unregistered',
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        ModifiedDate DATETIME,
                        UNIQUE(CompanyName, QuickBooksCustomerId)
                    );";

                string createInvoicesTable = @"
                    CREATE TABLE IF NOT EXISTS Invoices (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        CompanyName TEXT NOT NULL,
                        QuickBooksInvoiceId TEXT NOT NULL,
                        InvoiceNumber TEXT NOT NULL,
                        CustomerId INTEGER NOT NULL,
                        Amount DECIMAL NOT NULL,
                        TotalAmount DECIMAL DEFAULT 0,
                        TaxAmount DECIMAL DEFAULT 0,
                        DiscountAmount DECIMAL DEFAULT 0,
                        InvoiceDate DATETIME,
                        PaymentMode TEXT DEFAULT 'Cash',
                        Status TEXT DEFAULT 'Pending',
                        InvoiceType TEXT DEFAULT 'Invoice',
                        FBR_IRN TEXT,
                        FBR_QRCode TEXT,
                        UploadDate DATETIME,
                        ErrorMessage TEXT,
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        ModifiedDate DATETIME,
                        UNIQUE(CompanyName, QuickBooksInvoiceId),
                        FOREIGN KEY (CustomerId) REFERENCES Customers(Id)
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
                    );";

                ExecuteNonQuery(conn, createCompaniesTable);
                ExecuteNonQuery(conn, createCustomersTable);
                ExecuteNonQuery(conn, createInvoicesTable);
                ExecuteNonQuery(conn, createInvoiceItemsTable);
                ExecuteNonQuery(conn, createTransactionTypesTable);

                ExecuteNonQuerySafe(conn, "CREATE INDEX IF NOT EXISTS idx_customers_qbid ON Customers(QuickBooksCustomerId);");
                ExecuteNonQuerySafe(conn, "CREATE INDEX IF NOT EXISTS idx_customers_company ON Customers(CompanyName);");
                ExecuteNonQuerySafe(conn, "CREATE INDEX IF NOT EXISTS idx_customers_name ON Customers(CompanyName, CustomerName);");
                ExecuteNonQuerySafe(conn, "CREATE INDEX IF NOT EXISTS idx_invoices_customerid ON Invoices(CustomerId);");
                ExecuteNonQuerySafe(conn, "CREATE INDEX IF NOT EXISTS idx_invoices_qbid ON Invoices(QuickBooksInvoiceId);");
                ExecuteNonQuerySafe(conn, "CREATE INDEX IF NOT EXISTS idx_transaction_type_id ON TransactionTypes(TransactionTypeId);");
                ExecuteNonQuerySafe(conn, "CREATE INDEX IF NOT EXISTS idx_invoiceitems_invoiceid ON InvoiceItems(InvoiceId);");

                Console.WriteLine("✅ Database initialized successfully");
            }
        }

        private void ExecuteNonQuery(SQLiteConnection conn, string sql)
        {
            using var cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        private void ExecuteNonQuerySafe(SQLiteConnection conn, string sql)
        {
            try
            {
                using var cmd = new SQLiteCommand(sql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (SQLiteException ex) when (ex.Message.Contains("already exists"))
            {
                // Index already exists, ignore
            }
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
                        StrNo = reader["StrNo"] != DBNull.Value ? reader["StrNo"].ToString() : null,
                        SellerAddress = reader["SellerAddress"]?.ToString(),
                        SellerProvince = reader["SellerProvince"]?.ToString(),
                        SellerPhone = reader["SellerPhone"]?.ToString(),
                        SellerEmail = reader["SellerEmail"]?.ToString(),
                        LogoImage = reader["LogoImage"] != DBNull.Value ? (byte[])reader["LogoImage"] : null,
                        Environment = reader["Environment"] != DBNull.Value ? reader["Environment"].ToString() : "Sandbox",
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
                        StrNo = reader["StrNo"] != DBNull.Value ? reader["StrNo"].ToString() : null,
                        SellerAddress = reader["SellerAddress"]?.ToString(),
                        SellerProvince = reader["SellerProvince"]?.ToString(),
                        SellerPhone = reader["SellerPhone"]?.ToString(),
                        SellerEmail = reader["SellerEmail"]?.ToString(),
                        Environment = reader["Environment"] != DBNull.Value ? reader["Environment"].ToString() : "Sandbox",
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
                            StrNo = @StrNo,
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
                        INSERT INTO Companies (CompanyName, FBRToken, SellerNTN, StrNo, SellerAddress, SellerProvince, SellerPhone, SellerEmail, LogoImage, Environment, CreatedDate)
                        VALUES (@name, @token, @ntn, @StrNo, @address, @province, @phone, @email, @logoImage, @environment, @created)", conn);

                    cmd.Parameters.AddWithValue("@created", DateTime.Now);
                }

                cmd.Parameters.AddWithValue("@name", company.CompanyName);
                cmd.Parameters.AddWithValue("@token", company.FBRToken);
                cmd.Parameters.AddWithValue("@environment", company.Environment ?? "Sandbox");
                cmd.Parameters.AddWithValue("@ntn", company.SellerNTN ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@StrNo", company.StrNo ?? (object)DBNull.Value);
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

        #region Customers CRUD

        public Customer GetCustomerByQBId(string companyName, string qbCustomerId)
        {
            return ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    SELECT * FROM Customers 
                    WHERE CompanyName = @company 
                    AND QuickBooksCustomerId = @qbId", conn);

                cmd.Parameters.AddWithValue("@company", companyName);
                cmd.Parameters.AddWithValue("@qbId", qbCustomerId);

                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    return MapCustomerFromReader(reader);
                }

                return null;
            }, "GetCustomerByQBId");
        }

        public Customer GetCustomerById(int customerId)
        {
            return ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand("SELECT * FROM Customers WHERE Id = @id", conn);
                cmd.Parameters.AddWithValue("@id", customerId);

                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    return MapCustomerFromReader(reader);
                }

                return null;
            }, "GetCustomerById");
        }

        public List<Customer> GetCustomersByCompany(string companyName)
        {
            return ExecuteWithRetry(() =>
            {
                var customers = new List<Customer>();

                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    SELECT * FROM Customers 
                    WHERE CompanyName = @company 
                    ORDER BY CustomerName", conn);

                cmd.Parameters.AddWithValue("@company", companyName);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    customers.Add(MapCustomerFromReader(reader));
                }

                return customers;
            }, "GetCustomersByCompany");
        }

        public Customer GetOrCreateCustomer(string companyName, string qbCustomerId, CustomerData customerData)
        {
            try
            {
                var existing = GetCustomerByQBId(companyName, qbCustomerId);

                if (existing != null)
                {
                    existing.CustomerName = customerData.CustomerName ?? existing.CustomerName;
                    existing.CustomerNTN = customerData.NTN ?? existing.CustomerNTN;
                    existing.CustomerStrNo = customerData.StrNo ?? existing.CustomerStrNo;
                    existing.CustomerAddress = customerData.Address ?? existing.CustomerAddress;
                    existing.CustomerPhone = customerData.Phone ?? existing.CustomerPhone;
                    existing.CustomerEmail = customerData.Email ?? existing.CustomerEmail;
                    existing.CustomerType = customerData.CustomerType ?? existing.CustomerType;
                    existing.ModifiedDate = DateTime.Now;

                    UpdateCustomer(existing);
                    return existing;
                }
                else
                {
                    var newCustomer = new Customer
                    {
                        CompanyName = companyName,
                        QuickBooksCustomerId = qbCustomerId,
                        CustomerName = customerData.CustomerName ?? "Unknown Customer",
                        CustomerNTN = customerData.NTN,
                        CustomerStrNo = customerData.StrNo,
                        CustomerAddress = customerData.Address,
                        CustomerPhone = customerData.Phone,
                        CustomerEmail = customerData.Email,
                        CustomerType = customerData.CustomerType ?? "Unregistered",
                        CreatedDate = DateTime.Now
                    };

                    InsertCustomer(newCustomer);
                    return newCustomer;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ GetOrCreateCustomer failed: {ex.Message}");
                throw new Exception($"Failed to get or create customer: {ex.Message}", ex);
            }
        }

        public void InsertCustomer(Customer customer)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    INSERT INTO Customers (
                        CompanyName, QuickBooksCustomerId, CustomerName,
                        CustomerNTN, CustomerStrNo, CustomerAddress, CustomerPhone, CustomerEmail,
                        CustomerType, CreatedDate
                    ) VALUES (
                        @company, @qbId, @name,
                        @ntn, @strNo, @address, @phone, @email,
                        @type, @created
                    );
                    SELECT last_insert_rowid();", conn);

                cmd.Parameters.AddWithValue("@company", customer.CompanyName);
                cmd.Parameters.AddWithValue("@qbId", customer.QuickBooksCustomerId);
                cmd.Parameters.AddWithValue("@name", customer.CustomerName);
                cmd.Parameters.AddWithValue("@ntn", customer.CustomerNTN ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@strNo", customer.CustomerStrNo ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@address", customer.CustomerAddress ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@phone", customer.CustomerPhone ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@email", customer.CustomerEmail ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@type", customer.CustomerType ?? "Unregistered");
                cmd.Parameters.AddWithValue("@created", DateTime.Now);

                customer.Id = Convert.ToInt32(cmd.ExecuteScalar());

                Console.WriteLine($"✅ Created customer: {customer.CustomerName} (ID: {customer.Id})");
                return true;
            }, "InsertCustomer");
        }

        public void UpdateCustomer(Customer customer)
        {
            ExecuteWithRetry(() =>
            {
                using var conn = new SQLiteConnection(_connectionString);
                conn.Open();

                var cmd = new SQLiteCommand(@"
                    UPDATE Customers SET
                        CustomerName = @name,
                        CustomerNTN = @ntn,
                        CustomerStrNo = @strNo,
                        CustomerAddress = @address,
                        CustomerPhone = @phone,
                        CustomerEmail = @email,
                        CustomerType = @type,
                        ModifiedDate = @modified
                    WHERE Id = @id", conn);

                cmd.Parameters.AddWithValue("@id", customer.Id);
                cmd.Parameters.AddWithValue("@name", customer.CustomerName);
                cmd.Parameters.AddWithValue("@ntn", customer.CustomerNTN ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@strNo", customer.CustomerStrNo ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@address", customer.CustomerAddress ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@phone", customer.CustomerPhone ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@email", customer.CustomerEmail ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@type", customer.CustomerType ?? "Unregistered");
                cmd.Parameters.AddWithValue("@modified", DateTime.Now);

                cmd.ExecuteNonQuery();

                Console.WriteLine($"✅ Updated customer: {customer.CustomerName} (ID: {customer.Id})");
                return true;
            }, "UpdateCustomer");
        }

        private Customer MapCustomerFromReader(SQLiteDataReader reader)
        {
            return new Customer
            {
                Id = Convert.ToInt32(reader["Id"]),
                CompanyName = reader["CompanyName"].ToString(),
                QuickBooksCustomerId = reader["QuickBooksCustomerId"].ToString(),
                CustomerName = reader["CustomerName"].ToString(),
                CustomerNTN = reader["CustomerNTN"] != DBNull.Value ? reader["CustomerNTN"].ToString() : null,
                CustomerStrNo = reader["CustomerStrNo"] != DBNull.Value ? reader["CustomerStrNo"].ToString() : null,
                CustomerAddress = reader["CustomerAddress"] != DBNull.Value ? reader["CustomerAddress"].ToString() : null,
                CustomerPhone = reader["CustomerPhone"] != DBNull.Value ? reader["CustomerPhone"].ToString() : null,
                CustomerEmail = reader["CustomerEmail"] != DBNull.Value ? reader["CustomerEmail"].ToString() : null,
                CustomerType = reader["CustomerType"] != DBNull.Value ? reader["CustomerType"].ToString() : "Unregistered",
                CreatedDate = Convert.ToDateTime(reader["CreatedDate"]),
                ModifiedDate = reader["ModifiedDate"] != DBNull.Value ? Convert.ToDateTime(reader["ModifiedDate"]) : (DateTime?)null
            };
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

                SQLiteCommand cmd;

                if (existingId != null)
                {
                    cmd = new SQLiteCommand(@"
                        UPDATE Invoices SET
                            InvoiceNumber = @invNum,
                            CustomerId = @customerId,
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
                        WHERE Id = @id", conn);

                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(existingId));
                    cmd.Parameters.AddWithValue("@invNum", invoice.InvoiceNumber);
                    cmd.Parameters.AddWithValue("@customerId", invoice.CustomerId);
                    cmd.Parameters.AddWithValue("@amount", invoice.Amount);
                    cmd.Parameters.AddWithValue("@totalAmount", invoice.TotalAmount);
                    cmd.Parameters.AddWithValue("@taxAmount", invoice.TaxAmount);
                    cmd.Parameters.AddWithValue("@discountAmount", invoice.DiscountAmount);
                    cmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@paymentMode", invoice.PaymentMode ?? "Cash");
                    cmd.Parameters.AddWithValue("@status", invoice.Status ?? "Pending");
                    cmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@uploadDate", invoice.UploadDate ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@invoiceType", invoice.InvoiceType ?? "Invoice");
                    cmd.Parameters.AddWithValue("@modified", DateTime.Now);

                    cmd.ExecuteNonQuery();
                    invoice.Id = Convert.ToInt32(existingId);
                }
                else
                {
                    cmd = new SQLiteCommand(@"
                        INSERT INTO Invoices (
                            CompanyName, QuickBooksInvoiceId, InvoiceNumber, CustomerId,
                            Amount, TotalAmount, TaxAmount, DiscountAmount,
                            InvoiceDate, PaymentMode, Status, InvoiceType,
                            FBR_IRN, FBR_QRCode, UploadDate, ErrorMessage, CreatedDate
                        ) VALUES (
                            @company, @qbId, @invNum, @customerId,
                            @amount, @totalAmount, @taxAmount, @discountAmount,
                            @invDate, @paymentMode, @status, @invoiceType,
                            @irn, @qr, @uploadDate, @error, @created
                        );
                        SELECT last_insert_rowid();", conn);

                    cmd.Parameters.AddWithValue("@company", invoice.CompanyName);
                    cmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                    cmd.Parameters.AddWithValue("@invNum", invoice.InvoiceNumber);
                    cmd.Parameters.AddWithValue("@customerId", invoice.CustomerId);
                    cmd.Parameters.AddWithValue("@amount", invoice.Amount);
                    cmd.Parameters.AddWithValue("@totalAmount", invoice.TotalAmount);
                    cmd.Parameters.AddWithValue("@taxAmount", invoice.TaxAmount);
                    cmd.Parameters.AddWithValue("@discountAmount", invoice.DiscountAmount);
                    cmd.Parameters.AddWithValue("@invDate", invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate : (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@paymentMode", invoice.PaymentMode ?? "Cash");
                    cmd.Parameters.AddWithValue("@status", invoice.Status ?? "Pending");
                    cmd.Parameters.AddWithValue("@invoiceType", invoice.InvoiceType ?? "Invoice");
                    cmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@qr", invoice.FBR_QRCode ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@uploadDate", invoice.UploadDate ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@created", invoice.CreatedDate);

                    invoice.Id = Convert.ToInt32(cmd.ExecuteScalar());
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
                                CustomerId = @customerId,
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
                        updateCmd.Parameters.AddWithValue("@customerId", invoice.CustomerId);
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
                                CompanyName, QuickBooksInvoiceId, InvoiceNumber, CustomerId,
                                Amount, TotalAmount, TaxAmount, DiscountAmount,
                                InvoiceDate, PaymentMode, Status,
                                FBR_IRN, FBR_QRCode, UploadDate, ErrorMessage, InvoiceType, CreatedDate
                            ) VALUES (
                                @company, @qbId, @invNum, @customerId,
                                @amount, @totalAmount, @taxAmount, @discountAmount,
                                @invDate, @paymentMode, @status,
                                @irn, @qr, @uploadDate, @error, @invoiceType, @created
                            );
                            SELECT last_insert_rowid();", conn, transaction);

                        insertCmd.Parameters.AddWithValue("@company", invoice.CompanyName);
                        insertCmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
                        insertCmd.Parameters.AddWithValue("@invNum", invoice.InvoiceNumber);
                        insertCmd.Parameters.AddWithValue("@customerId", invoice.CustomerId);
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

                    // STEP 1: Get all invoices with customer info (NO items in this query)
                    string invoiceQuery = @"
                SELECT 
                    i.Id, i.CompanyName, i.QuickBooksInvoiceId, i.InvoiceNumber, i.CustomerId,
                    i.Amount, i.TotalAmount, i.TaxAmount, i.DiscountAmount, 
                    i.InvoiceDate, i.PaymentMode, i.Status, i.InvoiceType,
                    i.FBR_IRN, i.FBR_QRCode, i.UploadDate, i.ErrorMessage,
                    i.CreatedDate, i.ModifiedDate,
                    c.Id as Customer_Id,
                    c.QuickBooksCustomerId as Customer_QBId,
                    c.CustomerName,
                    c.CustomerNTN,
                    c.CustomerStrNo,
                    c.CustomerAddress,
                    c.CustomerPhone,
                    c.CustomerEmail,
                    c.CustomerType
                FROM Invoices i
                LEFT JOIN Customers c ON i.CustomerId = c.Id
                WHERE i.CompanyName = @CompanyName
                ORDER BY i.CreatedDate DESC";

                    using (var cmd = new SQLiteCommand(invoiceQuery, connection))
                    {
                        cmd.Parameters.AddWithValue("@CompanyName", companyName);

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int invoiceId = Convert.ToInt32(reader["Id"]);

                                var invoice = new Invoice
                                {
                                    Id = invoiceId,
                                    CompanyName = reader["CompanyName"]?.ToString(),
                                    QuickBooksInvoiceId = reader["QuickBooksInvoiceId"]?.ToString(),
                                    InvoiceNumber = reader["InvoiceNumber"]?.ToString(),
                                    CustomerId = Convert.ToInt32(reader["CustomerId"]),
                                    Amount = reader["Amount"] != DBNull.Value ? Convert.ToDecimal(reader["Amount"]) : 0,
                                    TotalAmount = reader["TotalAmount"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAmount"]) : 0,
                                    TaxAmount = reader["TaxAmount"] != DBNull.Value ? Convert.ToDecimal(reader["TaxAmount"]) : 0,
                                    DiscountAmount = reader["DiscountAmount"] != DBNull.Value ? Convert.ToDecimal(reader["DiscountAmount"]) : 0,
                                    InvoiceDate = reader["InvoiceDate"] != DBNull.Value ? Convert.ToDateTime(reader["InvoiceDate"]) : DateTime.MinValue,
                                    PaymentMode = reader["PaymentMode"]?.ToString() ?? "Cash",
                                    Status = reader["Status"]?.ToString() ?? "Pending",
                                    InvoiceType = reader["InvoiceType"]?.ToString() ?? "Invoice",
                                    FBR_IRN = reader["FBR_IRN"] != DBNull.Value ? reader["FBR_IRN"].ToString() : null,
                                    FBR_QRCode = reader["FBR_QRCode"] != DBNull.Value ? reader["FBR_QRCode"].ToString() : null,
                                    UploadDate = reader["UploadDate"] != DBNull.Value ? Convert.ToDateTime(reader["UploadDate"]) : (DateTime?)null,
                                    ErrorMessage = reader["ErrorMessage"] != DBNull.Value ? reader["ErrorMessage"].ToString() : null,
                                    CreatedDate = reader["CreatedDate"] != DBNull.Value ? Convert.ToDateTime(reader["CreatedDate"]) : DateTime.Now,
                                    ModifiedDate = reader["ModifiedDate"] != DBNull.Value ? Convert.ToDateTime(reader["ModifiedDate"]) : (DateTime?)null,
                                    Items = new List<InvoiceItem>()
                                };

                                // Populate Customer navigation property
                                if (reader["Customer_Id"] != DBNull.Value)
                                {
                                    invoice.Customer = new Customer
                                    {
                                        Id = Convert.ToInt32(reader["Customer_Id"]),
                                        QuickBooksCustomerId = reader["Customer_QBId"]?.ToString(),
                                        CustomerName = reader["CustomerName"]?.ToString(),
                                        CustomerNTN = reader["CustomerNTN"] != DBNull.Value ? reader["CustomerNTN"].ToString() : null,
                                        CustomerStrNo = reader["CustomerStrNo"] != DBNull.Value ? reader["CustomerStrNo"].ToString() : null,
                                        CustomerAddress = reader["CustomerAddress"] != DBNull.Value ? reader["CustomerAddress"].ToString() : null,
                                        CustomerPhone = reader["CustomerPhone"] != DBNull.Value ? reader["CustomerPhone"].ToString() : null,
                                        CustomerEmail = reader["CustomerEmail"] != DBNull.Value ? reader["CustomerEmail"].ToString() : null,
                                        CustomerType = reader["CustomerType"]?.ToString() ?? "Unregistered"
                                    };
                                }

                                invoiceDict[invoiceId] = invoice;
                            }
                        }
                    }

                    // STEP 2: Get all items separately (using standard column names)
                    if (invoiceDict.Count > 0)
                    {
                        string itemsQuery = @"
                    SELECT * FROM InvoiceItems 
                    WHERE InvoiceId IN (SELECT Id FROM Invoices WHERE CompanyName = @CompanyName)
                    ORDER BY InvoiceId, Id";

                        using (var cmd = new SQLiteCommand(itemsQuery, connection))
                        {
                            cmd.Parameters.AddWithValue("@CompanyName", companyName);

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    int invoiceId = Convert.ToInt32(reader["InvoiceId"]);

                                    if (invoiceDict.ContainsKey(invoiceId))
                                    {
                                        var item = new InvoiceItem
                                        {
                                            Id = Convert.ToInt32(reader["Id"]),
                                            InvoiceId = invoiceId,
                                            ItemName = reader["ItemName"]?.ToString(),
                                            ItemDescription = reader["ItemDescription"] != DBNull.Value ? reader["ItemDescription"].ToString() : null,
                                            Quantity = Convert.ToInt32(reader["Quantity"]),
                                            UnitPrice = Convert.ToDecimal(reader["UnitPrice"]),
                                            TotalPrice = Convert.ToDecimal(reader["TotalPrice"]),
                                            NetAmount = reader["NetAmount"] != DBNull.Value ? Convert.ToDecimal(reader["NetAmount"]) : 0,
                                            TaxRate = reader["TaxRate"] != DBNull.Value ? Convert.ToDecimal(reader["TaxRate"]) : 0,
                                            SalesTaxAmount = reader["SalesTaxAmount"] != DBNull.Value ? Convert.ToDecimal(reader["SalesTaxAmount"]) : 0,
                                            TotalValue = reader["TotalValue"] != DBNull.Value ? Convert.ToDecimal(reader["TotalValue"]) : 0,
                                            HSCode = reader["HSCode"]?.ToString(),
                                            UnitOfMeasure = reader["UnitOfMeasure"]?.ToString(),
                                            RetailPrice = reader["RetailPrice"] != DBNull.Value ? Convert.ToDecimal(reader["RetailPrice"]) : 0,
                                            ExtraTax = reader["ExtraTax"] != DBNull.Value ? Convert.ToDecimal(reader["ExtraTax"]) : 0,
                                            FurtherTax = reader["FurtherTax"] != DBNull.Value ? Convert.ToDecimal(reader["FurtherTax"]) : 0,
                                            FedPayable = reader["FedPayable"] != DBNull.Value ? Convert.ToDecimal(reader["FedPayable"]) : 0,
                                            SalesTaxWithheldAtSource = reader["SalesTaxWithheldAtSource"] != DBNull.Value ? Convert.ToDecimal(reader["SalesTaxWithheldAtSource"]) : 0,
                                            Discount = reader["Discount"] != DBNull.Value ? Convert.ToDecimal(reader["Discount"]) : 0,
                                            SaleType = reader["SaleType"]?.ToString(),
                                            SroScheduleNo = reader["SroScheduleNo"]?.ToString(),
                                            SroItemSerialNo = reader["SroItemSerialNo"]?.ToString()
                                        };

                                        invoiceDict[invoiceId].Items.Add(item);
                                    }
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
                        CustomerId = @customerId,
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
                    updateCmd.Parameters.AddWithValue("@customerId", invoice.CustomerId);
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
                        CompanyName, QuickBooksInvoiceId, InvoiceNumber, CustomerId,
                        Amount, TotalAmount, TaxAmount, DiscountAmount,
                        InvoiceDate, PaymentMode, Status,
                        FBR_IRN, FBR_QRCode, UploadDate, ErrorMessage, InvoiceType, CreatedDate
                    ) VALUES (
                        @company, @qbId, @invNum, @customerId,
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
                    insertCmd.Parameters.AddWithValue("@customerId", invoice.CustomerId);
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
            SELECT 
                i.*,
                c.Id as Customer_Id,
                c.QuickBooksCustomerId as Customer_QBId,
                c.CustomerName,
                c.CustomerNTN,
                c.CustomerStrNo,
                c.CustomerAddress,
                c.CustomerPhone,
                c.CustomerEmail,
                c.CustomerType
            FROM Invoices i
            LEFT JOIN Customers c ON i.CustomerId = c.Id
            WHERE i.CompanyName = @company 
            ORDER BY i.CreatedDate DESC", conn);

                cmd.Parameters.AddWithValue("@company", companyName);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var invoice = MapInvoiceFromReader(reader);

                    // Populate Customer navigation property
                    if (reader["Customer_Id"] != DBNull.Value)
                    {
                        invoice.Customer = new Customer
                        {
                            Id = Convert.ToInt32(reader["Customer_Id"]),
                            QuickBooksCustomerId = reader["Customer_QBId"]?.ToString(),
                            CustomerName = reader["CustomerName"]?.ToString(),
                            CustomerNTN = reader["CustomerNTN"] != DBNull.Value ? reader["CustomerNTN"].ToString() : null,
                            CustomerStrNo = reader["CustomerStrNo"] != DBNull.Value ? reader["CustomerStrNo"].ToString() : null,
                            CustomerAddress = reader["CustomerAddress"] != DBNull.Value ? reader["CustomerAddress"].ToString() : null,
                            CustomerPhone = reader["CustomerPhone"] != DBNull.Value ? reader["CustomerPhone"].ToString() : null,
                            CustomerEmail = reader["CustomerEmail"] != DBNull.Value ? reader["CustomerEmail"].ToString() : null,
                            CustomerType = reader["CustomerType"]?.ToString() ?? "Unregistered"
                        };
                    }

                    invoices.Add(invoice);
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

                // Load Invoice
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
                        invoice = MapInvoiceFromReader(reader);
                }

                if (invoice == null)
                    return null;

                // Load Customer
                if (invoice.CustomerId > 0)
                {
                    var customerCmd = new SQLiteCommand(
                        "SELECT * FROM Customers WHERE Id = @id", conn);

                    customerCmd.Parameters.AddWithValue("@id", invoice.CustomerId);

                    using var customerReader = customerCmd.ExecuteReader();
                    if (customerReader.Read())
                    {
                        invoice.Customer = MapCustomerFromReader(customerReader);
                    }
                }

                // Load Invoice Items
                var itemsCmd = new SQLiteCommand(
                    "SELECT * FROM InvoiceItems WHERE InvoiceId = @invoiceId", conn);

                itemsCmd.Parameters.AddWithValue("@invoiceId", invoice.Id);

                invoice.Items = new List<InvoiceItem>();

                using (var itemsReader = itemsCmd.ExecuteReader())
                {
                    while (itemsReader.Read())
                        invoice.Items.Add(MapInvoiceItemFromReader(itemsReader));
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
                CustomerId = Convert.ToInt32(reader["CustomerId"]),
                Amount = Convert.ToDecimal(reader["Amount"]),
                TotalAmount = reader["TotalAmount"] != DBNull.Value ? Convert.ToDecimal(reader["TotalAmount"]) : 0,
                TaxAmount = reader["TaxAmount"] != DBNull.Value ? Convert.ToDecimal(reader["TaxAmount"]) : 0,
                DiscountAmount = reader["DiscountAmount"] != DBNull.Value ? Convert.ToDecimal(reader["DiscountAmount"]) : 0,
                InvoiceDate = reader["InvoiceDate"] != DBNull.Value ? Convert.ToDateTime(reader["InvoiceDate"]) : DateTime.MinValue,
                PaymentMode = reader["PaymentMode"]?.ToString() ?? "Cash",
                Status = reader["Status"]?.ToString() ?? "Pending",
                InvoiceType = reader["InvoiceType"]?.ToString() ?? "Invoice",
                FBR_IRN = reader["FBR_IRN"] != DBNull.Value ? reader["FBR_IRN"].ToString() : null,
                FBR_QRCode = reader["FBR_QRCode"] != DBNull.Value ? reader["FBR_QRCode"].ToString() : null,
                UploadDate = reader["UploadDate"] != DBNull.Value ? Convert.ToDateTime(reader["UploadDate"]) : (DateTime?)null,
                ErrorMessage = reader["ErrorMessage"] != DBNull.Value ? reader["ErrorMessage"].ToString() : null,
                CreatedDate = Convert.ToDateTime(reader["CreatedDate"]),
                ModifiedDate = reader["ModifiedDate"] != DBNull.Value ? Convert.ToDateTime(reader["ModifiedDate"]) : (DateTime?)null,
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
                ItemDescription = reader["ItemDescription"]?.ToString(),
                Quantity = Convert.ToInt32(reader["Quantity"]),
                UnitPrice = Convert.ToDecimal(reader["UnitPrice"]),
                TotalPrice = Convert.ToDecimal(reader["TotalPrice"]),
                NetAmount = reader["NetAmount"] != DBNull.Value ? Convert.ToDecimal(reader["NetAmount"]) : 0,
                TaxRate = reader["TaxRate"] != DBNull.Value ? Convert.ToDecimal(reader["TaxRate"]) : 0,
                SalesTaxAmount = reader["SalesTaxAmount"] != DBNull.Value ? Convert.ToDecimal(reader["SalesTaxAmount"]) : 0,
                TotalValue = reader["TotalValue"] != DBNull.Value ? Convert.ToDecimal(reader["TotalValue"]) : 0,
                HSCode = reader["HSCode"]?.ToString(),
                UnitOfMeasure = reader["UnitOfMeasure"]?.ToString(),
                RetailPrice = reader["RetailPrice"] != DBNull.Value ? Convert.ToDecimal(reader["RetailPrice"]) : 0,
                ExtraTax = reader["ExtraTax"] != DBNull.Value ? Convert.ToDecimal(reader["ExtraTax"]) : 0,
                FurtherTax = reader["FurtherTax"] != DBNull.Value ? Convert.ToDecimal(reader["FurtherTax"]) : 0,
                FedPayable = reader["FedPayable"] != DBNull.Value ? Convert.ToDecimal(reader["FedPayable"]) : 0,
                SalesTaxWithheldAtSource = reader["SalesTaxWithheldAtSource"] != DBNull.Value ? Convert.ToDecimal(reader["SalesTaxWithheldAtSource"]) : 0,
                Discount = reader["Discount"] != DBNull.Value ? Convert.ToDecimal(reader["Discount"]) : 0,
                SaleType = reader["SaleType"]?.ToString(),
                SroScheduleNo = reader["SroScheduleNo"]?.ToString(),
                SroItemSerialNo = reader["SroItemSerialNo"]?.ToString()
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
    }
}