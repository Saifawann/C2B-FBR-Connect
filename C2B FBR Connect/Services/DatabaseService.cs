using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace C2B_FBR_Connect.Services
{
    public class DatabaseService
    {
        private readonly string _connectionString;
        private readonly string _dbPath;

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

            _connectionString = $"Data Source={_dbPath};Version=3;";
            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            using var conn = new SQLiteConnection(_connectionString);
            conn.Open();

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
                    Amount DECIMAL NOT NULL,
                    Status TEXT DEFAULT 'Pending',
                    FBR_IRN TEXT,
                    UploadDate DATETIME,
                    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                    ErrorMessage TEXT,
                    UNIQUE(CompanyName, QuickBooksInvoiceId)
                );";

            ExecuteNonQuery(conn, createCompaniesTable);
            ExecuteNonQuery(conn, createInvoicesTable);
        }

        private void ExecuteNonQuery(SQLiteConnection conn, string sql)
        {
            using var cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        // ✅ Companies CRUD
        public Company GetCompanyByName(string companyName)
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
                    CreatedDate = Convert.ToDateTime(reader["CreatedDate"])
                };
            }
            return null;
            return null;
        }

        public void SaveCompany(Company company)
        {
            using var conn = new SQLiteConnection(_connectionString);
            conn.Open();

            var cmd = new SQLiteCommand(@"
                INSERT OR REPLACE INTO Companies 
                (CompanyName, FBRToken, SellerNTN, SellerAddress, SellerProvince, SellerPhone, SellerEmail, ModifiedDate) 
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
        }

        // ✅ Invoices CRUD
        public List<Invoice> GetInvoices(string companyName)
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
                invoices.Add(new Invoice
                {
                    Id = Convert.ToInt32(reader["Id"]),
                    CompanyName = reader["CompanyName"].ToString(),
                    QuickBooksInvoiceId = reader["QuickBooksInvoiceId"].ToString(),
                    InvoiceNumber = reader["InvoiceNumber"].ToString(),
                    CustomerName = reader["CustomerName"].ToString(),
                    CustomerNTN = reader["CustomerNTN"]?.ToString(),
                    Amount = Convert.ToDecimal(reader["Amount"]),
                    Status = reader["Status"].ToString(),
                    FBR_IRN = reader["FBR_IRN"]?.ToString(),
                    UploadDate = reader["UploadDate"] != DBNull.Value
                        ? Convert.ToDateTime(reader["UploadDate"])
                        : (DateTime?)null,
                    ErrorMessage = reader["ErrorMessage"]?.ToString()
                });
            }
            return invoices;
        }

        public void SaveInvoice(Invoice invoice)
        {
            using var conn = new SQLiteConnection(_connectionString);
            conn.Open();

            var cmd = new SQLiteCommand(@"
                INSERT OR REPLACE INTO Invoices 
                (CompanyName, QuickBooksInvoiceId, InvoiceNumber, CustomerName, CustomerNTN,
                 Amount, Status, FBR_IRN, UploadDate, ErrorMessage) 
                VALUES 
                (@company, @qbId, @invNo, @customer, @ntn, @amount, @status, 
                 @irn, @upload, @error)", conn);

            cmd.Parameters.AddWithValue("@company", invoice.CompanyName);
            cmd.Parameters.AddWithValue("@qbId", invoice.QuickBooksInvoiceId);
            cmd.Parameters.AddWithValue("@invNo", invoice.InvoiceNumber);
            cmd.Parameters.AddWithValue("@customer", invoice.CustomerName);
            cmd.Parameters.AddWithValue("@ntn", invoice.CustomerNTN ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@amount", invoice.Amount);
            cmd.Parameters.AddWithValue("@status", invoice.Status);
            cmd.Parameters.AddWithValue("@irn", invoice.FBR_IRN ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@upload", invoice.UploadDate ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@error", invoice.ErrorMessage ?? (object)DBNull.Value);

            cmd.ExecuteNonQuery();
        }

        public void UpdateInvoiceStatus(string qbInvoiceId, string status, string irn = null, string error = null)
        {
            using var conn = new SQLiteConnection(_connectionString);
            conn.Open();

            var cmd = new SQLiteCommand(@"
                UPDATE Invoices 
                SET Status = @status, FBR_IRN = @irn,
                    UploadDate = @date, ErrorMessage = @error
                WHERE QuickBooksInvoiceId = @id", conn);

            cmd.Parameters.AddWithValue("@status", status);
            cmd.Parameters.AddWithValue("@irn", irn ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@date", status == "Uploaded" ? DateTime.Now : (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@error", error ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@id", qbInvoiceId);

            cmd.ExecuteNonQuery();
        }
    }
}
