using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using System;

namespace C2B_FBR_Connect.Managers
{
    public class CompanyManager
    {
        private readonly DatabaseService _db;

        public CompanyManager(DatabaseService db)
        {
            _db = db;
        }

        public Company GetCompany(string companyName)
        {
            return _db.GetCompanyByName(companyName);
        }

        public void SaveCompany(Company company)
        {
            // Validate required fields
            if (string.IsNullOrWhiteSpace(company.CompanyName))
                throw new ArgumentException("Company name is required");

            if (string.IsNullOrWhiteSpace(company.FBRToken))
                throw new ArgumentException("FBR token is required");

            if (string.IsNullOrWhiteSpace(company.SellerNTN))
                throw new ArgumentException("Seller NTN is required for FBR compliance");

            if (string.IsNullOrWhiteSpace(company.SellerProvince))
                throw new ArgumentException("Seller Province is required for FBR compliance");

            _db.SaveCompany(company);
        }

        public bool IsCompanyConfigured(string companyName)
        {
            var company = GetCompany(companyName);
            return company != null &&
                   !string.IsNullOrEmpty(company.FBRToken) &&
                   !string.IsNullOrEmpty(company.SellerNTN) &&
                   !string.IsNullOrEmpty(company.SellerProvince);
        }
    }
}