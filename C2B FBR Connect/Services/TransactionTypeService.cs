using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Services
{
    public class TransactionTypeService
    {
        private readonly DatabaseService _db;
        private readonly FBRApiService _fbrApi;

        public TransactionTypeService(DatabaseService db)
        {
            _db = db;
            _fbrApi = new FBRApiService();
        }

        public async Task<bool> FetchAndStoreTransactionTypesAsync(string fbrToken = null)
        {
            try
            {
                // Fetch transaction types from FBR API using FBRApiService
                var transactionTypes = await _fbrApi.FetchTransactionTypesAsync(fbrToken);

                if (transactionTypes != null && transactionTypes.Count > 0)
                {
                    // Save to database
                    _db.SaveTransactionTypes(transactionTypes);
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error fetching transaction types: {ex.Message}");
            }

            return false;
        }

        public List<TransactionType> GetTransactionTypes()
        {
            return _db.GetTransactionTypes();
        }

        public TransactionType GetTransactionTypeById(int transactionTypeId)
        {
            return _db.GetTransactionTypeById(transactionTypeId);
        }
    }
} 