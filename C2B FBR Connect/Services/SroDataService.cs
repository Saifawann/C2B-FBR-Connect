using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Services
{
    public class SroDataService
    {
        private readonly FBRApiService _fbrApi;
        private readonly TransactionTypeService _transactionTypeService;

        // ✅ Cache for rate lookups to avoid repeated API calls
        private readonly Dictionary<string, List<SaleTypeRate>> _rateCache = new Dictionary<string, List<SaleTypeRate>>();
        private readonly Dictionary<int, List<SroSchedule>> _scheduleCache = new Dictionary<int, List<SroSchedule>>();

        public SroDataService(FBRApiService fbrApi, TransactionTypeService transactionTypeService)
        {
            _fbrApi = fbrApi;
            _transactionTypeService = transactionTypeService;
        }

        /// <summary>
        /// Clear all caches - call when province or date changes
        /// </summary>
        public void ClearCache()
        {
            _rateCache.Clear();
            _scheduleCache.Clear();
        }

        #region Rate Extraction and Matching

        /// <summary>
        /// Extracts numeric tax rate from rate string for percentage-based rates
        /// Examples: "18%" → 18, "0.20%" → 0.2, "Exempt" → 0
        /// Returns null for non-percentage rates like "Rs.200"
        /// </summary>
        private decimal? ExtractNumericRate(string rateString)
        {
            if (string.IsNullOrWhiteSpace(rateString))
                return 0m;

            // Handle special cases that are NOT percentages
            if (rateString.Equals("Exempt", StringComparison.OrdinalIgnoreCase))
                return 0m;

            if (rateString.StartsWith("Rs.", StringComparison.OrdinalIgnoreCase) ||
                rateString.StartsWith("Rs ", StringComparison.OrdinalIgnoreCase))
            {
                return null; // Fixed rate format
            }

            if (rateString.Contains("along with", StringComparison.OrdinalIgnoreCase))
            {
                // Potassium Chlorate: "18% along with rupees 60 per kilogram" → 18
                var parts = rateString.Split(new[] { "along with" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length > 0)
                {
                    var percentPart = parts[0].Trim().Replace("%", "").Trim();
                    if (decimal.TryParse(percentPart, out decimal rate))
                        return rate;
                }
                return 18m; // Default for Potassium Chlorate
            }

            // Standard percentage format: "18%", "0.20%", "25%"
            string cleanedPercent = rateString.Replace("%", "").Trim();
            return decimal.TryParse(cleanedPercent, out decimal numericRate) ? numericRate : 0m;
        }

        /// <summary>
        /// Compares two rate strings for matching
        /// Handles both numeric rates (18%) and special formats (Rs.200, Exempt)
        /// </summary>
        private bool RatesMatch(string itemRate, string apiRateDesc, decimal? itemNumericRate, decimal apiRateValue)
        {
            // Handle special string-based rates (Rs.200, Exempt, etc.)
            if (!itemNumericRate.HasValue)
            {
                return itemRate.Equals(apiRateDesc, StringComparison.OrdinalIgnoreCase);
            }

            // Handle numeric percentage rates
            return Math.Abs(apiRateValue - itemNumericRate.Value) < 0.01m;
        }

        #endregion

        #region SRO Enrichment

        /// <summary>
        /// Enriches an invoice item with SRO Schedule and Item Serial Number data
        /// ✅ Updates item.Rate to FBR's exact RATE_DESC format
        /// ✅ Preserves default SRO values when API enrichment fails
        /// ✅ Uses caching to improve performance
        /// </summary>
        public async Task<bool> EnrichItemWithSroDataAsync(InvoiceItem item, string saleTypeDesc,
            int provinceId, DateTime invoiceDate, string authToken)
        {
            try
            {
                // Store existing default values
                string existingSchedule = item.SroScheduleNo;
                string existingSerial = item.SroItemSerialNo;
                bool hadDefaults = !string.IsNullOrEmpty(existingSchedule) && !string.IsNullOrEmpty(existingSerial);

                if (string.IsNullOrEmpty(authToken))
                {
                    return hadDefaults; // Return true if we have defaults
                }

                // Normalize sale type for API lookup
                string normalizedSaleType = ScenarioMapper.NormalizeSaleType(saleTypeDesc);

                // Step 1: Get TransactionType ID from SaleType description
                var transactionType = _transactionTypeService.GetTransactionTypes()
                    .FirstOrDefault(t => t.TransactionDesc.Equals(normalizedSaleType, StringComparison.OrdinalIgnoreCase));

                if (transactionType == null)
                {
                    if (hadDefaults)
                    {
                        item.SroScheduleNo = existingSchedule;
                        item.SroItemSerialNo = existingSerial;
                        return true;
                    }
                    return false;
                }

                int transTypeId = transactionType.TransactionTypeId;
                string dateStr = invoiceDate.ToString("dd-MMMyyyy");

                // Extract numeric rate from rate string
                decimal? numericTaxRate = ExtractNumericRate(item.Rate);

                // Step 2: Get rates for this transaction type and province (with caching)
                var rates = await GetRatesWithCacheAsync(dateStr, transTypeId, provinceId, authToken);

                if (rates == null || !rates.Any())
                {
                    return RestoreDefaultsIfAvailable(item, existingSchedule, existingSerial, hadDefaults);
                }

                // Step 3: Match the rate using smart comparison
                var matchingRate = rates.FirstOrDefault(rate => RatesMatch(item.Rate, rate.RATE_DESC, numericTaxRate, rate.RATE_VALUE));

                if (matchingRate == null)
                {
                    return RestoreDefaultsIfAvailable(item, existingSchedule, existingSerial, hadDefaults);
                }

                int rateId = matchingRate.RATE_ID;

                // Update item.Rate to FBR's exact RATE_DESC format
                item.Rate = matchingRate.RATE_DESC;

                // Step 4: Get SRO Schedule using the rate ID (with caching)
                var sroSchedules = await GetSchedulesWithCacheAsync(rateId, dateStr, provinceId, authToken);

                if (sroSchedules == null || !sroSchedules.Any())
                {
                    return RestoreDefaultsIfAvailable(item, existingSchedule, existingSerial, hadDefaults);
                }

                // Select first non-empty SRO schedule
                var selectedSchedule = sroSchedules.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s.SRO_DESC) && s.SRO_DESC != "0");

                if (selectedSchedule == null)
                {
                    return RestoreDefaultsIfAvailable(item, existingSchedule, existingSerial, hadDefaults);
                }

                string scheduleNo = selectedSchedule.SRO_DESC;

                // Step 5: Get SRO Item Serial Number using the SRO ID
                var sroItemSerialNo = await _fbrApi.FetchSroItemSerialNoAsync(selectedSchedule.SRO_ID, dateStr, authToken);

                // Only assign if we have BOTH schedule AND serial number
                if (!string.IsNullOrEmpty(scheduleNo) && !string.IsNullOrEmpty(sroItemSerialNo))
                {
                    item.SroScheduleNo = scheduleNo;
                    item.SroItemSerialNo = sroItemSerialNo;
                    return true;
                }
                else
                {
                    return RestoreDefaultsIfAvailable(item, existingSchedule, existingSerial, hadDefaults);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error enriching item with SRO data: {ex.Message}");

                // Don't clear if we have existing values
                if (!string.IsNullOrEmpty(item.SroScheduleNo) && !string.IsNullOrEmpty(item.SroItemSerialNo))
                {
                    return true;
                }

                item.SroScheduleNo = "";
                item.SroItemSerialNo = "";
                return false;
            }
        }

        /// <summary>
        /// Enriches all items in an invoice with SRO data
        /// ✅ Uses parallel processing for better performance
        /// ✅ Each item's Rate will be updated to FBR's exact format
        /// ✅ Preserves default SRO values when API enrichment fails
        /// </summary>
        public async Task<int> EnrichInvoiceItemsWithSroDataAsync(List<InvoiceItem> items,
            int provinceId, DateTime invoiceDate, string authToken)
        {
            int successCount = 0;

            Console.WriteLine($"🔍 Starting SRO enrichment for {items.Count} items...");

            // ✅ Process items that don't require SRO first (faster)
            var itemsToProcess = new List<(InvoiceItem item, bool requiresSro)>();

            foreach (var item in items)
            {
                if (string.IsNullOrEmpty(item.SaleType))
                {
                    continue;
                }

                // Extract numeric rate for SRO requirement check
                decimal? numericTaxRate = ExtractNumericRate(item.Rate);
                decimal rateForCheck = numericTaxRate ?? 0m;

                // Check if SRO is required based on BOTH sale type AND tax rate
                bool requiresSro = ScenarioMapper.RequiresSroDetails(item.SaleType, rateForCheck);

                if (!requiresSro)
                {
                    item.SroScheduleNo = "";
                    item.SroItemSerialNo = "";
                    continue;
                }

                itemsToProcess.Add((item, requiresSro));
            }

            // ✅ Process items requiring SRO (can be parallelized for large lists)
            if (itemsToProcess.Count <= 5)
            {
                // Sequential processing for small lists
                foreach (var (item, _) in itemsToProcess)
                {
                    var success = await EnrichItemWithSroDataAsync(item, item.SaleType, provinceId, invoiceDate, authToken);
                    if (success) successCount++;
                    else
                    {
                        item.SroScheduleNo = "";
                        item.SroItemSerialNo = "";
                    }
                }
            }
            else
            {
                // Parallel processing for larger lists (limit concurrency to avoid API throttling)
                var semaphore = new System.Threading.SemaphoreSlim(3); // Max 3 concurrent API calls
                var tasks = itemsToProcess.Select(async tuple =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        var success = await EnrichItemWithSroDataAsync(tuple.item, tuple.item.SaleType, provinceId, invoiceDate, authToken);
                        if (!success)
                        {
                            tuple.item.SroScheduleNo = "";
                            tuple.item.SroItemSerialNo = "";
                        }
                        return success;
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }).ToList();

                var results = await Task.WhenAll(tasks);
                successCount = results.Count(r => r);
            }

            // Final validation: Ensure no item has schedule without serial
            ValidateAndFixSroConsistency(items);

            return successCount;
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Get rates with caching to avoid repeated API calls
        /// </summary>
        private async Task<List<SaleTypeRate>> GetRatesWithCacheAsync(string dateStr, int transTypeId, int provinceId, string authToken)
        {
            string cacheKey = $"{dateStr}_{transTypeId}_{provinceId}";

            if (_rateCache.TryGetValue(cacheKey, out var cachedRates))
            {
                return cachedRates;
            }

            var rates = await _fbrApi.FetchSaleTypeRatesAsync(dateStr, transTypeId, provinceId, authToken);
            if (rates != null)
            {
                _rateCache[cacheKey] = rates;
            }

            return rates;
        }

        /// <summary>
        /// Get schedules with caching to avoid repeated API calls
        /// </summary>
        private async Task<List<SroSchedule>> GetSchedulesWithCacheAsync(int rateId, string dateStr, int provinceId, string authToken)
        {
            if (_scheduleCache.TryGetValue(rateId, out var cachedSchedules))
            {
                return cachedSchedules;
            }

            var schedules = await _fbrApi.FetchAllSroSchedulesAsync(rateId, dateStr, provinceId, authToken);
            if (schedules != null)
            {
                _scheduleCache[rateId] = schedules;
            }

            return schedules;
        }

        /// <summary>
        /// Restore default SRO values if available
        /// </summary>
        private bool RestoreDefaultsIfAvailable(InvoiceItem item, string existingSchedule, string existingSerial, bool hadDefaults)
        {
            if (hadDefaults)
            {
                item.SroScheduleNo = existingSchedule;
                item.SroItemSerialNo = existingSerial;
                return true;
            }

            item.SroScheduleNo = "";
            item.SroItemSerialNo = "";
            return false;
        }

        /// <summary>
        /// Validate and fix SRO consistency issues
        /// </summary>
        private void ValidateAndFixSroConsistency(List<InvoiceItem> items)
        {
            int errorCount = 0;

            foreach (var item in items)
            {
                bool hasSchedule = !string.IsNullOrWhiteSpace(item.SroScheduleNo);
                bool hasSerial = !string.IsNullOrWhiteSpace(item.SroItemSerialNo);

                if (hasSchedule && !hasSerial)
                {
                    errorCount++;
                    System.Diagnostics.Debug.WriteLine($"❌ ERROR: Item '{item.ItemName}' has schedule but no serial - clearing schedule");
                    item.SroScheduleNo = "";
                }
                else if (!hasSchedule && hasSerial)
                {
                    errorCount++;
                    System.Diagnostics.Debug.WriteLine($"❌ ERROR: Item '{item.ItemName}' has serial but no schedule - clearing serial");
                    item.SroItemSerialNo = "";
                }
            }

            if (errorCount > 0)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Fixed {errorCount} SRO inconsistencies to prevent upload errors");
            }
        }

        #endregion
    }
}