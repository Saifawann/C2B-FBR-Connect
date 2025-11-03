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

        public SroDataService(FBRApiService fbrApi, TransactionTypeService transactionTypeService)
        {
            _fbrApi = fbrApi;
            _transactionTypeService = transactionTypeService;
        }

        /// <summary>
        /// Enriches an invoice item with SRO Schedule and Item Serial Number data
        /// </summary>
        public async Task<bool> EnrichItemWithSroDataAsync(InvoiceItem item, string saleTypeDesc,
            int provinceId, DateTime invoiceDate, string authToken)
        {
            try
            {
                if (string.IsNullOrEmpty(authToken))
                {
                    System.Diagnostics.Debug.WriteLine("⚠️ Auth token is required for SRO data fetching");
                    return false;
                }

                // Step 1: Get TransactionType ID from SaleType description
                var transactionType = _transactionTypeService.GetTransactionTypes()
                    .FirstOrDefault(t => t.TransactionDesc.Equals(saleTypeDesc, StringComparison.OrdinalIgnoreCase));

                if (transactionType == null)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ Transaction type not found for: {saleTypeDesc}");
                    return false;
                }

                int transTypeId = transactionType.TransactionTypeId;
                string dateStr = invoiceDate.ToString("dd-MMMyyyy"); // Format: 22-Oct2025

                System.Diagnostics.Debug.WriteLine($"🔍 Fetching SRO data for item: {item.ItemName}");
                System.Diagnostics.Debug.WriteLine($"   TransTypeId: {transTypeId}, Province: {provinceId}, TaxRate: {item.TaxRate}%");

                // Step 2: Get rates for this transaction type and province
                var rates = await _fbrApi.FetchSaleTypeRatesAsync(dateStr, transTypeId, provinceId, authToken);

                if (rates == null || !rates.Any())
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ No rates found for transTypeId={transTypeId}, province={provinceId}");
                    return false;
                }

                // Step 3: Match the rate based on item's tax rate
                var matchingRate = rates.FirstOrDefault(r => Math.Abs(r.RATE_VALUE - item.TaxRate) < 0.01m);

                if (matchingRate == null)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ No matching rate found for TaxRate={item.TaxRate}%");
                    System.Diagnostics.Debug.WriteLine($"   Available rates: {string.Join(", ", rates.Select(r => $"{r.RATE_VALUE}%"))}");
                    return false;
                }

                int rateId = matchingRate.RATE_ID;
                System.Diagnostics.Debug.WriteLine($"✅ Found matching rate: {matchingRate.RATE_DESC} (ID: {rateId})");

                // Step 4: Get SRO Schedule using the rate ID
                // ✅ MODIFIED: A rate can return multiple SRO schedules, select first non-empty one
                var sroSchedules = await _fbrApi.FetchAllSroSchedulesAsync(rateId, dateStr, provinceId, authToken);

                if (sroSchedules == null || !sroSchedules.Any())
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ No SRO Schedules found for rateId={rateId}");
                    return false;
                }

                // ✅ Select first non-empty SRO schedule
                var selectedSchedule = sroSchedules
                    .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s.SRO_DESC) && s.SRO_DESC != "0");

                if (selectedSchedule == null)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ All {sroSchedules.Count} SRO Schedules are empty");
                    return false;
                }

                string scheduleNo = selectedSchedule.SRO_DESC;
                System.Diagnostics.Debug.WriteLine($"✅ Selected SRO Schedule: {scheduleNo} (from {sroSchedules.Count} schedules)");

                // Step 5: Get SRO Item Serial Number using the SRO ID
                // ✅ MODIFIED: Response may have multiple items, select first non-empty one
                var sroItemSerialNo = await _fbrApi.FetchSroItemSerialNoAsync(selectedSchedule.SRO_ID, dateStr, authToken);

                // ✅ CRITICAL: Only assign if we have BOTH schedule AND serial number
                if (!string.IsNullOrEmpty(scheduleNo) && !string.IsNullOrEmpty(sroItemSerialNo))
                {
                    item.SroScheduleNo = scheduleNo;
                    item.SroItemSerialNo = sroItemSerialNo;

                    System.Diagnostics.Debug.WriteLine($"✅ SRO Item Serial No: {item.SroItemSerialNo}");
                    System.Diagnostics.Debug.WriteLine($"✅ BOTH SRO fields successfully set");
                    return true;
                }
                else
                {
                    // ✅ CRITICAL: If serial number is missing, don't set either field
                    // This prevents FBR Error 0078
                    System.Diagnostics.Debug.WriteLine($"⚠️ SRO Item Serial Number is empty - clearing both fields to avoid Error 0078");
                    System.Diagnostics.Debug.WriteLine($"   Schedule was: {scheduleNo}");
                    System.Diagnostics.Debug.WriteLine($"   Serial was: {sroItemSerialNo ?? "(null)"}");

                    item.SroScheduleNo = "";
                    item.SroItemSerialNo = "";
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error enriching item with SRO data: {ex.Message}");

                // ✅ Clear both fields on error to prevent partial data
                item.SroScheduleNo = "";
                item.SroItemSerialNo = "";
                return false;
            }
        }

        /// <summary>
        /// Enriches all items in an invoice with SRO data
        /// </summary>
        public async Task<int> EnrichInvoiceItemsWithSroDataAsync(List<InvoiceItem> items,
    int provinceId, DateTime invoiceDate, string authToken)
        {
            int successCount = 0;

            foreach (var item in items)
            {
                if (string.IsNullOrEmpty(item.SaleType))
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ Item '{item.ItemName}' has no SaleType, skipping SRO enrichment");
                    continue;
                }

                // ✅ Check if SRO is required based on BOTH sale type AND tax rate
                bool requiresSro = ScenarioMapper.RequiresSroDetails(item.SaleType, item.TaxRate);

                if (!requiresSro)
                {
                    System.Diagnostics.Debug.WriteLine($"✅ Item '{item.ItemName}' (rate={item.TaxRate}%) does not require SRO - skipping");
                    item.SroScheduleNo = "";
                    item.SroItemSerialNo = "";
                    continue;
                }

                System.Diagnostics.Debug.WriteLine($"🔍 Item '{item.ItemName}' requires SRO (rate={item.TaxRate}%) - fetching from API...");

                var success = await EnrichItemWithSroDataAsync(item, item.SaleType, provinceId, invoiceDate, authToken);

                if (success)
                {
                    successCount++;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ SRO enrichment failed for '{item.ItemName}'");
                    System.Diagnostics.Debug.WriteLine($"   This item requires SRO data but couldn't fetch from API");

                    // Clear both fields to ensure consistency
                    item.SroScheduleNo = "";
                    item.SroItemSerialNo = "";
                }
            }

            System.Diagnostics.Debug.WriteLine($"📊 SRO Enrichment Complete: {successCount}/{items.Count} items enriched from API");

            // ✅ FINAL VALIDATION: Ensure no item has schedule without serial
            int errorCount = 0;
            foreach (var item in items)
            {
                bool hasSchedule = !string.IsNullOrWhiteSpace(item.SroScheduleNo);
                bool hasSerial = !string.IsNullOrWhiteSpace(item.SroItemSerialNo);

                if (hasSchedule && !hasSerial)
                {
                    errorCount++;
                    System.Diagnostics.Debug.WriteLine($"❌ ERROR: Item '{item.ItemName}' has schedule but no serial!");
                    System.Diagnostics.Debug.WriteLine($"   This will cause FBR Error 0078 - clearing schedule");
                    item.SroScheduleNo = "";
                }
                else if (!hasSchedule && hasSerial)
                {
                    errorCount++;
                    System.Diagnostics.Debug.WriteLine($"❌ ERROR: Item '{item.ItemName}' has serial but no schedule!");
                    System.Diagnostics.Debug.WriteLine($"   Clearing serial for consistency");
                    item.SroItemSerialNo = "";
                }
            }

            if (errorCount > 0)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Fixed {errorCount} SRO inconsistencies to prevent upload errors");
            }

            return successCount;
        }
    }
}