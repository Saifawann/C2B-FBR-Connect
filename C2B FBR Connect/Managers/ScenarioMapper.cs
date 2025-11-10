using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace C2B_FBR_Connect.Services
{
    /// <summary>
    /// Maps FBR sale types to scenarios and handles SRO requirements
    /// Based on official FBR e-invoice scenario testing guide
    /// </summary>
    public static class ScenarioMapper
    {
        // ✅ Static readonly for better performance
        private static readonly Dictionary<string, ScenarioDefinition> ScenarioMap = new Dictionary<string, ScenarioDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["Goods at Standard Rate (default)"] = new ScenarioDefinition("SN001", "Goods at standard rate to registered buyers", false, true),
            ["Goods at standard rate (default)"] = new ScenarioDefinition("SN001", "Goods at standard rate to registered buyers", false, true),
            ["Steel Melting and re-rolling"] = new ScenarioDefinition("SN003", "Sale of Steel (Melted and Re-Rolled)", false),
            ["Steel melting and re-rolling"] = new ScenarioDefinition("SN003", "Sale of Steel (Melted and Re-Rolled)", false),
            ["Ship breaking"] = new ScenarioDefinition("SN004", "Sale by Ship Breakers", false),
            ["Goods at Reduced Rate"] = new ScenarioDefinition("SN005", "Reduced rate sale", true),
            ["Exempt goods"] = new ScenarioDefinition("SN006", "Exempt goods sale", true),
            ["Goods at zero-rate"] = new ScenarioDefinition("SN007", "Zero rated sale", true),
            ["3rd Schedule Goods"] = new ScenarioDefinition("SN008", "Sale of 3rd schedule goods", false),
            ["Cotton Ginners"] = new ScenarioDefinition("SN009", "Cotton Spinners purchase from Cotton Ginners", false),
            ["Cotton ginners"] = new ScenarioDefinition("SN009", "Cotton Spinners purchase from Cotton Ginners", false),
            ["Telecommunication services"] = new ScenarioDefinition("SN010", "Telecom services rendered or provided", false),
            ["Toll Manufacturing"] = new ScenarioDefinition("SN011", "Toll Manufacturing sale by Steel sector", false),
            ["Petroleum Products"] = new ScenarioDefinition("SN012", "Sale of Petroleum products", true),
            ["Electricity Supply to Retailers"] = new ScenarioDefinition("SN013", "Electricity Supply to Retailers", true),
            ["Gas to CNG stations"] = new ScenarioDefinition("SN014", "Sale of Gas to CNG stations", false),
            ["Mobile Phones"] = new ScenarioDefinition("SN015", "Sale of mobile phones", true),
            ["Processing/ Conversion of Goods"] = new ScenarioDefinition("SN016", "Processing / Conversion of Goods", false),
            ["Processing/Conversion of Goods"] = new ScenarioDefinition("SN016", "Processing / Conversion of Goods", false),
            ["Goods (FED in ST Mode)"] = new ScenarioDefinition("SN017", "Sale of Goods where FED is charged in ST mode", false),
            ["Services (FED in ST Mode)"] = new ScenarioDefinition("SN018", "Services rendered or provided where FED is charged in ST mode", false),
            ["Services"] = new ScenarioDefinition("SN019", "Services rendered or provided", true),
            ["Electric Vehicle"] = new ScenarioDefinition("SN020", "Sale of Electric Vehicles", true),
            ["Cement /Concrete Block"] = new ScenarioDefinition("SN021", "Sale of Cement /Concrete Block", false),
            ["Potassium Chlorate"] = new ScenarioDefinition("SN022", "Sale of Potassium Chlorate", true),
            ["CNG Sales"] = new ScenarioDefinition("SN023", "Sale of CNG", true),
            ["Goods as per SRO.297(|)/2023"] = new ScenarioDefinition("SN024", "Goods sold that are listed in SRO 297(1)/2023", true),
            ["Goods as per SRO.297(I)/2023"] = new ScenarioDefinition("SN024", "Goods sold that are listed in SRO 297(1)/2023", true),
            ["Non-Adjustable Supplies"] = new ScenarioDefinition("SN025", "Drugs sold at fixed ST rate under serial 81 of Eighth Schedule Table 1", true),
            ["Goods at standard rate to end consumer"] = new ScenarioDefinition("SN026", "Sale to End Consumer by retailers", false),
            ["3rd Schedule Goods to end consumer"] = new ScenarioDefinition("SN027", "Sale to End Consumer by retailers - 3rd Schedule", false),
            ["Goods at Reduced Rate to end consumer"] = new ScenarioDefinition("SN028", "Sale to End Consumer by retailers - Reduced Rate", true)
        };

        // ✅ Normalization map as static readonly
        private static readonly Dictionary<string, string> NormalizationMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["Electricity Supply to Retailer"] = "Electricity Supply to Retailers",
            ["Electricity supply to retailer"] = "Electricity Supply to Retailers",
            ["Electricity supply to retailers"] = "Electricity Supply to Retailers",
            ["Cotton Ginners"] = "Cotton ginners",
            ["Goods at Standard Rate"] = "Goods at Standard Rate (default)",
            ["Goods at standard rate"] = "Goods at Standard Rate (default)",
            ["Standard Rate"] = "Goods at Standard Rate (default)",
            ["3rd Schedule Good"] = "3rd Schedule Goods",
            ["Third Schedule Goods"] = "3rd Schedule Goods",
            ["Third Schedule Good"] = "3rd Schedule Goods",
            ["Processing/Conversion"] = "Processing/Conversion of Goods",
            ["Exempt Good"] = "Exempt goods",
            ["exempt good"] = "Exempt goods",
            ["Goods at zero rate"] = "Goods at zero-rate",
            ["Zero Rate Goods"] = "Goods at zero-rate",
            ["zero-rated goods"] = "Goods at zero-rate",
            ["Goods as per SRO 297(I)/2023"] = "Goods as per SRO.297(|)/2023",
            ["Goods as per SRO 297(|)/2023"] = "Goods as per SRO.297(|)/2023",
            ["Goods as per SRO.297(I)/2023"] = "Goods as per SRO.297(|)/2023",
            ["Service"] = "Services",
            ["service"] = "Services",
            ["Cement/Concrete Block"] = "Cement /Concrete Block",
            ["Cement / Concrete Block"] = "Cement /Concrete Block",
            ["Cement"] = "Cement /Concrete Block"
        };

        #region Public Methods

        /// <summary>
        /// Determines scenario ID based on sale type and buyer registration
        /// </summary>
        public static string DetermineScenarioId(string saleType, string buyerRegistrationType = "Registered")
        {
            if (string.IsNullOrEmpty(saleType))
                return "SN001";

            string normalizedSaleType = NormalizeSaleType(saleType);

            if (ScenarioMap.TryGetValue(normalizedSaleType, out var scenario))
            {
                return scenario.BuyerDependant && buyerRegistrationType == "Unregistered" ? "SN002" : scenario.ScenarioId;
            }

            // Check for partial match
            var matchedScenario = ScenarioMap.FirstOrDefault(kvp =>
                normalizedSaleType.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase) ||
                kvp.Key.Contains(normalizedSaleType, StringComparison.OrdinalIgnoreCase));

            if (!matchedScenario.Equals(default(KeyValuePair<string, ScenarioDefinition>)))
            {
                return matchedScenario.Value.BuyerDependant && buyerRegistrationType == "Unregistered"
                    ? "SN002"
                    : matchedScenario.Value.ScenarioId;
            }

            return "SN001";
        }

        /// <summary>
        /// Check if scenario requires SRO details
        /// </summary>
        public static bool RequiresSroDetails(string saleType, decimal taxRate)
        {
            // CRITICAL: If rate is not 18%, SRO is ALWAYS required (FBR Rule)
            if (taxRate != 18m && taxRate != 0m) // Exempt/Zero-rated are exceptions
            {
                return true;
            }

            if (string.IsNullOrEmpty(saleType))
                return false;

            string normalizedSaleType = NormalizeSaleType(saleType);

            if (ScenarioMap.TryGetValue(normalizedSaleType, out var scenario))
            {
                return scenario.RequiresSRO;
            }

            // Check for partial match
            var matchedScenario = ScenarioMap.FirstOrDefault(kvp =>
                normalizedSaleType.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase) ||
                kvp.Key.Contains(normalizedSaleType, StringComparison.OrdinalIgnoreCase));

            return !matchedScenario.Equals(default(KeyValuePair<string, ScenarioDefinition>)) && matchedScenario.Value.RequiresSRO;
        }

        /// <summary>
        /// Get scenario definition for display/validation
        /// </summary>
        public static ScenarioDefinition GetScenarioDefinition(string saleType)
        {
            if (string.IsNullOrEmpty(saleType))
                return new ScenarioDefinition("SN001", "Default", false);

            string trimmedSaleType = saleType.Trim();

            if (ScenarioMap.TryGetValue(trimmedSaleType, out var scenario))
            {
                return scenario;
            }

            // Check for partial match
            var matchedScenario = ScenarioMap.FirstOrDefault(kvp =>
                trimmedSaleType.Contains(kvp.Key.Trim(), StringComparison.OrdinalIgnoreCase) ||
                kvp.Key.Trim().Contains(trimmedSaleType, StringComparison.OrdinalIgnoreCase));

            return !matchedScenario.Equals(default(KeyValuePair<string, ScenarioDefinition>))
                ? matchedScenario.Value
                : new ScenarioDefinition("SN001", "Default", false);
        }

        /// <summary>
        /// Get all sale types for dropdown/selection
        /// </summary>
        public static List<string> GetAllSaleTypes()
        {
            return ScenarioMap.Keys.Distinct().OrderBy(k => k).ToList();
        }

        /// <summary>
        /// Validate SRO data completeness for InvoiceItem
        /// </summary>
        public static ValidationResult ValidateSroData(InvoiceItem item)
        {
            var result = new ValidationResult { IsValid = true };

            if (!RequiresSroDetails(item.SaleType, item.TaxRate))
                return result;

            if (string.IsNullOrWhiteSpace(item.SroScheduleNo))
            {
                result.IsValid = false;
                result.Errors.Add($"⚠️ SRO Schedule Number is required for sale type: {item.SaleType} with rate {item.TaxRate}%");
            }

            if (string.IsNullOrWhiteSpace(item.SroItemSerialNo))
            {
                result.IsValid = false;
                result.Errors.Add($"⚠️ SRO Item Serial Number is required for sale type: {item.SaleType} with rate {item.TaxRate}%");
            }

            return result;
        }

        /// <summary>
        /// Auto-fill SRO defaults - clears fields if SRO not required
        /// </summary>
        public static void AutoFillSroDefaults(InvoiceItem item)
        {
            if (!RequiresSroDetails(item.SaleType, item.TaxRate))
            {
                item.SroScheduleNo = "";
                item.SroItemSerialNo = "";
            }
        }

        /// <summary>
        /// Normalizes sale type to match FBR expected values
        /// </summary>
        public static string NormalizeSaleType(string saleType)
        {
            if (string.IsNullOrWhiteSpace(saleType))
                return saleType;

            string normalized = saleType.Trim();

            // Check normalization map first
            return NormalizationMap.TryGetValue(normalized, out string mappedValue) ? mappedValue : normalized;
        }

        #endregion
    }

    /// <summary>
    /// Scenario definition class
    /// </summary>
    public class ScenarioDefinition
    {
        public string ScenarioId { get; set; }
        public string Name { get; set; }
        public bool RequiresSRO { get; set; }
        public bool BuyerDependant { get; set; }

        public ScenarioDefinition() { }

        public ScenarioDefinition(string scenarioId, string name, bool requiresSRO, bool buyerDependant = false)
        {
            ScenarioId = scenarioId;
            Name = name;
            RequiresSRO = requiresSRO;
            BuyerDependant = buyerDependant;
        }
    }

    /// <summary>
    /// Validation result class
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
    }
}