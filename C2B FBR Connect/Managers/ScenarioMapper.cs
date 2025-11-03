using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace C2B_FBR_Connect.Services
{
    /// <summary>
    /// Maps FBR sale types to scenarios and handles SRO requirements
    /// Based on official FBR e-invoice scenario testing guide
    /// Works with InvoiceItem model
    /// </summary>
    public class ScenarioMapper
    {
        // Scenario definitions with their SRO requirements (based on FBR official guide)
        private static readonly Dictionary<string, ScenarioDefinition> ScenarioMap = new Dictionary<string, ScenarioDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            // SN001 & SN002 - Standard Rate (buyer registration determines which)
            ["Goods at Standard Rate (default)"] = new ScenarioDefinition
            {
                ScenarioId = "SN001",
                Name = "Goods at standard rate to registered buyers",
                RequiresSRO = false,
                BuyerDependant = true // SN001 for Registered, SN002 for Unregistered
            },
            ["Goods at standard rate (default)"] = new ScenarioDefinition
            {
                ScenarioId = "SN001",
                Name = "Goods at standard rate to registered buyers",
                RequiresSRO = false,
                BuyerDependant = true
            },

            // SN003 - Steel Melting
            ["Steel Melting and re-rolling"] = new ScenarioDefinition
            {
                ScenarioId = "SN003",
                Name = "Sale of Steel (Melted and Re-Rolled)",
                RequiresSRO = false
            },
            ["Steel melting and re-rolling"] = new ScenarioDefinition
            {
                ScenarioId = "SN003",
                Name = "Sale of Steel (Melted and Re-Rolled)",
                RequiresSRO = false
            },

            // SN004 - Ship Breaking
            ["Ship breaking"] = new ScenarioDefinition
            {
                ScenarioId = "SN004",
                Name = "Sale by Ship Breakers",
                RequiresSRO = false
            },

            // SN005 - Reduced Rate ✅ REQUIRES SRO
            ["Goods at Reduced Rate"] = new ScenarioDefinition
            {
                ScenarioId = "SN005",
                Name = "Reduced rate sale",
                RequiresSRO = true
            },

            // SN006 - Exempt Goods ✅ REQUIRES SRO
            ["Exempt Goods"] = new ScenarioDefinition
            {
                ScenarioId = "SN006",
                Name = "Exempt goods sale",
                RequiresSRO = true
            },
            ["Exempt goods"] = new ScenarioDefinition
            {
                ScenarioId = "SN006",
                Name = "Exempt goods sale",
                RequiresSRO = true
            },

            // SN007 - Zero Rated ✅ REQUIRES SRO
            ["Goods at zero-rate"] = new ScenarioDefinition
            {
                ScenarioId = "SN007",
                Name = "Zero rated sale",
                RequiresSRO = true
            },

            // SN008 - 3rd Schedule (NO SRO in guide examples)
            ["3rd Schedule Goods"] = new ScenarioDefinition
            {
                ScenarioId = "SN008",
                Name = "Sale of 3rd schedule goods",
                RequiresSRO = false
            },

            // SN009 - Cotton Ginners
            ["Cotton Ginners"] = new ScenarioDefinition
            {
                ScenarioId = "SN009",
                Name = "Cotton Spinners purchase from Cotton Ginners",
                RequiresSRO = false
            },
            ["Cotton ginners"] = new ScenarioDefinition
            {
                ScenarioId = "SN009",
                Name = "Cotton Spinners purchase from Cotton Ginners",
                RequiresSRO = false
            },

            // SN010 - Telecommunication
            ["Telecommunication services"] = new ScenarioDefinition
            {
                ScenarioId = "SN010",
                Name = "Telecom services rendered or provided",
                RequiresSRO = false
            },

            // SN011 - Toll Manufacturing
            ["Toll Manufacturing"] = new ScenarioDefinition
            {
                ScenarioId = "SN011",
                Name = "Toll Manufacturing sale by Steel sector",
                RequiresSRO = false
            },

            // SN012 - Petroleum ✅ REQUIRES SRO
            ["Petroleum Products"] = new ScenarioDefinition
            {
                ScenarioId = "SN012",
                Name = "Sale of Petroleum products",
                RequiresSRO = true
            },

            // SN013 - Electricity ✅ REQUIRES SRO
            ["Electricity Supply to Retailers"] = new ScenarioDefinition
            {
                ScenarioId = "SN013",
                Name = "Electricity Supply to Retailers",
                RequiresSRO = true
            },

            // SN014 - Gas to CNG
            ["Gas to CNG stations"] = new ScenarioDefinition
            {
                ScenarioId = "SN014",
                Name = "Sale of Gas to CNG stations",
                RequiresSRO = false
            },

            // SN015 - Mobile Phones ✅ REQUIRES SRO
            ["Mobile Phones"] = new ScenarioDefinition
            {
                ScenarioId = "SN015",
                Name = "Sale of mobile phones",
                RequiresSRO = true
            },

            // SN016 - Processing/Conversion
            ["Processing/ Conversion of Goods"] = new ScenarioDefinition
            {
                ScenarioId = "SN016",
                Name = "Processing / Conversion of Goods",
                RequiresSRO = false
            },
            ["Processing/Conversion of Goods"] = new ScenarioDefinition
            {
                ScenarioId = "SN016",
                Name = "Processing / Conversion of Goods",
                RequiresSRO = false
            },

            // SN017 - Goods (FED in ST Mode)
            ["Goods (FED in ST Mode)"] = new ScenarioDefinition
            {
                ScenarioId = "SN017",
                Name = "Sale of Goods where FED is charged in ST mode",
                RequiresSRO = false
            },

            // SN018 - Services (FED in ST Mode)
            ["Services (FED in ST Mode)"] = new ScenarioDefinition
            {
                ScenarioId = "SN018",
                Name = "Services rendered or provided where FED is charged in ST mode",
                RequiresSRO = false
            },

            // SN019 - Services ✅ REQUIRES SRO
            ["Services"] = new ScenarioDefinition
            {
                ScenarioId = "SN019",
                Name = "Services rendered or provided",
                RequiresSRO = true
            },

            // SN020 - Electric Vehicle ✅ REQUIRES SRO
            ["Electric Vehicle"] = new ScenarioDefinition
            {
                ScenarioId = "SN020",
                Name = "Sale of Electric Vehicles",
                RequiresSRO = true
            },

            // SN021 - Cement (NO SRO in guide)
            ["Cement /Concrete Block"] = new ScenarioDefinition
            {
                ScenarioId = "SN021",
                Name = "Sale of Cement /Concrete Block",
                RequiresSRO = false
            },

            // SN022 - Potassium Chlorate ✅ REQUIRES SRO
            ["Potassium Chlorate"] = new ScenarioDefinition
            {
                ScenarioId = "SN022",
                Name = "Sale of Potassium Chlorate",
                RequiresSRO = true
            },

            // SN023 - CNG Sales ✅ REQUIRES SRO
            ["CNG Sales"] = new ScenarioDefinition
            {
                ScenarioId = "SN023",
                Name = "Sale of CNG",
                RequiresSRO = true
            },

            // SN024 - SRO 297 ✅ REQUIRES SRO
            ["Goods as per SRO.297(|)/2023"] = new ScenarioDefinition
            {
                ScenarioId = "SN024",
                Name = "Goods sold that are listed in SRO 297(1)/2023",
                RequiresSRO = true
            },
            ["Goods as per SRO.297(I)/2023"] = new ScenarioDefinition
            {
                ScenarioId = "SN024",
                Name = "Goods sold that are listed in SRO 297(1)/2023",
                RequiresSRO = true
            },

            // SN025 - Non-Adjustable Supplies ✅ REQUIRES SRO
            ["Non-Adjustable Supplies"] = new ScenarioDefinition
            {
                ScenarioId = "SN025",
                Name = "Drugs sold at fixed ST rate under serial 81 of Eighth Schedule Table 1",
                RequiresSRO = true
            },

            // SN026 - Retailer to End Consumer (Standard Rate)
            ["Goods at standard rate to end consumer"] = new ScenarioDefinition
            {
                ScenarioId = "SN026",
                Name = "Sale to End Consumer by retailers",
                RequiresSRO = false
            },

            // SN027 - Retailer 3rd Schedule
            ["3rd Schedule Goods to end consumer"] = new ScenarioDefinition
            {
                ScenarioId = "SN027",
                Name = "Sale to End Consumer by retailers - 3rd Schedule",
                RequiresSRO = false
            },

            // SN028 - Retailer Reduced Rate ✅ REQUIRES SRO
            ["Goods at Reduced Rate to end consumer"] = new ScenarioDefinition
            {
                ScenarioId = "SN028",
                Name = "Sale to End Consumer by retailers - Reduced Rate",
                RequiresSRO = true
            }
        };

        /// <summary>
        /// Determines scenario ID based on sale type and buyer registration
        /// </summary>
        public static string DetermineScenarioId(string saleType, string buyerRegistrationType = "Registered")
        {
            if (string.IsNullOrEmpty(saleType))
                return "SN001"; // Default scenario

            // ✅ Trim whitespace from input
            string trimmedSaleType = saleType?.Trim();

            // Check for exact match first
            if (ScenarioMap.TryGetValue(trimmedSaleType, out var scenario))
            {
                // Special handling for buyer-dependent scenarios (SN001/SN002)
                if (scenario.BuyerDependant && buyerRegistrationType == "Unregistered")
                {
                    return "SN002"; // Goods at standard rate to unregistered buyers
                }

                return scenario.ScenarioId;
            }

            // Check for case-insensitive partial match with trimming
            var matchedScenario = ScenarioMap
                .FirstOrDefault(kvp =>
                    trimmedSaleType.Contains(kvp.Key.Trim(), StringComparison.OrdinalIgnoreCase) ||
                    kvp.Key.Trim().Contains(trimmedSaleType, StringComparison.OrdinalIgnoreCase));

            if (!matchedScenario.Equals(default(KeyValuePair<string, ScenarioDefinition>)))
            {
                if (matchedScenario.Value.BuyerDependant && buyerRegistrationType == "Unregistered")
                {
                    return "SN002";
                }
                return matchedScenario.Value.ScenarioId;
            }

            // Default to SN001 if no match found
            return "SN001";
        }

        /// <summary>
        /// Check if scenario requires SRO details
        /// </summary>
        public static bool RequiresSroDetails(string saleType, decimal taxRate)
        {
            // ✅ CRITICAL: If rate is not 18%, SRO is ALWAYS required (FBR Rule)
            if (taxRate != 18m)
            {
                return true;
            }

            // For 18% rate, check if sale type requires SRO
            if (string.IsNullOrEmpty(saleType))
                return false;

            // ✅ Trim whitespace from input
            string trimmedSaleType = saleType?.Trim();

            if (ScenarioMap.TryGetValue(trimmedSaleType, out var scenario))
            {
                return scenario.RequiresSRO;
            }

            // Check for partial match with trimming
            var matchedScenario = ScenarioMap
                .FirstOrDefault(kvp =>
                    trimmedSaleType.Contains(kvp.Key.Trim(), StringComparison.OrdinalIgnoreCase) ||
                    kvp.Key.Trim().Contains(trimmedSaleType, StringComparison.OrdinalIgnoreCase));

            return !matchedScenario.Equals(default(KeyValuePair<string, ScenarioDefinition>)) &&
                   matchedScenario.Value.RequiresSRO;
        }


        /// <summary>
        /// Get scenario definition for display/validation
        /// </summary>
        public static ScenarioDefinition GetScenarioDefinition(string saleType)
        {
            if (string.IsNullOrEmpty(saleType))
                return new ScenarioDefinition { ScenarioId = "SN001", Name = "Default", RequiresSRO = false };

            // ✅ Trim whitespace from input
            string trimmedSaleType = saleType?.Trim();

            if (ScenarioMap.TryGetValue(trimmedSaleType, out var scenario))
            {
                return scenario;
            }

            // Check for partial match with trimming
            var matchedScenario = ScenarioMap
                .FirstOrDefault(kvp =>
                    trimmedSaleType.Contains(kvp.Key.Trim(), StringComparison.OrdinalIgnoreCase) ||
                    kvp.Key.Trim().Contains(trimmedSaleType, StringComparison.OrdinalIgnoreCase));

            if (!matchedScenario.Equals(default(KeyValuePair<string, ScenarioDefinition>)))
            {
                return matchedScenario.Value;
            }

            return new ScenarioDefinition { ScenarioId = "SN001", Name = "Default", RequiresSRO = false };
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

            // ✅ Check BOTH sale type AND tax rate
            if (!RequiresSroDetails(item.SaleType, item.TaxRate))
                return result; // No SRO required

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
        /// Auto-fill SRO defaults - now just clears fields if SRO not required
        /// Actual SRO data should be fetched from API via SroDataService
        /// </summary>
        public static void AutoFillSroDefaults(InvoiceItem item)
        {
            // ✅ Check BOTH sale type AND tax rate
            if (!RequiresSroDetails(item.SaleType, item.TaxRate))
            {
                // Clear SRO fields if not required
                item.SroScheduleNo = "";
                item.SroItemSerialNo = "";
                System.Diagnostics.Debug.WriteLine($"✅ Cleared SRO fields for '{item.ItemName}' (rate={item.TaxRate}%, not required)");
            }
            else
            {
                // SRO is required but not filled
                if (string.IsNullOrWhiteSpace(item.SroScheduleNo) || string.IsNullOrWhiteSpace(item.SroItemSerialNo))
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ SRO data required for '{item.ItemName}' (rate={item.TaxRate}%) - must be fetched from API");
                }
            }
        }
    }

    /// <summary>
    /// Scenario definition class
    /// </summary>
    public class ScenarioDefinition
    {
        public string ScenarioId { get; set; }
        public string Name { get; set; }
        public bool RequiresSRO { get; set; }
        public bool BuyerDependant { get; set; } // For scenarios that change based on buyer registration
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