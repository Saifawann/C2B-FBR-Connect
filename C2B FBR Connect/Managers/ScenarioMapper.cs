using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace C2B_FBR_Connect.Services
{
    /// <summary>
    /// Maps FBR sale types to scenarios and handles SRO requirements
    /// </summary>
    public static class ScenarioMapper
    {
        // ✅ QuickBooks short names (≤30 chars) for custom field
        public static class QBSaleTypes
        {
            public const string StandardRate = "Standard rate";                  // 13 - SN001
            public const string UnregisteredBuyer = "Unregistered buyer";        // 18 - SN002
            public const string SteelMelting = "Steel melting/re-rolling";       // 24 - SN003
            public const string ShipBreaking = "Ship breaking";                  // 13 - SN004
            public const string ReducedRate = "Reduced rate";                    // 12 - SN005
            public const string Exempt = "Exempt";                               // 6  - SN006
            public const string ZeroRate = "Zero-rate";                          // 9  - SN007
            public const string ThirdSchedule = "3rd Schedule";                  // 12 - SN008
            public const string CottonGinners = "Cotton ginners";                // 14 - SN009
            public const string Telecom = "Telecom services";                    // 16 - SN010
            public const string TollManufacturing = "Toll manufacturing";        // 18 - SN011
            public const string Petroleum = "Petroleum";                         // 9  - SN012
            public const string ElectricityRetail = "Electricity-retailers";     // 21 - SN013
            public const string GasCNG = "Gas to CNG";                           // 10 - SN014
            public const string MobilePhones = "Mobile phones";                  // 13 - SN015
            public const string Processing = "Processing/Conversion";            // 21 - SN016
            public const string GoodsFED = "Goods-FED ST mode";                  // 17 - SN017
            public const string ServicesFED = "Services-FED ST mode";            // 20 - SN018
            public const string Services = "Services";                           // 8  - SN019
            public const string ElectricVehicle = "Electric vehicle";            // 16 - SN020
            public const string Cement = "Cement/Concrete";                      // 15 - SN021
            public const string PotassiumChlorate = "Potassium chlorate";        // 18 - SN022
            public const string CNGSales = "CNG sales";                          // 9  - SN023
            public const string SRO297 = "SRO 297(I)/2023";                      // 15 - SN024
            public const string NonAdjustable = "Non-adjustable";                // 14 - SN025
            public const string ConsumerStandard = "Consumer-standard";          // 17 - SN026
            public const string Consumer3rdSchedule = "Consumer-3rd Sched";      // 18 - SN027
            public const string ConsumerReduced = "Consumer-reduced";            // 16 - SN028
        }

        // ✅ FBR API exact values (what FBR expects)
        public static class FBRSaleTypes
        {
            public const string StandardRate = "Goods at standard rate (default)";
            public const string UnregisteredBuyer = "Goods at standard rate (default)"; // Same, scenario changes
            public const string SteelMelting = "Steel melting and re-rolling";
            public const string ShipBreaking = "Ship breaking";
            public const string ReducedRate = "Goods at Reduced Rate";
            public const string Exempt = "Exempt goods";
            public const string ZeroRate = "Goods at zero-rate";
            public const string ThirdSchedule = "3rd Schedule Goods";
            public const string CottonGinners = "Cotton ginners";
            public const string Telecom = "Telecommunication services";
            public const string TollManufacturing = "Toll Manufacturing";
            public const string Petroleum = "Petroleum Products";
            public const string ElectricityRetail = "Electricity Supply to Retailers";
            public const string GasCNG = "Gas to CNG stations";
            public const string MobilePhones = "Mobile Phones";
            public const string Processing = "Processing/Conversion of Goods";
            public const string GoodsFED = "Goods (FED in ST Mode)";
            public const string ServicesFED = "Services (FED in ST Mode)";
            public const string Services = "Services";
            public const string ElectricVehicle = "Electric Vehicle";
            public const string Cement = "Cement /Concrete Block";
            public const string PotassiumChlorate = "Potassium Chlorate";
            public const string CNGSales = "CNG Sales";
            public const string SRO297 = "Goods as per SRO.297(|)/2023";
            public const string NonAdjustable = "Non-Adjustable Supplies";
            public const string ConsumerStandard = "Goods at standard rate (default)";
            public const string Consumer3rdSchedule = "3rd Schedule Goods";
            public const string ConsumerReduced = "Goods at Reduced Rate";
        }

        // ✅ QB Short Name → FBR API Value
        private static readonly Dictionary<string, string> QBToFBRMap = new(StringComparer.OrdinalIgnoreCase)
        {
            [QBSaleTypes.StandardRate] = FBRSaleTypes.StandardRate,
            [QBSaleTypes.UnregisteredBuyer] = FBRSaleTypes.UnregisteredBuyer,
            [QBSaleTypes.SteelMelting] = FBRSaleTypes.SteelMelting,
            [QBSaleTypes.ShipBreaking] = FBRSaleTypes.ShipBreaking,
            [QBSaleTypes.ReducedRate] = FBRSaleTypes.ReducedRate,
            [QBSaleTypes.Exempt] = FBRSaleTypes.Exempt,
            [QBSaleTypes.ZeroRate] = FBRSaleTypes.ZeroRate,
            [QBSaleTypes.ThirdSchedule] = FBRSaleTypes.ThirdSchedule,
            [QBSaleTypes.CottonGinners] = FBRSaleTypes.CottonGinners,
            [QBSaleTypes.Telecom] = FBRSaleTypes.Telecom,
            [QBSaleTypes.TollManufacturing] = FBRSaleTypes.TollManufacturing,
            [QBSaleTypes.Petroleum] = FBRSaleTypes.Petroleum,
            [QBSaleTypes.ElectricityRetail] = FBRSaleTypes.ElectricityRetail,
            [QBSaleTypes.GasCNG] = FBRSaleTypes.GasCNG,
            [QBSaleTypes.MobilePhones] = FBRSaleTypes.MobilePhones,
            [QBSaleTypes.Processing] = FBRSaleTypes.Processing,
            [QBSaleTypes.GoodsFED] = FBRSaleTypes.GoodsFED,
            [QBSaleTypes.ServicesFED] = FBRSaleTypes.ServicesFED,
            [QBSaleTypes.Services] = FBRSaleTypes.Services,
            [QBSaleTypes.ElectricVehicle] = FBRSaleTypes.ElectricVehicle,
            [QBSaleTypes.Cement] = FBRSaleTypes.Cement,
            [QBSaleTypes.PotassiumChlorate] = FBRSaleTypes.PotassiumChlorate,
            [QBSaleTypes.CNGSales] = FBRSaleTypes.CNGSales,
            [QBSaleTypes.SRO297] = FBRSaleTypes.SRO297,
            [QBSaleTypes.NonAdjustable] = FBRSaleTypes.NonAdjustable,
            [QBSaleTypes.ConsumerStandard] = FBRSaleTypes.ConsumerStandard,
            [QBSaleTypes.Consumer3rdSchedule] = FBRSaleTypes.Consumer3rdSchedule,
            [QBSaleTypes.ConsumerReduced] = FBRSaleTypes.ConsumerReduced,
        };

        // ✅ Legacy/alternate names → QB short name (for backward compatibility)
        private static readonly Dictionary<string, string> AliasToQBMap = new(StringComparer.OrdinalIgnoreCase)
        {
            // Old FBR names → QB short
            ["Goods at Standard Rate (default)"] = QBSaleTypes.StandardRate,
            ["Goods at standard rate (default)"] = QBSaleTypes.StandardRate,
            ["Goods at Standard Rate"] = QBSaleTypes.StandardRate,
            ["Goods at standard rate"] = QBSaleTypes.StandardRate,
            ["SN001 Standard"] = QBSaleTypes.StandardRate,

            ["Steel Melting and re-rolling"] = QBSaleTypes.SteelMelting,
            ["Steel melting and re-rolling"] = QBSaleTypes.SteelMelting,

            ["Goods at Reduced Rate"] = QBSaleTypes.ReducedRate,
            ["Goods at reduced rate"] = QBSaleTypes.ReducedRate,

            ["Exempt goods"] = QBSaleTypes.Exempt,
            ["Exempt Goods"] = QBSaleTypes.Exempt,

            ["Goods at zero-rate"] = QBSaleTypes.ZeroRate,
            ["Zero-rated goods"] = QBSaleTypes.ZeroRate,
            ["Zero Rate Goods"] = QBSaleTypes.ZeroRate,

            ["3rd Schedule Goods"] = QBSaleTypes.ThirdSchedule,
            ["3rd Schedule goods"] = QBSaleTypes.ThirdSchedule,
            ["Third Schedule Goods"] = QBSaleTypes.ThirdSchedule,

            ["Cotton Ginners"] = QBSaleTypes.CottonGinners,

            ["Telecommunication services"] = QBSaleTypes.Telecom,
            ["Telecommunication Services"] = QBSaleTypes.Telecom,
            ["Telecom"] = QBSaleTypes.Telecom,

            ["Toll Manufacturing"] = QBSaleTypes.TollManufacturing,

            ["Petroleum Products"] = QBSaleTypes.Petroleum,
            ["Petroleum products"] = QBSaleTypes.Petroleum,

            ["Electricity Supply to Retailers"] = QBSaleTypes.ElectricityRetail,
            ["Electricity to Retailers"] = QBSaleTypes.ElectricityRetail,
            ["Electricity to retailers"] = QBSaleTypes.ElectricityRetail,

            ["Gas to CNG stations"] = QBSaleTypes.GasCNG,
            ["Gas to CNG Stations"] = QBSaleTypes.GasCNG,

            ["Mobile Phones"] = QBSaleTypes.MobilePhones,

            ["Processing/ Conversion of Goods"] = QBSaleTypes.Processing,
            ["Processing/Conversion of Goods"] = QBSaleTypes.Processing,
            ["Processing/Conversion of goods"] = QBSaleTypes.Processing,

            ["Goods (FED in ST Mode)"] = QBSaleTypes.GoodsFED,
            ["Goods (FED in ST mode)"] = QBSaleTypes.GoodsFED,

            ["Services (FED in ST Mode)"] = QBSaleTypes.ServicesFED,
            ["Services (FED in ST mode)"] = QBSaleTypes.ServicesFED,

            ["Service"] = QBSaleTypes.Services,

            ["Electric Vehicle"] = QBSaleTypes.ElectricVehicle,
            ["Electric Vehicles"] = QBSaleTypes.ElectricVehicle,

            ["Cement /Concrete Block"] = QBSaleTypes.Cement,
            ["Cement/Concrete Block"] = QBSaleTypes.Cement,
            ["Cement/Concrete block"] = QBSaleTypes.Cement,

            ["Potassium Chlorate"] = QBSaleTypes.PotassiumChlorate,

            ["CNG Sales"] = QBSaleTypes.CNGSales,
            ["CNG"] = QBSaleTypes.CNGSales,

            ["Goods as per SRO.297(|)/2023"] = QBSaleTypes.SRO297,
            ["Goods as per SRO.297(I)/2023"] = QBSaleTypes.SRO297,
            ["Goods as per SRO 297(I)/2023"] = QBSaleTypes.SRO297,

            ["Non-Adjustable Supplies"] = QBSaleTypes.NonAdjustable,
            ["Non-adjustable supplies"] = QBSaleTypes.NonAdjustable,

            ["Goods at standard rate to end consumer"] = QBSaleTypes.ConsumerStandard,
            ["3rd Schedule Goods to end consumer"] = QBSaleTypes.Consumer3rdSchedule,
            ["Goods at Reduced Rate to end consumer"] = QBSaleTypes.ConsumerReduced,
        };

        // ✅ Scenario definitions (keyed by QB short name)
        private static readonly Dictionary<string, ScenarioDefinition> ScenarioMap = new(StringComparer.OrdinalIgnoreCase)
        {
            [QBSaleTypes.StandardRate] = new("SN001", "Standard rate to registered", false, true),
            [QBSaleTypes.UnregisteredBuyer] = new("SN002", "Standard rate to unregistered", false),
            [QBSaleTypes.SteelMelting] = new("SN003", "Steel melting and re-rolling", false),
            [QBSaleTypes.ShipBreaking] = new("SN004", "Ship breakers", false),
            [QBSaleTypes.ReducedRate] = new("SN005", "Reduced rate", true),
            [QBSaleTypes.Exempt] = new("SN006", "Exempt goods", true),
            [QBSaleTypes.ZeroRate] = new("SN007", "Zero rated", true),
            [QBSaleTypes.ThirdSchedule] = new("SN008", "3rd schedule goods", true),
            [QBSaleTypes.CottonGinners] = new("SN009", "Cotton ginners", false),
            [QBSaleTypes.Telecom] = new("SN010", "Telecom services", false),
            [QBSaleTypes.TollManufacturing] = new("SN011", "Toll manufacturing", false),
            [QBSaleTypes.Petroleum] = new("SN012", "Petroleum products", true),
            [QBSaleTypes.ElectricityRetail] = new("SN013", "Electricity to retailers", true),
            [QBSaleTypes.GasCNG] = new("SN014", "Gas to CNG stations", false),
            [QBSaleTypes.MobilePhones] = new("SN015", "Mobile phones", true),
            [QBSaleTypes.Processing] = new("SN016", "Processing/Conversion", false),
            [QBSaleTypes.GoodsFED] = new("SN017", "Goods FED in ST mode", false),
            [QBSaleTypes.ServicesFED] = new("SN018", "Services FED in ST mode", false),
            [QBSaleTypes.Services] = new("SN019", "Services", true),
            [QBSaleTypes.ElectricVehicle] = new("SN020", "Electric vehicles", true),
            [QBSaleTypes.Cement] = new("SN021", "Cement/Concrete block", false),
            [QBSaleTypes.PotassiumChlorate] = new("SN022", "Potassium chlorate", true),
            [QBSaleTypes.CNGSales] = new("SN023", "CNG sales", true),
            [QBSaleTypes.SRO297] = new("SN024", "SRO 297(I)/2023 goods", true),
            [QBSaleTypes.NonAdjustable] = new("SN025", "Non-adjustable supplies", true),
            [QBSaleTypes.ConsumerStandard] = new("SN026", "End consumer standard", false),
            [QBSaleTypes.Consumer3rdSchedule] = new("SN027", "End consumer 3rd schedule", true),
            [QBSaleTypes.ConsumerReduced] = new("SN028", "End consumer reduced", true),
        };

        #region Public Methods

        /// <summary>
        /// Converts any sale type (QB short, FBR long, or alias) to FBR API format
        /// This is what gets sent to FBR
        /// </summary>
        public static string ToFBRSaleType(string saleType)
        {
            if (string.IsNullOrWhiteSpace(saleType))
                return FBRSaleTypes.StandardRate;

            string trimmed = saleType.Trim();

            // Already a QB short name? Map to FBR
            if (QBToFBRMap.TryGetValue(trimmed, out string fbrValue))
                return fbrValue;

            // Is it an alias? Convert to QB first, then to FBR
            if (AliasToQBMap.TryGetValue(trimmed, out string qbName))
            {
                if (QBToFBRMap.TryGetValue(qbName, out string fbrFromAlias))
                    return fbrFromAlias;
            }

            // Already an FBR value? Return as-is
            if (QBToFBRMap.Values.Contains(trimmed, StringComparer.OrdinalIgnoreCase))
                return trimmed;

            // Partial match
            var partial = AliasToQBMap.FirstOrDefault(kvp =>
                trimmed.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase) ||
                kvp.Key.Contains(trimmed, StringComparison.OrdinalIgnoreCase));

            if (!partial.Equals(default(KeyValuePair<string, string>)))
            {
                if (QBToFBRMap.TryGetValue(partial.Value, out string fbrPartial))
                    return fbrPartial;
            }

            return FBRSaleTypes.StandardRate;
        }

        /// <summary>
        /// Converts any sale type to QB short name (for storage/display)
        /// </summary>
        public static string ToQBSaleType(string saleType)
        {
            if (string.IsNullOrWhiteSpace(saleType))
                return QBSaleTypes.StandardRate;

            string trimmed = saleType.Trim();

            // Already a QB short name?
            if (QBToFBRMap.ContainsKey(trimmed))
                return trimmed;

            // Is it an alias?
            if (AliasToQBMap.TryGetValue(trimmed, out string qbName))
                return qbName;

            // Partial match
            var partial = AliasToQBMap.FirstOrDefault(kvp =>
                trimmed.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase) ||
                kvp.Key.Contains(trimmed, StringComparison.OrdinalIgnoreCase));

            if (!partial.Equals(default(KeyValuePair<string, string>)))
                return partial.Value;

            return QBSaleTypes.StandardRate;
        }

        /// <summary>
        /// LEGACY: Alias for ToFBRSaleType (backward compatibility)
        /// </summary>
        public static string NormalizeSaleType(string saleType) => ToFBRSaleType(saleType);

        /// <summary>
        /// Get scenario ID from sale type
        /// </summary>
        public static string GetScenarioId(string saleType, string buyerType = "Registered")
        {
            string qbType = ToQBSaleType(saleType);

            if (ScenarioMap.TryGetValue(qbType, out var scenario))
            {
                if (scenario.BuyerDependant && buyerType == "Unregistered")
                    return "SN002";
                return scenario.ScenarioId;
            }

            return "SN001";
        }

        /// <summary>
        /// LEGACY: Alias for GetScenarioId
        /// </summary>
        public static string DetermineScenarioId(string saleType, string buyerType = "Registered")
            => GetScenarioId(saleType, buyerType);

        /// <summary>
        /// Check if SRO details required
        /// </summary>
        public static bool RequiresSroDetails(string saleType, decimal taxRate)
        {
            // Non-standard rates need SRO (except 0%)
            if (taxRate != 18m && taxRate != 0m)
                return true;

            string qbType = ToQBSaleType(saleType);
            return ScenarioMap.TryGetValue(qbType, out var scenario) && scenario.RequiresSRO;
        }

        /// <summary>
        /// Get scenario definition
        /// </summary>
        public static ScenarioDefinition GetScenarioDefinition(string saleType)
        {
            string qbType = ToQBSaleType(saleType);
            return ScenarioMap.TryGetValue(qbType, out var scenario)
                ? scenario
                : new ScenarioDefinition("SN001", "Default", false);
        }

        /// <summary>
        /// Get all QB sale types for dropdown (short names ≤30 chars)
        /// </summary>
        public static List<string> GetAllQBSaleTypes()
        {
            return QBToFBRMap.Keys.OrderBy(k => k).ToList();
        }

        /// <summary>
        /// LEGACY: Alias for GetAllQBSaleTypes
        /// </summary>
        public static List<string> GetAllSaleTypes() => GetAllQBSaleTypes();

        /// <summary>
        /// Validate SRO data
        /// </summary>
        public static ValidationResult ValidateSroData(InvoiceItem item)
        {
            var result = new ValidationResult { IsValid = true };

            if (!RequiresSroDetails(item.SaleType, item.TaxRate))
                return result;

            if (string.IsNullOrWhiteSpace(item.SroScheduleNo))
            {
                result.IsValid = false;
                result.Errors.Add($"SRO Schedule required for {item.SaleType} at {item.TaxRate}%");
            }

            if (string.IsNullOrWhiteSpace(item.SroItemSerialNo))
            {
                result.IsValid = false;
                result.Errors.Add($"SRO Serial required for {item.SaleType} at {item.TaxRate}%");
            }

            return result;
        }

        /// <summary>
        /// Clear SRO if not required
        /// </summary>
        public static void AutoFillSroDefaults(InvoiceItem item)
        {
            if (!RequiresSroDetails(item.SaleType, item.TaxRate))
            {
                item.SroScheduleNo = "";
                item.SroItemSerialNo = "";
            }
        }

        #endregion
    }

    public class ScenarioDefinition
    {
        public string ScenarioId { get; set; }
        public string Name { get; set; }
        public bool RequiresSRO { get; set; }
        public bool BuyerDependant { get; set; }

        public ScenarioDefinition() { }
        public ScenarioDefinition(string id, string name, bool sro, bool buyer = false)
        {
            ScenarioId = id;
            Name = name;
            RequiresSRO = sro;
            BuyerDependant = buyer;
        }
    }

    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> Errors { get; set; } = new();
        public List<string> Warnings { get; set; } = new();
    }
}