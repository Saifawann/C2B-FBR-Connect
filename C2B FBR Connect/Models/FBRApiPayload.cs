using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace C2B_FBR_Connect.Models
{
    /// <summary>
    /// Payload format for FBR API - uses exact camelCase property names required by FBR
    /// </summary>
    public class FBRApiPayload
    {
        [JsonProperty("invoiceType")]
        public string InvoiceType { get; set; }

        [JsonProperty("invoiceDate")]
        public string InvoiceDate { get; set; }

        [JsonProperty("sellerBusinessName")]
        public string SellerBusinessName { get; set; }

        [JsonProperty("sellerProvince")]
        public string SellerProvince { get; set; }

        [JsonProperty("sellerNTNCNIC")]
        public string SellerNTNCNIC { get; set; }

        [JsonProperty("sellerAddress")]
        public string SellerAddress { get; set; }

        [JsonProperty("buyerNTNCNIC")]
        public string BuyerNTNCNIC { get; set; }

        [JsonProperty("buyerBusinessName")]
        public string BuyerBusinessName { get; set; }

        [JsonProperty("buyerProvince")]
        public string BuyerProvince { get; set; }

        [JsonProperty("buyerAddress")]
        public string BuyerAddress { get; set; }

        [JsonProperty("invoiceRefNo")]
        public string InvoiceRefNo { get; set; }

        [JsonProperty("scenarioId", NullValueHandling = NullValueHandling.Ignore)]
        public string ScenarioId { get; set; }

        [JsonProperty("buyerRegistrationType")]
        public string BuyerRegistrationType { get; set; }

        [JsonProperty("items")]
        public List<FBRApiItem> Items { get; set; }

        public FBRApiPayload()
        {
            Items = new List<FBRApiItem>();
        }
    }

    /// <summary>
    /// Line item format for FBR API
    /// </summary>
    public class FBRApiItem
    {
        [JsonProperty("hsCode")]
        public string HsCode { get; set; }

        [JsonProperty("productDescription")]
        public string ProductDescription { get; set; }

        [JsonProperty("rate")]
        public string Rate { get; set; }

        [JsonProperty("uoM")]
        public string UoM { get; set; }

        [JsonProperty("quantity")]
        public int Quantity { get; set; }

        [JsonProperty("totalValues")]
        public decimal TotalValues { get; set; }

        [JsonProperty("valueSalesExcludingST")]
        public decimal ValueSalesExcludingST { get; set; }

        [JsonProperty("fixedNotifiedValueOrRetailPrice")]
        public decimal FixedNotifiedValueOrRetailPrice { get; set; }

        [JsonProperty("salesTaxApplicable")]
        public decimal SalesTaxApplicable { get; set; }

        [JsonProperty("salesTaxWithheldAtSource")]
        public decimal SalesTaxWithheldAtSource { get; set; }

        [JsonProperty("extraTax")]
        public string ExtraTax { get; set; }

        [JsonProperty("furtherTax")]
        public decimal FurtherTax { get; set; }

        [JsonProperty("sroScheduleNo")]
        public string SroScheduleNo { get; set; }

        [JsonProperty("fedPayable")]
        public decimal FedPayable { get; set; }

        [JsonProperty("discount")]
        public decimal Discount { get; set; }

        [JsonProperty("saleType")]
        public string SaleType { get; set; }

        [JsonProperty("sroItemSerialNo")]
        public string SroItemSerialNo { get; set; }
    }
}