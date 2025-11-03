using System;
using System.Collections.Generic;

namespace C2B_FBR_Connect.Models
{
    // ✅ Response model for FBR invoice upload
    public class FBRResponse
    {
        public bool Success { get; set; }
        public string IRN { get; set; }
        public string ErrorMessage { get; set; }
        public string ResponseData { get; set; }
    }

    // ✅ Main payload structure for FBR invoice upload
    public class FBRInvoicePayload

    {
        public string ScenarioId { get; set; } = "SN001"; 
        public string BuyerRegistrationType { get; set; } = "Registered";

        // Basic invoice details
        public string InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public string InvoiceType { get; set; } // e.g., "Sale Invoice"

        // Seller (Company) details
        public string SellerNTN { get; set; }
        public string SellerBusinessName { get; set; }
        public string SellerProvince { get; set; }
        public string SellerAddress { get; set; }

        // Buyer (Customer) details
        public string CustomerName { get; set; }
        public string CustomerNTN { get; set; }
        public string BuyerProvince { get; set; }
        public string BuyerAddress { get; set; }
        public string BuyerPhone { get; set; }
        public string BuyerEmail { get; set; }

        // Financial summary
        public decimal Subtotal { get; set; }       // Amount before tax
        public decimal TaxAmount { get; set; }      // Total tax amount
        public decimal TotalAmount { get; set; }    // Grand total

        // Line items
        public List<InvoiceItem> Items { get; set; }

        public FBRInvoicePayload()
        {
            InvoiceDate = DateTime.Now;
            InvoiceType = "Sale Invoice";
            Items = new List<InvoiceItem>();
        }
    }

    // ✅ Line item structure for invoice

}
