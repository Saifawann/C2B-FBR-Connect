using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Models
{

    public class InvoiceItem
    {
        // Database fields
        public int Id { get; set; }
        public int InvoiceId { get; set; }

        // Basic item details
        public string ItemName { get; set; }
        public string ItemDescription { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal TotalPrice { get; set; }
        public decimal NetAmount { get; set; }

        // Tax details
        public decimal TaxRate { get; set; }
        public string Rate { get; set; }  // ✅ String for API: "18%", "Exempt", "0%", etc.

        public decimal SalesTaxAmount { get; set; }
        public decimal TotalValue { get; set; }

        // FBR-specific fields
        public string HSCode { get; set; }
        public string UnitOfMeasure { get; set; }
        public decimal RetailPrice { get; set; }
        public decimal ExtraTax { get; set; }
        public decimal FurtherTax { get; set; }
        public string SroScheduleNo { get; set; }
        public decimal FedPayable { get; set; }
        public decimal SalesTaxWithheldAtSource { get; set; }
        public decimal Discount { get; set; }
        public string SaleType { get; set; }
        public string SroItemSerialNo { get; set; }
    }

}
