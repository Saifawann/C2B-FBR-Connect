using System;
using System.Collections.Generic;

namespace C2B_FBR_Connect.Models
{
    /// <summary>
    /// Fully normalized Invoice model - customer data accessed via CustomerId reference
    /// </summary>
    public class Invoice
    {
        // Primary Keys & Identifiers
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public string QuickBooksInvoiceId { get; set; }
        public string InvoiceNumber { get; set; }
        public string InvoiceType { get; set; }

        // ✅ Customer Reference (Foreign Key)
        public int CustomerId { get; set; }

        // ✅ Navigation Property (populated via JOIN or separate query)
        public Customer Customer { get; set; }

        // ✅ QuickBooks Customer ID (used for fetching/syncing)
        public string QuickBooksCustomerId { get; set; }

        // Invoice Financial Details
        public decimal Amount { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal TaxAmount { get; set; }
        public decimal DiscountAmount { get; set; }

        // Invoice Metadata
        public DateTime InvoiceDate { get; set; }
        public string PaymentMode { get; set; } // Cash, Credit, Bank Transfer, etc.
        public string Status { get; set; } // Pending, Uploaded, Failed

        // FBR Related
        public string FBR_IRN { get; set; }
        public string FBR_QRCode { get; set; }
        public DateTime? UploadDate { get; set; }
        public string ErrorMessage { get; set; }

        // System Timestamps
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }

        // Invoice Line Items
        public List<InvoiceItem> Items { get; set; } = new List<InvoiceItem>();

        // ✅ Helper Properties (computed from Customer navigation property)
        public string CustomerName => Customer?.CustomerName ?? "Unknown";
        public string CustomerNTN => Customer?.CustomerNTN;
        public string CustomerAddress => Customer?.CustomerAddress;
        public string CustomerPhone => Customer?.CustomerPhone;
        public string CustomerEmail => Customer?.CustomerEmail;
        public string CustomerType => Customer?.CustomerType ?? "Unregistered";
    }

    public enum InvoiceStatus
    {
        Pending,
        Uploaded,
        Failed
    }
}