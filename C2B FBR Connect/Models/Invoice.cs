using System;
using System.Collections.Generic;

namespace C2B_FBR_Connect.Models
{
    public class Invoice
    {
        // Primary Keys & Identifiers
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public string QuickBooksInvoiceId { get; set; }
        public string InvoiceNumber { get; set; }

        // Customer Information
        public string CustomerName { get; set; }
        public string CustomerNTN { get; set; }
        public string CustomerAddress { get; set; }
        public string CustomerPhone { get; set; }
        public string CustomerEmail { get; set; }

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
    }
    public enum InvoiceStatus
    {
        Pending,
        Uploaded,
        Failed
    }

}