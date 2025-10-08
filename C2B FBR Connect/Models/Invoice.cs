using System;
using System.Collections.Generic;

namespace C2B_FBR_Connect.Models
{
    public class Invoice
    {
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public string QuickBooksInvoiceId { get; set; }
        public string InvoiceNumber { get; set; }
        public string CustomerName { get; set; }
        public string CustomerNTN { get; set; } // Add this property
        public decimal Amount { get; set; }
        public string Status { get; set; } // Pending, Uploaded, Failed
        public string FBR_IRN { get; set; }
        public string FBR_QRCode { get; set; }
        public DateTime? UploadDate { get; set; }
        public DateTime CreatedDate { get; set; }
        public string ErrorMessage { get; set; }
    }

    public enum InvoiceStatus
    {
        Pending,
        Uploaded,
        Failed
    }

}