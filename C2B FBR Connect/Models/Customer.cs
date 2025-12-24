using System;

namespace C2B_FBR_Connect.Models
{
    public class Customer
    {
        // Primary Keys & Identifiers
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public string QuickBooksCustomerId { get; set; }  // CustomerRef.ListID from QuickBooks

        // Customer Information
        public string CustomerName { get; set; }
        public string CustomerNTN { get; set; }
        public string CustomerStrNo { get; set; }
        public string CustomerAddress { get; set; }
        public string CustomerPhone { get; set; }
        public string CustomerEmail { get; set; }
        public string CustomerType { get; set; } = "Unregistered";  // Registered, Unregistered

        // System Timestamps
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}