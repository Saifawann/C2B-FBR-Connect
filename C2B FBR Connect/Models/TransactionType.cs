using System;

namespace C2B_FBR_Connect.Models
{
    public class TransactionType
    {
        public int Id { get; set; }
        public int TransactionTypeId { get; set; }
        public string TransactionDesc { get; set; }
        public DateTime LastUpdated { get; set; }
    }
}