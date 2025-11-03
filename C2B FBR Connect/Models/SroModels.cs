using System;

namespace C2B_FBR_Connect.Models
{
    // Response model for SaleTypeToRate API
    public class SaleTypeRate
    {
        public int RATE_ID { get; set; }
        public string RATE_DESC { get; set; }
        public decimal RATE_VALUE { get; set; }
    }

    // Response model for SroSchedule API
    public class SroSchedule
    {
        public int SRO_ID { get; set; }
        public int SerNo { get; set; } 
        public string SRO_DESC { get; set; }
    }

    // Response model for SROItem API
    public class SroItem
    {
        public int SRO_ITEM_ID { get; set; }
        public string SRO_ITEM_DESC { get; set; }
    }
}