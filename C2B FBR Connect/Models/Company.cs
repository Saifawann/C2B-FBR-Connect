using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Models
{
    public class Company
    {
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public string Environment { get; set; } = "Sandbox";
        public string FBRToken { get; set; }

        // Seller Information for FBR
        public string SellerNTN { get; set; }
        public string SellerAddress { get; set; }
        public string SellerProvince { get; set; }
        public string SellerPhone { get; set; }
        public string SellerEmail { get; set; }

        public byte[] LogoImage { get; set; }

        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}