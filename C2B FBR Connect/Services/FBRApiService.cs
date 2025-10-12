using C2B_FBR_Connect.Models;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Services
{
    public class FBRApiService : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl = "https://gw.fbr.gov.pk/di_data/v1/di/";

        public FBRApiService()
        {
            _httpClient = new HttpClient();
        }

        public async Task<FBRResponse> UploadInvoice(FBRInvoicePayload invoice, string authToken)
        {
            try
            {
                // ✅ Basic validation
                if (string.IsNullOrEmpty(invoice.SellerNTN))
                {
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = "Seller NTN is required. Please configure it in Company Settings."
                    };
                }

                // ✅ Map payload into FBR's expected format
                var fbrBody = new
                {
                    invoiceType = invoice.InvoiceType,
                    invoiceDate = invoice.InvoiceDate.ToString("yyyy-MM-dd"),
                    invoiceRefNo = invoice.InvoiceNumber,
                    scenarioId = "SN002", // default scenario
                    sellerNTNCNIC = invoice.SellerNTN,
                    sellerBusinessName = invoice.SellerBusinessName,
                    sellerAddress = invoice.SellerAddress,
                    sellerProvince = invoice.SellerProvince,
                    buyerNTNCNIC = invoice.CustomerNTN,
                    buyerBusinessName = invoice.CustomerName,
                    buyerAddress = invoice.BuyerAddress,
                    buyerProvince = invoice.BuyerProvince,
                    buyerRegistrationType = invoice.BuyerRegistrationType,
                    items = invoice.Items?.Select(item => new
                    {
                        hsCode = item.HSCode,
                        productDescription = item.ItemName,
                        rate = $"{item.TaxRate}%",
                        uoM = item.UnitOfMeasure,
                        quantity = item.Quantity,
                        totalValues = item.TotalValue,
                        valueSalesExcludingST = item.TotalPrice,
                        fixedNotifiedValueOrRetailPrice = item.RetailPrice,
                        salesTaxApplicable = item.SalesTaxAmount,
                        extraTax = item.ExtraTax,
                        furtherTax = item.FurtherTax,
                        discount = item.Discount,
                        fedPayable = item.FedPayable,
                        salesTaxWithheldAtSource = item.SalesTaxWithheldAtSource,
                        saleType = "Goods at standard rate (default)"
                    }).ToList(),
                    totalSalesValue = invoice.Subtotal,
                    totalTaxCharged = invoice.TaxAmount,
                    totalInvoiceValue = invoice.TotalAmount
                };

                var json = JsonConvert.SerializeObject(fbrBody, Formatting.Indented);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {authToken}");

                var response = await _httpClient.PostAsync($"{_baseUrl}postinvoicedata_sb", content);
                var responseString = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = $"HTTP {response.StatusCode}: {responseString}",
                        ResponseData = responseString
                    };
                }

                // ✅ Parse the correct FBR response structure
                dynamic parsed = JsonConvert.DeserializeObject(responseString);

                string invoiceNumber = parsed?.invoiceNumber ?? "";
                string dated = parsed?.dated ?? "";
                string statusCode = parsed?.validationResponse?.statusCode ?? "";
                string status = parsed?.validationResponse?.status ?? "";
                string error = parsed?.validationResponse?.error ?? "";

                // Extract invoice number from the first item in invoiceStatuses array
                string fbrInvoiceNo = "";
                if (parsed?.validationResponse?.invoiceStatuses != null &&
                    parsed.validationResponse.invoiceStatuses.Count > 0)
                {
                    fbrInvoiceNo = parsed.validationResponse.invoiceStatuses[0].invoiceNo ?? "";
                }

                // Check if validation was successful
                bool isSuccess = statusCode == "00" && status == "Valid";

                return new FBRResponse
                {
                    Success = isSuccess,
                    IRN = fbrInvoiceNo, // This is the FBR-generated invoice number
                    ErrorMessage = isSuccess ? "" : error,
                    ResponseData = responseString
                };
            }
            catch (Exception ex)
            {
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"Exception: {ex.Message}",
                    ResponseData = ex.ToString()
                };
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}