using C2B_FBR_Connect.Models;
using iText.Kernel.Geom;
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
                var fbrBody = BuildFBRPayload(invoice);

                // ✅ Log request JSON
                var json = JsonConvert.SerializeObject(fbrBody, Formatting.Indented);
                System.Diagnostics.Debug.WriteLine("=== FBR Request ===");
                System.Diagnostics.Debug.WriteLine(json);

                // ✅ Prepare request
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken);

                // ✅ Make API call
                var response = await _httpClient.PostAsync($"{_baseUrl}postinvoicedata_sb", content);
                var responseString = await response.Content.ReadAsStringAsync();

                System.Diagnostics.Debug.WriteLine("=== FBR Response ===");
                System.Diagnostics.Debug.WriteLine(responseString);

                // ✅ Safely parse JSON response
                string fbrInvoiceNo = "";
                string? errorMessage = null;

                try
                {
                    dynamic? parsed = JsonConvert.DeserializeObject(responseString);
                    fbrInvoiceNo = parsed?.validationResponse?.invoiceStatuses?[0]?.invoiceNo?.ToString() ?? "";
                    errorMessage = parsed?.validationResponse?.message?.ToString() ?? responseString;
                }
                catch
                {
                    // If JSON parsing fails, just keep the raw response
                    errorMessage = responseString;
                }

                // ✅ Return consistent FBRResponse
                return new FBRResponse
                {
                    Success = response.IsSuccessStatusCode && !string.IsNullOrEmpty(fbrInvoiceNo),
                    IRN = fbrInvoiceNo,
                    ErrorMessage = errorMessage,
                    ResponseData = responseString
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("=== Exception ===");
                System.Diagnostics.Debug.WriteLine(ex);

                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    ResponseData = ex.ToString()
                };
            }
        }

        // ✅ Shared method to build FBR payload
        public object BuildFBRPayload(FBRInvoicePayload invoice)
        {
            string scenarioId = invoice.BuyerRegistrationType?.Equals("Registered", StringComparison.OrdinalIgnoreCase) == true
                ? "SN001" // Registered buyer
                : "SN002"; // Unregistered buyer

            return new
            {
                invoiceType = invoice.InvoiceType,
                invoiceDate = invoice.InvoiceDate.ToString("yyyy-MM-dd"),
                sellerNTNCNIC = invoice.SellerNTN,
                sellerBusinessName = invoice.SellerBusinessName,
                sellerProvince = invoice.SellerProvince,
                sellerAddress = invoice.SellerAddress,
                buyerNTNCNIC = invoice.CustomerNTN,
                buyerBusinessName = invoice.CustomerName,
                buyerProvince = invoice.BuyerProvince,
                buyerAddress = invoice.BuyerAddress,
                buyerRegistrationType = invoice.BuyerRegistrationType,
                invoiceRefNo = invoice.InvoiceNumber,
                scenarioId = scenarioId,
                items = invoice.Items?.Select(item => new
                {
                    hsCode = item.HSCode,
                    productDescription = item.ItemName,
                    rate = item.TaxRate > 0 ? $"{item.TaxRate:N0}%" : "0%",
                    uoM = item.UnitOfMeasure,
                    quantity = item.Quantity,
                    totalValues = item.TotalValue,
                    valueSalesExcludingST = item.TotalPrice,
                    fixedNotifiedValueOrRetailPrice = item.RetailPrice,
                    salesTaxApplicable = item.SalesTaxAmount,
                    extraTax = item.ExtraTax,
                    furtherTax = item.FurtherTax,
                    discount = 0.00,
                    fedPayable = 0.00,
                    salesTaxWithheldAtSource = 0.00,
                    saleType = "Goods at standard rate (default)"
                }).ToList()
            };
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}