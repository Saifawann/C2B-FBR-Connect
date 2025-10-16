using C2B_FBR_Connect.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
        private readonly HttpClientHandler _httpClientHandler;
        private readonly string _baseUrl = "https://gw.fbr.gov.pk/di_data/v1/di/";

        public FBRApiService()
        {
            _httpClientHandler = new HttpClientHandler
            {
                UseCookies = true,
                AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate,
                CookieContainer = new System.Net.CookieContainer()
            };

            _httpClient = new HttpClient(_httpClientHandler);
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
        }

        public async Task<FBRResponse> UploadInvoice(FBRInvoicePayload invoice, string authToken)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(authToken))
                {
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = "Authentication token is required",
                        ResponseData = "No token provided"
                    };
                }

                var fbrBody = BuildFBRPayload(invoice);
                var json = JsonConvert.SerializeObject(fbrBody, Formatting.Indented);

                var request = new HttpRequestMessage(HttpMethod.Post, $"{_baseUrl}postinvoicedata_sb")
                {
                    Content = new StringContent(json, Encoding.UTF8, "application/json")
                };

                request.Headers.TryAddWithoutValidation("User-Agent", "PostmanRuntime/7.26.8");
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());

                System.Diagnostics.Debug.WriteLine($"🔹 Sending invoice request to FBR...");

                var startTime = DateTime.Now;
                var response = await _httpClient.SendAsync(request);
                var responseString = await response.Content.ReadAsStringAsync();
                var duration = DateTime.Now - startTime;

                System.Diagnostics.Debug.WriteLine($"✅ HTTP {(int)response.StatusCode} {response.StatusCode} ({duration.TotalMilliseconds:F0} ms)");

                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"❌ Request failed with status {response.StatusCode}");
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = $"HTTP {(int)response.StatusCode}: {response.ReasonPhrase}",
                        ResponseData = responseString
                    };
                }

                if (string.IsNullOrWhiteSpace(responseString))
                {
                    System.Diagnostics.Debug.WriteLine("⚠️ Response body is empty");
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = "Empty response from FBR API",
                        ResponseData = "Empty Response"
                    };
                }

                return ParseFBRResponse(responseString);
            }
            catch (HttpRequestException ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ HTTP Error: {ex.Message}");
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"Network Error: {ex.Message}",
                    ResponseData = ex.ToString()
                };
            }
            catch (TaskCanceledException)
            {
                System.Diagnostics.Debug.WriteLine("❌ Request timed out");
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = "Request timed out after 30 seconds",
                    ResponseData = "Timeout"
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Unexpected error: {ex.Message}");
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"Unexpected error: {ex.Message}",
                    ResponseData = ex.ToString()
                };
            }
        }

        private FBRResponse ParseFBRResponse(string responseString)
        {
            try
            {
                var jsonResponse = JObject.Parse(responseString);

                // Check for authentication fault
                var fault = jsonResponse["fault"];
                if (fault != null)
                {
                    var faultMessage = fault["message"]?.ToString();
                    var faultDescription = fault["description"]?.ToString();

                    System.Diagnostics.Debug.WriteLine($"❌ Authentication fault: {faultMessage}");

                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = $"Authentication Failed: {faultMessage} - {faultDescription}",
                        ResponseData = responseString
                    };
                }

                // ✅ Check for top-level invoice number (success response)
                var topLevelInvoiceNumber = jsonResponse["invoiceNumber"]?.ToString();

                // Check for validation response
                var validationResponse = jsonResponse["validationResponse"];
                if (validationResponse != null)
                {
                    var statusCode = validationResponse["statusCode"]?.ToString();
                    var status = validationResponse["status"]?.ToString();
                    var topLevelErrorCode = validationResponse["errorCode"]?.ToString();
                    var topLevelError = validationResponse["error"]?.ToString();

                    string invoiceNo = topLevelInvoiceNumber ?? ""; // ✅ Use top-level invoice number first
                    string detailedError = "";
                    var invoiceStatuses = validationResponse["invoiceStatuses"];

                    if (invoiceStatuses != null && invoiceStatuses.Any())
                    {
                        var firstItem = invoiceStatuses[0];

                        // ✅ Only use item-level invoice number if top-level is missing
                        if (string.IsNullOrEmpty(invoiceNo))
                        {
                            invoiceNo = firstItem?["invoiceNo"]?.ToString() ?? "";
                        }

                        var itemErrorCode = firstItem?["errorCode"]?.ToString();
                        var itemError = firstItem?["error"]?.ToString();

                        if (!string.IsNullOrEmpty(itemError))
                        {
                            detailedError = $"Error Code {itemErrorCode}: {itemError}";
                        }
                    }

                    bool isSuccess = statusCode == "00" && !string.IsNullOrEmpty(invoiceNo);

                    if (isSuccess)
                    {
                        System.Diagnostics.Debug.WriteLine($"✅ Invoice submitted successfully! IRN: {invoiceNo}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"❌ Validation failed: {status} (Code: {statusCode})");
                    }

                    string errorMessage = null;
                    if (!isSuccess)
                    {
                        if (!string.IsNullOrEmpty(detailedError))
                        {
                            errorMessage = detailedError;
                        }
                        else if (!string.IsNullOrEmpty(topLevelError))
                        {
                            errorMessage = $"Error Code {topLevelErrorCode}: {topLevelError} (Status: {status})";
                        }
                        else
                        {
                            errorMessage = $"Status: {status}, Code: {statusCode}";
                        }
                    }

                    return new FBRResponse
                    {
                        Success = isSuccess,
                        IRN = invoiceNo,
                        ErrorMessage = errorMessage,
                        ResponseData = responseString
                    };
                }

                // Check for direct error format
                var code = jsonResponse["Code"]?.ToString();
                var errorMsg = jsonResponse["error"]?.ToString();

                if (!string.IsNullOrEmpty(code) && code != "00")
                {
                    System.Diagnostics.Debug.WriteLine($"❌ Error response: Code {code}");
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = $"Code: {code}, Error: {errorMsg}",
                        ResponseData = responseString
                    };
                }

                System.Diagnostics.Debug.WriteLine("⚠️ Unexpected response format");
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = "Unexpected response format from FBR API",
                    ResponseData = responseString
                };
            }
            catch (JsonException ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ JSON parsing error: {ex.Message}");
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"Invalid JSON response: {ex.Message}",
                    ResponseData = responseString
                };
            }
        }

        public object BuildFBRPayload(FBRInvoicePayload invoice)
        {
            string scenarioId = invoice.BuyerRegistrationType?.Equals("Registered", StringComparison.OrdinalIgnoreCase) == true
                ? "SN001"
                : "SN002";

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
                    rate = $"{item.TaxRate}%",
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
            _httpClientHandler?.Dispose();
        }
    }
}