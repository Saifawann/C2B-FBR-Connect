using C2B_FBR_Connect.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace C2B_FBR_Connect.Services
{
    public class FBRApiService : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly HttpClientHandler _httpClientHandler;
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

        private string GetBaseUrl(string environment)
        {
            if (environment?.Equals("Sandbox", StringComparison.OrdinalIgnoreCase) == true)
            {
                return "https://gw.fbr.gov.pk/di_data/v1/di/postinvoicedata_sb";
            }

            return "https://gw.fbr.gov.pk/di_data/v1/di/postinvoicedata";
        }

        public async Task<List<TransactionType>> FetchTransactionTypesAsync(string authToken = null)
        {
            var transactionTypes = new List<TransactionType>();

            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, "https://gw.fbr.gov.pk/pdi/v1/transtypecode");

                if (!string.IsNullOrWhiteSpace(authToken))
                {
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());
                }

                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.TryAddWithoutValidation("User-Agent", "C2B_FBR_Connect/1.0");

                var response = await _httpClient.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    var jsonString = await response.Content.ReadAsStringAsync();
                    var options = new System.Text.Json.JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    };

                    var apiResponse = System.Text.Json.JsonSerializer.Deserialize<List<TransactionTypeResponse>>(jsonString, options);

                    if (apiResponse != null)
                    {
                        foreach (var item in apiResponse)
                        {
                            transactionTypes.Add(new TransactionType
                            {
                                TransactionTypeId = item.TransactioN_TYPE_ID,
                                TransactionDesc = item.TransactioN_DESC?.Trim() ?? "",
                                LastUpdated = DateTime.Now
                            });
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"✅ Fetched {transactionTypes.Count} transaction types from FBR");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ Failed to fetch transaction types: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error fetching transaction types: {ex.Message}");
            }

            return transactionTypes;
        }

        public async Task<List<Province>> FetchProvincesAsync(string authToken = null)
        {
            var provinces = new List<Province>();

            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, "https://gw.fbr.gov.pk/pdi/v1/provinces");

                if (!string.IsNullOrWhiteSpace(authToken))
                {
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());
                }

                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.TryAddWithoutValidation("User-Agent", "C2B_FBR_Connect/1.0");

                var response = await _httpClient.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    var jsonString = await response.Content.ReadAsStringAsync();
                    var options = new System.Text.Json.JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    };

                    var apiResponse = System.Text.Json.JsonSerializer.Deserialize<List<ProvinceResponse>>(jsonString, options);

                    if (apiResponse != null)
                    {
                        foreach (var item in apiResponse)
                        {
                            provinces.Add(new Province
                            {
                                StateProvinceCode = item.StateProvinceCode,
                                StateProvinceDesc = item.StateProvinceDesc
                            });
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"✅ Fetched {provinces.Count} provinces from FBR");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ Failed to fetch provinces: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error fetching provinces: {ex.Message}");
            }

            return provinces;
        }

        public async Task<FBRResponse> UploadInvoice(FBRInvoicePayload invoice, string authToken, string environment = "Production")
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

                // Build and enrich the payload
                var enrichedPayload = BuildFBRPayload(invoice);

                // Convert to API format with correct property names
                // ✅ PASS environment parameter
                var apiPayload = ConvertToApiPayload(enrichedPayload, environment);

                // ✅ VALIDATE before sending
                var validationErrors = ValidateFBRPayload(apiPayload);
                if (validationErrors.Count > 0)
                {
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = "Validation failed: " + string.Join("; ", validationErrors),
                        ResponseData = string.Join("\n", validationErrors)
                    };
                }

                // Serialize to JSON
                var json = JsonConvert.SerializeObject(apiPayload, Formatting.Indented);

                Console.WriteLine($"\n📤 Sending to FBR API ({environment}):");
                Console.WriteLine($"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                Console.WriteLine(json);
                Console.WriteLine($"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

                // Get the correct URL based on environment
                string apiUrl = GetBaseUrl(environment);

                var request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
                {
                    Content = new StringContent(json, Encoding.UTF8, "application/json")
                };

                request.Headers.TryAddWithoutValidation("User-Agent", "PostmanRuntime/7.26.8");
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());

                System.Diagnostics.Debug.WriteLine($"🌍 Environment: {environment}");
                System.Diagnostics.Debug.WriteLine($"📍 API URL: {apiUrl}");
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

        public async Task<string> GetUOMDescriptionAsync(string hsCode, string authToken)
        {
            if (string.IsNullOrWhiteSpace(authToken))
                return "Error: Missing authentication token";

            string url = $"https://gw.fbr.gov.pk/pdi/v2/HS_UOM?hs_code={hsCode}&annexure_id=3";

            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.TryAddWithoutValidation("User-Agent", "C2B_FBR_Connect/1.0");

                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    return $"Error: {response.StatusCode} - {response.ReasonPhrase}\n{json}";
                }

                var options = new System.Text.Json.JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                var items = System.Text.Json.JsonSerializer.Deserialize<List<FbrUomItem>>(json, options);

                if (items != null && items.Any())
                    return items.First().Description;

                return "Description not found";
            }
            catch (HttpRequestException ex)
            {
                return $"Network error: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        public async Task<List<SaleTypeRate>> FetchSaleTypeRatesAsync(string date, int transTypeId, int provinceId, string authToken = null)
        {
            var rates = new List<SaleTypeRate>();

            if (string.IsNullOrWhiteSpace(authToken))
            {
                System.Diagnostics.Debug.WriteLine("⚠️ Auth token is missing for FetchSaleTypeRatesAsync");
                return rates;
            }

            string url = $"https://gw.fbr.gov.pk/pdi/v2/SaleTypeToRate?date={date}&transTypeId={transTypeId}&originationSupplier={provinceId}";

            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.TryAddWithoutValidation("User-Agent", "C2B_FBR_Connect/1.0");

                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"❌ FetchSaleTypeRates failed: {response.StatusCode} - {json}");
                    return rates;
                }

                var options = new System.Text.Json.JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                rates = System.Text.Json.JsonSerializer.Deserialize<List<SaleTypeRate>>(json, options) ?? new List<SaleTypeRate>();
                System.Diagnostics.Debug.WriteLine($"✅ Fetched {rates.Count} rates for transTypeId={transTypeId}, province={provinceId}");

                return rates;
            }
            catch (HttpRequestException ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Network error in FetchSaleTypeRates: {ex.Message}");
                return rates;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error in FetchSaleTypeRates: {ex.Message}");
                return rates;
            }
        }

        public async Task<SroSchedule> FetchSroScheduleAsync(int rateId, string date, int provinceId, string authToken = null)
        {
            if (string.IsNullOrWhiteSpace(authToken))
            {
                System.Diagnostics.Debug.WriteLine("⚠️ Auth token is missing for FetchSroScheduleAsync");
                return null;
            }

            string url = $"https://gw.fbr.gov.pk/pdi/v1/SroSchedule?rate_id={rateId}&date={date}&origination_supplier_csv={provinceId}";

            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.TryAddWithoutValidation("User-Agent", "C2B_FBR_Connect/1.0");

                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"❌ FetchSroSchedule failed: {response.StatusCode} - {json}");
                    return null;
                }

                var options = new System.Text.Json.JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                var schedules = System.Text.Json.JsonSerializer.Deserialize<List<SroSchedule>>(json, options);

                if (schedules != null && schedules.Count > 0)
                {
                    var schedule = schedules.First();
                    System.Diagnostics.Debug.WriteLine($"✅ Fetched SRO Schedule: {schedule.SRO_DESC} (ID: {schedule.SRO_ID})");
                    return schedule;
                }

                System.Diagnostics.Debug.WriteLine($"⚠️ No SRO Schedule found for rateId={rateId}");
                return null;
            }
            catch (HttpRequestException ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Network error in FetchSroSchedule: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error in FetchSroSchedule: {ex.Message}");
                return null;
            }
        }

        public async Task<string> FetchSroItemSerialNoAsync(int sroId, string date, string authToken = null)
        {
            if (string.IsNullOrWhiteSpace(authToken))
            {
                System.Diagnostics.Debug.WriteLine("⚠️ Auth token is missing for FetchSroItemSerialNoAsync");
                return null;
            }

            string formattedDate = date;
            try
            {
                DateTime parsedDate = DateTime.ParseExact(date, "dd-MMMyyyy", System.Globalization.CultureInfo.InvariantCulture);
                formattedDate = parsedDate.ToString("yyyy-MM-dd");
                System.Diagnostics.Debug.WriteLine($"📅 Date converted: {date} → {formattedDate}");
            }
            catch (FormatException ex)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Date parsing failed for '{date}': {ex.Message}");
                if (DateTime.TryParse(date, out DateTime fallbackDate))
                {
                    formattedDate = fallbackDate.ToString("yyyy-MM-dd");
                    System.Diagnostics.Debug.WriteLine($"📅 Date converted (fallback): {date} → {formattedDate}");
                }
            }

            string url = $"https://gw.fbr.gov.pk/pdi/v2/SROItem?date={formattedDate}&sro_id={sroId}";

            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                System.Diagnostics.Debug.WriteLine($"📤 Request URL: {url}");

                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"❌ FetchSroItemSerialNo failed: {response.StatusCode} - {json}");
                    return null;
                }

                var options = new System.Text.Json.JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                var items = System.Text.Json.JsonSerializer.Deserialize<List<SroItem>>(json, options);

                if (items != null && items.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"📋 Found {items.Count} SRO item(s) for sroId={sroId}");

                    foreach (var item in items)
                    {
                        if (!string.IsNullOrWhiteSpace(item.SRO_ITEM_DESC) &&
                            item.SRO_ITEM_DESC != "0" &&
                            item.SRO_ITEM_DESC.Trim().Length > 0)
                        {
                            string serialNo = item.SRO_ITEM_DESC.Trim();
                            System.Diagnostics.Debug.WriteLine($"✅ Selected SRO Item Serial No: {serialNo}");
                            return serialNo;
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"⚠️ All {items.Count} SRO items have empty descriptions");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"⚠️ No SRO Items found for sroId={sroId}");
                return null;
            }
            catch (HttpRequestException ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Network error in FetchSroItemSerialNo: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error in FetchSroItemSerialNo: {ex.Message}");
                return null;
            }
        }

        public async Task<List<SroSchedule>> FetchAllSroSchedulesAsync(int rateId, string date, int provinceId, string authToken = null)
        {
            if (string.IsNullOrWhiteSpace(authToken))
            {
                System.Diagnostics.Debug.WriteLine("⚠️ Auth token is missing for FetchAllSroSchedulesAsync");
                return new List<SroSchedule>();
            }

            string url = $"https://gw.fbr.gov.pk/pdi/v1/SroSchedule?rate_id={rateId}&date={date}&origination_supplier_csv={provinceId}";

            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authToken.Trim());
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.TryAddWithoutValidation("User-Agent", "C2B_FBR_Connect/1.0");

                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"❌ FetchAllSroSchedules failed: {response.StatusCode} - {json}");
                    return new List<SroSchedule>();
                }

                var options = new System.Text.Json.JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                var schedules = System.Text.Json.JsonSerializer.Deserialize<List<SroSchedule>>(json, options);

                if (schedules != null && schedules.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"📋 Found {schedules.Count} SRO Schedule(s) for rateId={rateId}");

                    foreach (var schedule in schedules)
                    {
                        System.Diagnostics.Debug.WriteLine($"   - {schedule.SRO_DESC} (ID: {schedule.SRO_ID})");
                    }

                    return schedules;
                }

                System.Diagnostics.Debug.WriteLine($"⚠️ No SRO Schedule found for rateId={rateId}");
                return new List<SroSchedule>();
            }
            catch (HttpRequestException ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Network error in FetchAllSroSchedules: {ex.Message}");
                return new List<SroSchedule>();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error in FetchAllSroSchedules: {ex.Message}");
                return new List<SroSchedule>();
            }
        }

        private FBRResponse ParseFBRResponse(string responseString)
        {
            try
            {
                var jsonResponse = JObject.Parse(responseString);

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

                var topLevelInvoiceNumber = jsonResponse["invoiceNumber"]?.ToString();

                var validationResponse = jsonResponse["validationResponse"];
                if (validationResponse != null)
                {
                    var statusCode = validationResponse["statusCode"]?.ToString();
                    var status = validationResponse["status"]?.ToString();
                    var topLevelErrorCode = validationResponse["errorCode"]?.ToString();
                    var topLevelError = validationResponse["error"]?.ToString();

                    string invoiceNo = topLevelInvoiceNumber ?? "";
                    string detailedError = "";
                    var invoiceStatuses = validationResponse["invoiceStatuses"];

                    if (invoiceStatuses != null && invoiceStatuses.Any())
                    {
                        var firstItem = invoiceStatuses[0];

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
            catch (Newtonsoft.Json.JsonException ex)
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

        public FBRInvoicePayload BuildFBRPayload(FBRInvoicePayload details)
        {
            string scenarioId = DetermineScenarioFromItems(details.Items, details.BuyerRegistrationType);
            details.ScenarioId = scenarioId;

            foreach (var item in details.Items)
            {
                ScenarioMapper.AutoFillSroDefaults(item);

                if (string.IsNullOrEmpty(item.SaleType))
                {
                    item.SaleType = "Goods at Standard Rate (default)";
                }
            }

            return details;
        }

        private string DetermineScenarioFromItems(List<InvoiceItem> items, string buyerRegistrationType)
        {
            if (items == null || items.Count == 0)
                return "SN001";

            var scenarios = items
                .Where(i => !string.IsNullOrEmpty(i.SaleType))
                .Select(i => ScenarioMapper.DetermineScenarioId(i.SaleType, buyerRegistrationType))
                .Distinct()
                .ToList();

            if (scenarios.Count == 0)
                return "SN001";

            var specialScenario = scenarios.FirstOrDefault(s => s != "SN001" && s != "SN002");
            if (specialScenario != null)
                return specialScenario;

            return scenarios.First();
        }

        public FBRApiPayload ConvertToApiPayload(FBRInvoicePayload details, string environment = "Production")
        {
            Console.WriteLine($"\n🔄 Converting to FBR API Payload format...");
            Console.WriteLine($"   Invoice: {details.InvoiceNumber}");
            Console.WriteLine($"   Items: {details.Items?.Count ?? 0}");
            Console.WriteLine($"   Environment: {environment}");

            var apiPayload = new FBRApiPayload
            {
                InvoiceType = details.InvoiceType == "Sale Invoice" ? "Sale Invoice" : "Debit Note",
                InvoiceDate = details.InvoiceDate.ToString("yyyy-MM-dd"),
                SellerBusinessName = details.SellerBusinessName ?? "",
                SellerProvince = details.SellerProvince ?? "",
                SellerNTNCNIC = details.SellerNTN ?? "",
                SellerAddress = details.SellerAddress ?? "",
                BuyerNTNCNIC = details.CustomerNTN ?? "",
                BuyerBusinessName = details.CustomerName ?? "",
                BuyerProvince = details.BuyerProvince ?? "",
                BuyerAddress = details.BuyerAddress ?? "",
                InvoiceRefNo = "",
                BuyerRegistrationType = details.BuyerRegistrationType ?? "Registered",
                Items = new List<FBRApiItem>()
            };

            // ✅ ONLY add ScenarioId if environment is Sandbox
            if (environment?.Equals("Sandbox", StringComparison.OrdinalIgnoreCase) == true)
            {
                apiPayload.ScenarioId = details.ScenarioId ?? "SN001";
                Console.WriteLine($"   📋 ScenarioId included: {apiPayload.ScenarioId} (Sandbox mode)");
            }
            else
            {
                Console.WriteLine($"   📋 ScenarioId excluded (Production mode)");
            }

            if (details.Items != null)
            {
                Console.WriteLine($"\n   📦 Converting {details.Items.Count} items:");

                foreach (var item in details.Items)
                {
                    decimal valueSalesExcludingST = item.NetAmount;

                    if (valueSalesExcludingST <= 0)
                    {
                        Console.WriteLine($"      ⚠️ WARNING: Item '{item.ItemName}' has invalid NetAmount={item.NetAmount}");
                        Console.WriteLine($"         Using TotalPrice={item.TotalPrice} as fallback");
                        valueSalesExcludingST = Math.Abs(item.TotalPrice);
                    }

                    decimal totalValues = item.TotalValue;

                    Console.WriteLine($"      ✅ {item.ItemName}");
                    Console.WriteLine($"         valueSalesExcludingST: {valueSalesExcludingST}");
                    Console.WriteLine($"         totalValues: {totalValues}");
                    Console.WriteLine($"         quantity: {item.Quantity}");
                    Console.WriteLine($"         rate: {item.TaxRate}%");
                    Console.WriteLine($"         saleType: {item.SaleType}");

                    var apiItem = new FBRApiItem
                    {
                        HsCode = item.HSCode ?? "",
                        ProductDescription = item.ItemName ?? "",
                        Rate = item.Rate,
                        UoM = item.UnitOfMeasure ?? "",
                        Quantity = item.Quantity,
                        ValueSalesExcludingST = valueSalesExcludingST,
                        TotalValues = totalValues,
                        FixedNotifiedValueOrRetailPrice = item.RetailPrice,
                        SalesTaxApplicable = item.SalesTaxAmount,
                        SalesTaxWithheldAtSource = item.SalesTaxWithheldAtSource,
                        ExtraTax = item.ExtraTax > 0 ? item.ExtraTax.ToString() : "",
                        FurtherTax = item.FurtherTax,
                        SroScheduleNo = item.SroScheduleNo ?? "",
                        FedPayable = item.FedPayable,
                        Discount = item.Discount,
                        SaleType = ScenarioMapper.NormalizeSaleType(item.SaleType) ?? "Goods at standard rate (default)",
                        SroItemSerialNo = item.SroItemSerialNo ?? ""
                    };

                    apiPayload.Items.Add(apiItem);
                }
            }

            Console.WriteLine($"✅ Conversion complete\n");
            return apiPayload;
        }

        public List<string> ValidateFBRPayload(FBRApiPayload payload)
        {
            var errors = new List<string>();

            Console.WriteLine($"\n🔍 Validating FBR Payload...");

            if (string.IsNullOrEmpty(payload.SellerNTNCNIC))
                errors.Add("Seller NTN is required");

            if (payload.Items == null || payload.Items.Count == 0)
                errors.Add("At least one item is required");

            if (payload.Items != null)
            {
                for (int i = 0; i < payload.Items.Count; i++)
                {
                    var item = payload.Items[i];
                    string itemPrefix = $"Item {i + 1} ({item.ProductDescription})";

                    if (item.Quantity <= 0)
                        errors.Add($"{itemPrefix}: Quantity must be greater than 0");

                    if (item.ValueSalesExcludingST <= 0)
                        errors.Add($"{itemPrefix}: valueSalesExcludingST must be greater than 0 (currently: {item.ValueSalesExcludingST})");

                    if (item.TotalValues <= 0)
                        errors.Add($"{itemPrefix}: totalValues must be greater than 0 (currently: {item.TotalValues})");

                    if (string.IsNullOrEmpty(item.SaleType))
                        errors.Add($"{itemPrefix}: Sale Type is required");

                    decimal taxRate = 18m;
                    if (!string.IsNullOrEmpty(item.Rate))
                    {
                        string rateStr = item.Rate.Replace("%", "").Trim();
                        decimal.TryParse(rateStr, out taxRate);
                    }

                    String[] ignoredSaleTypes = ["Processing/Conversion of Goods", "Goods (FED in ST Mode)", "Telecommunication services", "Services (FED in ST Mode)"];

                    if (taxRate != 18m && !ignoredSaleTypes.Contains(item.SaleType))
                    {
                        if (string.IsNullOrWhiteSpace(item.SroScheduleNo))
                            errors.Add($"{itemPrefix}: SRO Schedule Number is required for rate {taxRate}% (Error 0077)");

                        if (string.IsNullOrWhiteSpace(item.SroItemSerialNo))
                            errors.Add($"{itemPrefix}: SRO Item Serial Number is required for rate {taxRate}% (Error 0077)");
                    }

                    bool hasSchedule = !string.IsNullOrWhiteSpace(item.SroScheduleNo);
                    bool hasSerial = !string.IsNullOrWhiteSpace(item.SroItemSerialNo);

                    if (hasSchedule && !hasSerial)
                        errors.Add($"{itemPrefix}: SRO Schedule provided but Serial Number missing (Error 0078)");

                    if (!hasSchedule && hasSerial)
                        errors.Add($"{itemPrefix}: SRO Serial Number provided but Schedule missing");
                }
            }

            if (errors.Count > 0)
            {
                Console.WriteLine($"❌ Validation failed with {errors.Count} error(s):");
                foreach (var error in errors)
                {
                    Console.WriteLine($"   - {error}");
                }
            }
            else
            {
                Console.WriteLine($"✅ Validation passed");
            }

            return errors;
        }

        private class FbrUomItem
        {
            public int UoM_ID { get; set; }
            public string Description { get; set; } = string.Empty;
        }

        private class TransactionTypeResponse
        {
            public int TransactioN_TYPE_ID { get; set; }
            public string TransactioN_DESC { get; set; }
        }

        private class ProvinceResponse
        {
            public int StateProvinceCode { get; set; }
            public string StateProvinceDesc { get; set; }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
            _httpClientHandler?.Dispose();
        }
    }
}