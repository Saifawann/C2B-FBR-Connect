using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Http;
using Newtonsoft.Json;
using System.Threading.Tasks;
using C2B_FBR_Connect.Models;

namespace C2B_FBR_Connect.Services
{
    public class FBRApiService : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl = "https://gw.fbr.gov.pk/di_data/v1/di/"; // Replace with actual FBR API URL

        public FBRApiService()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
        }

        /// <summary>
        /// Uploads an invoice to the FBR Digital Invoicing system
        /// </summary>
        /// <param name="invoice">The invoice payload with all FBR required fields</param>
        /// <param name="authToken">FBR authentication token</param>
        /// <returns>FBR response with IRN if successful</returns>
        public async Task<FBRResponse> UploadInvoice(FBRInvoicePayload invoice, string authToken)
        {
            try
            {
                // Validate invoice data before sending
                if (string.IsNullOrEmpty(invoice.SellerNTN))
                {
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = "Seller NTN is required. Please configure it in Company Settings."
                    };
                }

                if (string.IsNullOrEmpty(invoice.InvoiceNumber))
                {
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = "Invoice number is required."
                    };
                }

                // Set authentication header
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {authToken}");

                // Serialize invoice to JSON
                var json = JsonConvert.SerializeObject(invoice, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    Formatting = Formatting.Indented
                });

                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // Send request to FBR API
                var response = await _httpClient.PostAsync($"{_baseUrl}postinvoicedata_sb", content);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    // Parse successful response
                    dynamic result = JsonConvert.DeserializeObject(responseContent);

                    return new FBRResponse
                    {
                        Success = true,
                        IRN = result?.invoiceNumber?.ToString() ?? "",
                        ResponseData = responseContent
                    };
                }
                else
                {
                    // Handle error response
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = $"FBR API Error {response.StatusCode}: {responseContent}",
                        ResponseData = responseContent
                    };
                }
            }
            catch (HttpRequestException ex)
            {
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"Network error: {ex.Message}"
                };
            }
            catch (TaskCanceledException ex)
            {
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = "Request timeout. Please try again."
                };
            }
            catch (JsonException ex)
            {
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"JSON error: {ex.Message}"
                };
            }
            catch (Exception ex)
            {
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"Unexpected error: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Verifies the connection to FBR API with the provided token
        /// </summary>
        /// <param name="authToken">FBR authentication token</param>
        /// <returns>True if connection is successful</returns>
        public async Task<FBRResponse> VerifyConnection(string authToken)
        {
            try
            {
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {authToken}");

                var response = await _httpClient.GetAsync($"{_baseUrl}verify");
                var responseContent = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    return new FBRResponse
                    {
                        Success = true,
                        ResponseData = responseContent
                    };
                }
                else
                {
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = $"Connection verification failed: {response.StatusCode}",
                        ResponseData = responseContent
                    };
                }
            }
            catch (Exception ex)
            {
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"Connection error: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Retrieves invoice status from FBR by IRN
        /// </summary>
        /// <param name="irn">Invoice Reference Number</param>
        /// <param name="authToken">FBR authentication token</param>
        /// <returns>FBR response with invoice status</returns>
        public async Task<FBRResponse> GetInvoiceStatus(string irn, string authToken)
        {
            try
            {
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {authToken}");

                var response = await _httpClient.GetAsync($"{_baseUrl}validateinvoicedata_sb?irn={irn}");
                var responseContent = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    return new FBRResponse
                    {
                        Success = true,
                        ResponseData = responseContent
                    };
                }
                else
                {
                    return new FBRResponse
                    {
                        Success = false,
                        ErrorMessage = $"Status check failed: {response.StatusCode}",
                        ResponseData = responseContent
                    };
                }
            }
            catch (Exception ex)
            {
                return new FBRResponse
                {
                    Success = false,
                    ErrorMessage = $"Error checking status: {ex.Message}"
                };
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}