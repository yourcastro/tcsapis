using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;

public class ApiClient
{
    private static readonly HttpClientHandler httpClientHandler = new HttpClientHandler
    {
        Credentials = CredentialCache.DefaultNetworkCredentials // Use default credentials
    };

    private static readonly HttpClient httpClient = new HttpClient(httpClientHandler);

    public async Task<string> PostDataAsync(string url, object data)
    {
        try
        {
            // Send a POST request with the JSON body
            HttpResponseMessage response = await httpClient.PostAsJsonAsync(url, data);

            // Ensure the response is successful
            response.EnsureSuccessStatusCode();

            // Read the response content as a string
            var responseData = await response.Content.ReadAsStringAsync();
            return responseData; // Return the response data
        }
        catch (HttpRequestException e)
        {
            Console.WriteLine("Error calling API: " + e.Message);
            return null; // Or handle error as needed
        }
    }
}
