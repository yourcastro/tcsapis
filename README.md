using System;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;

public class ApiClient
{
    private static readonly HttpClient httpClient = new HttpClient();

    public async Task PostDataAsync(string url, object data)
    {
        try
        {
            // Send a POST request with the JSON body
            HttpResponseMessage response = await httpClient.PostAsJsonAsync(url, data);

            // Ensure the response is successful
            response.EnsureSuccessStatusCode();

            // Optionally read the response content
            var responseData = await response.Content.ReadAsStringAsync();
            Console.WriteLine("Response from API: " + responseData);
        }
        catch (HttpRequestException e)
        {
            Console.WriteLine("Error calling API: " + e.Message);
        }
    }
}
