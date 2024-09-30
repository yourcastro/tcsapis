using System;
using System.Net;
using System.Net.Http;
using RestSharp;

class Program
{
    static void Main(string[] args)
    {
        // Define proxy settings
        var proxy = new WebProxy("http://your-proxy-address:port")
        {
            Credentials = new NetworkCredential("username", "password") // If required
        };

        // Create an HttpClientHandler that uses the proxy
        var handler = new HttpClientHandler()
        {
            Proxy = proxy,
            UseProxy = true
        };

        // Create a RestClient with custom HttpClient
        var client = new RestClient(new RestClientOptions("https://api.example.com")
        {
            ConfigureHandler = httpClient => httpClient = new HttpClient(handler)
        });

        // Create a request
        var request = new RestRequest("endpoint", Method.Get);

        // Execute the request
        var response = client.Execute(request);

        // Check the response
        if (response.IsSuccessful)
        {
            Console.WriteLine(response.Content);
        }
        else
        {
            Console.WriteLine($"Error: {response.StatusCode} - {response.ErrorMessage}");
        }
    }
}
