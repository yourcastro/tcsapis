using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        var siteName = "your-site-name";
        var listName = "your-list-name";
        var itemId = "your-item-id"; // ID of the item
        var versionId = "your-version-id"; // ID of the version you want

        var accessToken = "your-access-token"; // Ensure you have a valid access token

        using (var client = new HttpClient())
        {
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // Retrieve specific version
            var url = $"https://your-sharepoint-site-url/sites/{siteName}/_api/web/lists/getbytitle('{listName}')/items({itemId})/versions({versionId})/$value";

            var response = await client.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsByteArrayAsync();
                System.IO.File.WriteAllBytes("YourFileName.xlsx", content);
                Console.WriteLine("File downloaded successfully.");
            }
            else
            {
                Console.WriteLine($"Error: {response.StatusCode}");
            }
        }
    }
}
