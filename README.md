using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

class Program
{
    private static readonly string siteUrl = "https://yourtenant.sharepoint.com/sites/yoursite";
    private static readonly string sourceLibrary = "SourceLibraryName";
    private static readonly string destinationLibrary = "DestinationLibraryName";
    private static readonly string fileName = "sourceFile.txt";
    private static readonly string newFileName = "newFileName.txt";
    private static readonly string accessToken = "YOUR_ACCESS_TOKEN"; // Obtain this via OAuth

    static async Task Main(string[] args)
    {
        await CopyFileAsync();
    }

    private static async Task CopyFileAsync()
    {
        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // Step 1: Get the file from the source library
            string sourceFileUrl = $"{siteUrl}/_api/web/GetFolderByServerRelativeUrl('{sourceLibrary}')/Files('{fileName}')/$value";
            HttpResponseMessage response = await client.GetAsync(sourceFileUrl);

            if (response.IsSuccessStatusCode)
            {
                byte[] fileContent = await response.Content.ReadAsByteArrayAsync();

                // Step 2: Upload the file to the destination library
                string destinationFileUrl = $"{siteUrl}/_api/web/GetFolderByServerRelativeUrl('{destinationLibrary}')/Files/Add(url='{newFileName}',overwrite=true)";
                using (var content = new ByteArrayContent(fileContent))
                {
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    HttpResponseMessage uploadResponse = await client.PostAsync(destinationFileUrl, content);

                    if (uploadResponse.IsSuccessStatusCode)
                    {
                        Console.WriteLine("File copied successfully.");
                    }
                    else
                    {
                        Console.WriteLine($"Error uploading file: {uploadResponse.ReasonPhrase}");
                    }
                }
            }
            else
            {
                Console.WriteLine($"Error fetching file: {response.ReasonPhrase}");
            }
        }
    }
}
