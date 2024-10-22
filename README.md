using Microsoft.Office.Interop.Excel;
using RestSharp;
using System;
using System.IO;
using System.Threading.Tasks;

public class SharePointClient
{
    private string sharepointSiteUrl = "https://<your-tenant>.sharepoint.com/sites/<site-name>";
    private string accessToken = "<your-access-token>";  // OAuth token for authentication

    public SharePointClient()
    {
        // You can implement logic here to get an access token via OAuth2 if required
    }

    public async Task<Workbook> GetFileVersionAndLoadIntoExcelAsync(string fileRelativePath, string versionLabel)
    {
        var client = new RestClient(sharepointSiteUrl);

        // Step 1: Download the specific version of the file
        var versionDownloadRequest = new RestRequest($"{fileRelativePath}/versions/{versionLabel}/$value", Method.Get);
        versionDownloadRequest.AddHeader("Authorization", "Bearer " + accessToken);

        var versionResponse = await client.ExecuteAsync(versionDownloadRequest);

        if (!versionResponse.IsSuccessful)
        {
            Console.WriteLine("Failed to download specific file version.");
            return null;
        }

        // Step 2: Write the file to a temporary location
        string tempFilePath = Path.Combine(Path.GetTempPath(), "tempDownloadedFile.xlsx");
        File.WriteAllBytes(tempFilePath, versionResponse.RawBytes);

        // Step 3: Open the Excel file using Microsoft.Office.Interop.Excel
        Application excelApp = new Application();
        Workbook workbook = excelApp.Workbooks.Open(tempFilePath);

        return workbook;
    }
}
