using OfficeOpenXml;
using System.IO;

// Step 1: Assume `responseSource.RawBytes` contains the raw Excel file bytes
byte[] excelFileBytes = responseSource?.RawBytes ?? Array.Empty<byte>();

// Step 2: Use MemoryStream to avoid saving to a temporary location
using (MemoryStream memoryStream = new MemoryStream(excelFileBytes))
{
    // Step 3: Set the LicenseContext for non-commercial use
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    // Step 4: Load the Excel package from the memory stream
    using (ExcelPackage package = new ExcelPackage(memoryStream))
    {
        // Now you have the ExcelPackage object ready to use
        // You can access the workbook, worksheets, and manipulate the data
        var workbook = package.Workbook;

        // Example: Access the first worksheet
        var worksheet = workbook.Worksheets[0];

        // Do whatever you need with the Excel data
        // For example, read data from the first cell in the first worksheet
        string firstCellValue = worksheet.Cells[1, 1].Text;

        // Return the package, or perform operations on it
        return package; // Or work with the data here
    }
}
