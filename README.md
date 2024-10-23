 // Step 2: Write the file to a temporary location
 string tempFilePath = Path.Combine(Path.GetTempPath(), "tempDownloadedFile.xlsx");
 File.WriteAllBytes(tempFilePath, responseSource?.RawBytes ?? Array.Empty<byte>());

 // Step 3: Open the Excel file using OfficeOpenXml - Change to appropriate license if using commercially
 ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

 // Load the workbook from the file
 FileInfo fileInfo = new FileInfo(tempFilePath);
 var package = new ExcelPackage(fileInfo);
