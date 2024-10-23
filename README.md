 The exception handler configured on ExceptionHandlerOptions produced a 404 status response. This InvalidOperationException containing the original exception was thrown since this is often due to a misconfigured ExceptionHandlingPath. If the exception handler is expected to return 404 status responses then set AllowStatusCode404Response to true.
 ---> System.IO.FileNotFoundException: Could not load file or assembly 'office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'. The system cannot find the file specified.
File name: 'office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'



warning NU1701: Package 'Microsoft.Office.Interop.Excel 15.0.4420.1017' was restored using '.NETFramework,Version=v4.6.1, .NETFramework,Version=v4.6.2, .NETFramework,Version=v4.7, .NETFramework,Version=v4.7.1, .NETFramework,Version=v4.7.2, .NETFramework,Version=v4.8, .NETFramework,Version=v4.8.1' instead of the project target framework 'net8.0'. This package may not be fully compatible with your project.



Application excelApp = new Application();
Workbook workbook = excelApp.Workbooks.Open(tempFilePath);


// Set the EPPlus license context (required by EPPlus 5+)
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  // Change to appropriate license if using commercially

    // Load the workbook from the file
    using (var package = new ExcelPackage(new FileInfo(tempFilePath)))
    {
        // Access the first worksheet (for example)
        var workbook = package.Workbook;
        var worksheet = workbook.Worksheets[0];  // Access the first worksheet

        // You can now work with the worksheet
        var cellValue = worksheet.Cells[1, 1].Text;  // Get value from cell A1
        // Continue with your logic
    }
