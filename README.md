 The exception handler configured on ExceptionHandlerOptions produced a 404 status response. This InvalidOperationException containing the original exception was thrown since this is often due to a misconfigured ExceptionHandlingPath. If the exception handler is expected to return 404 status responses then set AllowStatusCode404Response to true.
 ---> System.IO.FileNotFoundException: Could not load file or assembly 'office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'. The system cannot find the file specified.
File name: 'office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'



warning NU1701: Package 'Microsoft.Office.Interop.Excel 15.0.4420.1017' was restored using '.NETFramework,Version=v4.6.1, .NETFramework,Version=v4.6.2, .NETFramework,Version=v4.7, .NETFramework,Version=v4.7.1, .NETFramework,Version=v4.7.2, .NETFramework,Version=v4.8, .NETFramework,Version=v4.8.1' instead of the project target framework 'net8.0'. This package may not be fully compatible with your project.



Application excelApp = new Application();
Workbook workbook = excelApp.Workbooks.Open(tempFilePath);
