public string SaveScoreCardFileData(int EntityScorecardID, string FilePath, string TemplateName, string userID)
{
    string xmlstring = "";
    string fileName = Path.GetFileName(FilePath);
    string result = "";
    string TranslationTypeId = string.Empty;
    string Sessionkey = CreateSessionId();
    string excelPwd = string.Empty;

    try
    {
        // Get the translation type ID using your existing function
        ClsPDScoreCardFunctions.GetScoredCardTranslationTypeId(
            _appSettings.Application_Name,
            _appSettings.SQLUserName,
            _appSettings.SQLPassword,
            _appSettings.CR_DB_Operational,
            Sessionkey, TemplateName, ref TranslationTypeId);

        excelPwd = _appSettings.WorkBookPwd ?? string.Empty;

        // Load the Excel file using EPPlus
        var fileInfo = new FileInfo(FilePath);
        using (var package = new ExcelPackage(fileInfo, new OfficeOpenXml.ExcelProtectionProvider(excelPwd)))
        {
            // Access the workbook
            ExcelWorkbook workbook = package.Workbook;
            
            // Perform your operations here - For example, generating XML from the workbook
            xmlstring = GetScoreCardXML(workbook, TranslationTypeId, TemplateName);

            // Save scorecard data
            result = SaveScoreCardData(EntityScorecardID, fileName, userID, xmlstring);
        }

        return result;
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
    finally
    {
        // Clean-up is done automatically by EPPlus within the using block
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
