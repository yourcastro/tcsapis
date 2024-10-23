 public string SaveScoreCardFileData(int EntityScorecardID, string FilePath, string TemplateName, string userID)
 {
     string xmlstring = "";
     string fileName = Path.GetFileName(FilePath);
     string result = "";
     Application xlApp = null;
     Workbook xlWorkBook = null;
     string TranslationTypeId = string.Empty;
     string Sessionkey = CreateSessionId();
     //Cryptography objCryptography = new Cryptography();
     string excelPwd = string.Empty;
     int? myExcelProcessId = 0;
     string strProcessName = string.Empty;

     try
     {
         ClsPDScoreCardFunctions.GetScoredCardTranslationTypeId(
             _appSettings.Application_Name,
             _appSettings.SQLUserName,
             _appSettings.SQLPassword,
             _appSettings.CR_DB_Operational,
             CreateSessionId(), TemplateName, ref TranslationTypeId);

         excelPwd = _appSettings.WorkBookPwd ?? string.Empty;

         xlApp = new Application();
         myExcelProcessId = GetExcelProcessID();
         xlApp.DisplayAlerts = false;
         xlWorkBook = xlApp.Workbooks.Open(FilePath, Password: excelPwd, ReadOnly: true, IgnoreReadOnlyRecommended: true, UpdateLinks: false);
         xlApp.Visible = false;

         xmlstring = GetScoreCardXML(xlWorkBook, TranslationTypeId, TemplateName);
         result = SaveScoreCardData(EntityScorecardID, fileName, userID, xmlstring);

         xlWorkBook.Close(SaveChanges: false);

         return "";
     }
     catch (Exception ex)
     {
         if (xlApp != null && xlWorkBook != null)
         {
             xlWorkBook.Close(SaveChanges: false);
         }
         return ex.Message;
     }
     finally
     {
         if (xlWorkBook != null)
         {
             System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkBook);
             xlWorkBook = null;
         }
         if (xlApp != null)
         {
             Workbooks wbs = xlApp.Workbooks;
             wbs.Close();
             System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wbs);
             wbs = null;

             xlApp.Quit();
             System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
             xlApp = null;
         }

         Thread.Sleep(200);
         GC.Collect();
         GC.WaitForPendingFinalizers();

         try
         {
             Process aProcess = Process.GetProcessById(myExcelProcessId ?? 0);
             strProcessName = aProcess.ProcessName;
             if (aProcess != null && strProcessName.ToUpper() == "EXCEL")
             {
                 aProcess.Kill();
             }
         }
         catch (Exception ex)
         {
         }
     }
 }
