 private string GetScoreCardXML(Workbook xlWorkBook, string TranslationTypeId, string TemplateName)
 {
     string strXML = "";
     
     DateTime dt;
     object nameRange;
     string result = "";
     Worksheet worksheet;

     System.Data.DataSet ds = ClsPDScoreCardFunctions.GetScoredCardTranslationDetails(
             _appSettings.Application_Name,
             _appSettings.SQLUserName,
             _appSettings.SQLPassword,
             _appSettings.CR_DB_Operational,
             CreateSessionId(), TranslationTypeId);


     strXML = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + "\r\n";
     strXML += "<PDScoreCord>" + "\r\n";
     strXML += "<Data>" + "\r\n";

     System.Data.DataTable tb = ds.Tables[0];

     if (TemplateName == "Public Finance")
     {
         for (int i = 0; i < tb.Rows.Count; i++)
         {
             worksheet = i <= 10 ? (Worksheet)xlWorkBook.Sheets["Summary"] : (Worksheet)xlWorkBook.Sheets["Template Assessment Sheet"];

             //Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Sheets["Template Assessment Sheet"];

             if (IsCVErr(worksheet.Range[tb.Rows[i]["ValueOut1"].ToString()].Value))
             {
                 strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"\"/>" + "\r\n";
             }
             else
             {
                 nameRange = worksheet.Range[tb.Rows[i]["ValueOut1"].ToString()].Value;

                 if (nameRange is double && nameRange.ToString().Contains("."))
                 {
                     if (nameRange.ToString().Split('.')[1].Length > 15)
                     {
                         result = string.Format("{0:0.##############E+00}", nameRange);
                         nameRange = result;
                     }
                 }

                 if (nameRange is DateTime)
                 {
                     dt = (DateTime)nameRange;
                     nameRange = dt.ToShortDateString();
                 }

                 strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"" + GetRangeValue(nameRange) + "\"/>" + "\r\n";
             }
         }
     }
     else
     {
         for (int i = 0; i < tb.Rows.Count; i++)
         {
             if (string.IsNullOrEmpty(tb.Rows[i]["ValueOut1"].ToString()))
             {
                 strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"\"/>" + "\r\n";
             }
             else if (RangeNameExists(xlWorkBook, tb.Rows[i]["ValueOut1"].ToString()))
             {
                 if (IsCVErr(xlWorkBook.Names.Item(tb.Rows[i]["ValueOut1"].ToString()).RefersToRange.Value))
                 {
                     strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"\"/>" + "\r\n";
                 }
                 else
                 {
                     nameRange = xlWorkBook.Names.Item(tb.Rows[i]["ValueOut1"].ToString()).RefersToRange.Value;

                     if (nameRange is double && nameRange.ToString().Contains("."))
                     {
                         if (nameRange.ToString().Split('.')[1].Length > 15)
                         {
                             result = string.Format("{0:0.##############E+00}", nameRange);
                             nameRange = result;
                         }
                     }

                     if (nameRange is DateTime)
                     {
                         dt = (DateTime)nameRange;
                         nameRange = dt.ToShortDateString();
                     }

                     strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"" + GetRangeValue(nameRange) + "\"/>" + "\r\n";
                 }
             }
         }
     }

     strXML += "</Data>" + "\r\n";
     strXML += "</PDScoreCord>" + "\r\n";
     strXML = strXML.Replace("<populated from EDS>", "");

     return strXML;
 }

 private bool RangeNameExists(Workbook activeWorkbook, string nname)
 {
     bool rangeNameExists = false;
     foreach (Name n in activeWorkbook.Names)
     {
         if (n.Name == nname)
         {
             rangeNameExists = true;
             break; // Exit the function
         }
     }
     return rangeNameExists;
 }


 private string GetRangeValue(object r)
 {
     if (r == null)
     {
         return "";
     }

     string myNumString = "";

     if (r.GetType().IsArray)
     {
         Array arrayResult = (Array)r;
         for (int i = 1; i <= arrayResult.Length; i++)
         {
             if (arrayResult.GetValue(1, i) != null)
             {
                 myNumString = arrayResult.GetValue(1, i).ToString();
             }
         }
     }
     else
     {
         myNumString = r.ToString();
     }

     myNumString = myNumString.Replace("&", "&amp;");
     myNumString = myNumString.Replace("<", "&lt;");
     myNumString = myNumString.Replace(">", "&gt;");
     myNumString = myNumString.Replace("\"", "&quot;");
     myNumString = myNumString.Replace("'", "&apos;");

     return myNumString;
 }
