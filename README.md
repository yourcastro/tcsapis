using OfficeOpenXml;
using System.Data;

private string GetScoreCardXML(ExcelWorkbook workbook, string TranslationTypeId, string TemplateName)
{
    string strXML = "";
    DateTime dt;
    object nameRange;
    string result = "";
    ExcelWorksheet worksheet;

    // Get scorecard translation details from the database
    DataSet ds = ClsPDScoreCardFunctions.GetScoredCardTranslationDetails(
        _appSettings.Application_Name,
        _appSettings.SQLUserName,
        _appSettings.SQLPassword,
        _appSettings.CR_DB_Operational,
        CreateSessionId(), TranslationTypeId);

    strXML = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + "\r\n";
    strXML += "<PDScoreCord>" + "\r\n";
    strXML += "<Data>" + "\r\n";

    DataTable tb = ds.Tables[0];

    if (TemplateName == "Public Finance")
    {
        for (int i = 0; i < tb.Rows.Count; i++)
        {
            // Get the worksheet based on the index condition
            worksheet = i <= 10 ? workbook.Worksheets["Summary"] : workbook.Worksheets["Template Assessment Sheet"];

            if (IsCVErr(worksheet.Cells[tb.Rows[i]["ValueOut1"].ToString()].Value))
            {
                strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"\"/>" + "\r\n";
            }
            else
            {
                nameRange = worksheet.Cells[tb.Rows[i]["ValueOut1"].ToString()].Value;

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
            else if (RangeNameExists(workbook, tb.Rows[i]["ValueOut1"].ToString()))
            {
                nameRange = workbook.Names[tb.Rows[i]["ValueOut1"].ToString()].Value;

                if (IsCVErr(nameRange))
                {
                    strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"\"/>" + "\r\n";
                }
                else
                {
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

private bool RangeNameExists(ExcelWorkbook workbook, string name)
{
    return workbook.Names[name] != null;
}

private string GetRangeValue(object r)
{
    if (r == null)
    {
        return "";
    }

    string myNumString = r.ToString();

    myNumString = myNumString.Replace("&", "&amp;");
    myNumString = myNumString.Replace("<", "&lt;");
    myNumString = myNumString.Replace(">", "&gt;");
    myNumString = myNumString.Replace("\"", "&quot;");
    myNumString = myNumString.Replace("'", "&apos;");

    return myNumString;
}
