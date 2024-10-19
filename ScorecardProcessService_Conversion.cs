public string GetUpdatedScoreCardFile(string filepath, string templatename, string userID)
{
    int index = 0;
    long id = 0;
    string name = "";
    string GICS_sector = "";
    string GICS_industry_group = "";
    string result = "";
    string lehmans_sector = "";
    Excel.Application xlApp = null;
    Excel.Workbook xlWorkBook = null;
    Excel.Names excelNames = null;
    Cryptography objCryptography = new Cryptography();
    List<string> wbNames = new List<string>();
    string fileName = Path.GetFileName(filepath);
    string initCellValue = "<populated from EDS>";
    string excelPwd = string.Empty;
    int myExcelProcessId = 0;
    string strProcessName = string.Empty;

    try
    {
        switch (templatename)
        {
            case "Financial Institutions - Commercial Banks":
                result = GetScoredCardForCommerialBank(fileName, ref id, ref name);
                break;
            case "Generic Large Corporate":
                result = GetScoredcardForGenericLargeCorporate(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group, ref lehmans_sector);
                break;
            case "Financial Institutions - Insurance CIQ":
                result = GetScoredCardForInsuranceCIQ(fileName, ref id, ref name);
                break;
            case "Financial Institutions - Insurance Non-CIQ":
                result = GetScoredCardForInsuranceNonCIQ(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group);
                break;
            case "Lease Finance":
                result = GetScoredcardForLeaseFinance(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group);
                break;
            case "Mortgage Bonds":
                result = GetScoredcardForMortgageBonds(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group);
                break;
            case "Financial Institutions - Non-Bank - CIQ":
                result = GetScoredCardForNonBankFinancialCIQ(fileName, ref id, ref name);
                break;
            case "Financial Institutions - Non-Bank - Non-CIQ":
                result = GetScoredCardForNonBankFinancialNonCIQ(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group);
                break;
            case "Oil and Gas – Exploration and Production CIQ":
                result = GetScoredCardForOilGasExplorationProductionCIQ(fileName, ref id, ref name, ref lehmans_sector);
                break;
            case "Oil and Gas – Exploration and Production Non-CIQ":
                result = GetScoredCardForOilGasExplorationProductionNonCIQ(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group, ref lehmans_sector);
                break;
            case "Oil and Gas – Midstream CIQ":
                result = GetScoredCardForOilGasMidstreamCIQ(fileName, ref id, ref name, ref lehmans_sector);
                break;
            case "Oil and Gas – Midstream Non-CIQ":
                result = GetScoredCardForOilGasMidstreamNonCIQ(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group, ref lehmans_sector);
                break;
            case "Oil and Gas – Oil Field Services CIQ":
                result = GetScoredCardForOilGasOilFieldServicesCIQ(fileName, ref id, ref name, ref lehmans_sector);
                break;
            case "Oil and Gas – Oil Field Services Non-CIQ":
                result = GetScoredCardForOilGasOilFieldServicesNonCIQ(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group, ref lehmans_sector);
                break;
            case "Oil and Gas – Refining and Marketing CIQ":
                result = GetScoredCardForOilGasRefiningMarketingCIQ(fileName, ref id, ref name, ref lehmans_sector);
                break;
            case "Oil and Gas – Refining and Marketing Non-CIQ":
                result = GetScoredCardForOilGasRefiningMarketingNonCIQ(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group, ref lehmans_sector);
                break;
            case "Project Finance":
                result = GetScoredCardForProjFinance(fileName, ref id, ref name);
                break;
            case "Public Finance":
                result = GetScoredCardForPublicFinancNonUS(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group);
                break;
            case "Real Estate Investment":
                result = GetScoredCardForRealEstateInvestment(fileName, ref id, ref name);
                break;
            case "Small Medium Enterprise":
                result = GetScoredcardForSmallMediumEnterprises(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group, ref lehmans_sector);
                break;
            case "University, School or Hospital":
                result = GetScoredCardForUniversitySchoolHospital(fileName, ref id, ref name, ref GICS_sector, ref GICS_industry_group);
                break;
            case "Utilities":
                result = GetScoredCardForUtilities(fileName, ref id, ref name);
                break;
        }

        excelPwd = ClsPDScoreCardEncryption.Decrypt("", AppSettings.Get(CONFIG_EXCEL_PWD));

        if (result == "OK")
        {
            CheckForExistingExcellProcesses();

            xlApp = new Excel.Application();

            myExcelProcessId = GetExcelProcessID();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(filepath, Password: excelPwd, ReadOnly: false, IgnoreReadOnlyRecommended: true, UpdateLinks: false);
            xlApp.Visible = false;

            if (templatename == "Public Finance")
            {
                xlWorkBook.Sheets["Template Assessment Sheet"].Range["D6"].Value = id;
                xlWorkBook.Sheets["Template Assessment Sheet"].Range["D4"].Value = name;
                xlWorkBook.Sheets["Template Assessment Sheet"].Range["D13"].Value = GICS_sector;
                xlWorkBook.Sheets["Template Assessment Sheet"].Range["D15"].Value = GICS_industry_group;

                if (id != 0)
                {
                    xlWorkBook.Sheets["Template Assessment Sheet"].Range["D11"].Value = DateTime.Now;
                    xlWorkBook.Sheets["Template Assessment Sheet"].Range["D9"].Value = userID;
                }
            }
            else
            {
                excelNames = xlWorkBook.Names;

                foreach (Excel.Name n in excelNames)
                {
                    wbNames.Add(n.Name);
                }

                if (wbNames.Contains("EntityID"))
                {
                    xlWorkBook.Names.Item("EntityID").RefersToRange.Value = id;
                }

                if (wbNames.Contains("EntityName"))
                {
                    xlWorkBook.Names.Item("EntityName").RefersToRange.Value = name;
                }

                if (wbNames.Contains("GICSSector") && !string.IsNullOrEmpty(GICS_sector))
                {
                    xlWorkBook.Names.Item("GICSSector").RefersToRange.Value = GICS_sector;
                }

                if (wbNames.Contains("GICSIndustryGroup") && !string.IsNullOrEmpty(GICS_industry_group))
                {
                    xlWorkBook.Names.Item("GICSIndustryGroup").RefersToRange.Value = GICS_industry_group;
                }

                if (wbNames.Contains("BarclaysName") && !string.IsNullOrEmpty(lehmans_sector))
                {
                    xlWorkBook.Names.Item("BarclaysName").RefersToRange.Value = lehmans_sector;
                }

                if (id != 0)
                {
                    if (wbNames.Contains("DateOfAnalysis"))
                    {
                        xlWorkBook.Names.Item("DateOfAnalysis").RefersToRange.Value = DateTime.Now;
                    }

                    if (wbNames.Contains("DateofAnalysis"))
                    {
                        xlWorkBook.Names.Item("DateofAnalysis").RefersToRange.Value = DateTime.Now;
                    }

                    if (wbNames.Contains("Analyst"))
                    {
                        xlWorkBook.Names.Item("Analyst").RefersToRange.Value = userID;
                    }

                    if (wbNames.Contains("Ccy"))
                    {
                        xlWorkBook.Names.Item("Ccy").RefersToRange.Value = userID;
                    }

                    if (wbNames.Contains("AnalystName"))
                    {
                        xlWorkBook.Names.Item("AnalystName").RefersToRange.Value = userID;
                    }
                }
            }

            xlWorkBook.Close(SaveChanges: true);
        }

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
        if (excelNames != null)
        {
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelNames);
            excelNames = null;
        }

        if (xlWorkBook != null)
        {
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkBook);
            xlWorkBook = null;
        }

        if (xlApp != null)
        {
            Excel.Workbooks wbs = xlApp.Workbooks;
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wbs);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }

        KillExcelProcess(myExcelProcessId);
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}



public string SaveScoreCardFileData(int EntityScorecardID, string FilePath, string TemplateName, string userID)
{
    string xmlstring = "";
    string fileName = Path.GetFileName(FilePath);
    string result = "";
    Excel.Application xlApp = null;
    Excel.Workbook xlWorkBook = null;
    string TranslationTypeId;
    string Sessionkey = CreateSessionId();
    Cryptography objCryptography = new Cryptography();
    string excelPwd = string.Empty;
    int myExcelProcessId = 0;
    string strProcessName = string.Empty;

    try
    {
        DataSet ds = new DataSet();
        sGenericTableRequestArguments objArguments = new sGenericTableRequestArguments(AppSettings.Get(CONFIG_APPNAME));

        objArguments.TableNames = new string[1];
        objArguments.FilterConditions = new string[1];
        objArguments.dsData = ds;
        objArguments.ConnectionDatabase = AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION);
        objArguments.ConnectionUser = AppSettings.Get(CONFIG_DB_USER);
        objArguments.ConnectionPwd = AppSettings.Get(CONFIG_DB_PWD);
        objArguments.TableNames[0] = "TranslationType";
        objArguments.FilterConditions[0] = "Code = '" + TemplateName + "'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION);
        objArguments.AuditConnectionUsers[0] = AppSettings.Get(CONFIG_DB_USER);
        objArguments.AuditConnectionPwds[0] = AppSettings.Get(CONFIG_DB_PWD);

        GetData(objArguments, Sessionkey);

        DataTable tb = ds.Tables[0];
        DataRow row = tb.Rows[0];
        TranslationTypeId = row["TranslationTypeId"].ToString();

        excelPwd = ClsPDScoreCardEncryption.Decrypt("", AppSettings.Get(CONFIG_EXCEL_PWD));

        xlApp = new Excel.Application();
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
            Excel.Workbooks wbs = xlApp.Workbooks;
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
            Process aProcess = Process.GetProcessById(myExcelProcessId);
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



private string GetScoreCardXML(Excel.Workbook xlWorkBook, string TranslationTypeId, string TemplateName)
{
    string strXML = "";
    DataSet ds = new DataSet();
    DateTime dt;
    object nameRange;
    string result = "";
    object worksheet;
    sGenericTableRequestArguments objArguments = new sGenericTableRequestArguments(AppSettings.Get(CONFIG_APPNAME));

    objArguments.TableNames = new string[1];
    objArguments.FilterConditions = new string[1];
    objArguments.dsData = ds;
    objArguments.ConnectionDatabase = AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION);
    objArguments.ConnectionUser = AppSettings.Get(CONFIG_DB_USER);
    objArguments.ConnectionPwd = AppSettings.Get(CONFIG_DB_PWD);
    objArguments.TableNames[0] = "Translation";
    objArguments.FilterConditions[0] = "TranslationTypeId = '" + TranslationTypeId + "'";

    objArguments.AuditConnectionDatabases = new string[1];
    objArguments.AuditConnectionUsers = new string[1];
    objArguments.AuditConnectionPwds = new string[1];
    objArguments.AuditConnectionDatabases[0] = AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION);
    objArguments.AuditConnectionUsers[0] = AppSettings.Get(CONFIG_DB_USER);
    objArguments.AuditConnectionPwds[0] = AppSettings.Get(CONFIG_DB_PWD);

    GetData(objArguments, CreateSessionId());

    strXML = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + "\r\n";
    strXML += "<PDScoreCord>" + "\r\n";
    strXML += "<Data>" + "\r\n";

    DataTable tb = ds.Tables[0];

    if (TemplateName == "Public Finance")
    {
        for (int i = 0; i < tb.Rows.Count; i++)
        {
            worksheet = i <= 10 ? xlWorkBook.Sheets["Summary"] : xlWorkBook.Sheets["Template Assessment Sheet"];

            if (IsCVErr(worksheet.Range(tb.Rows[i]["ValueOut1"].ToString()).Value))
            {
                strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"\"/>" + "\r\n";
            }
            else
            {
                nameRange = worksheet.Range(tb.Rows[i]["ValueOut1"].ToString()).Value;

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
                if (IsCVErr(xlWorkBook.Names[tb.Rows[i]["ValueOut1"].ToString()].RefersToRange.Value))
                {
                    strXML += "<" + tb.Rows[i]["ValueIn1"].ToString() + " value=\"\"/>" + "\r\n";
                }
                else
                {
                    nameRange = xlWorkBook.Names[tb.Rows[i]["ValueOut1"].ToString()].RefersToRange.Value;

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




