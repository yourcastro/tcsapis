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


