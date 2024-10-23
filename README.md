using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;

public string GetUpdatedScoreCardFile(string filepath, string templatename, string userID)
{
    int id = 0;
    string name = "";
    string GICS_sector = "";
    string GICS_industry_group = "";
    string result = "";
    string lehmans_sector = "";
    CryptorEngine objCryptography = new CryptorEngine();
    List<string> wbNames = new List<string>();
    string fileName = Path.GetFileName(filepath);
    string excelPwd = string.Empty;

    try
    {
        // Select the appropriate template
        switch (templatename)
        {
            case "Financial Institutions - Commercial Banks":
                result = GetScoredCardForCommerialBank(fileName, ref id, ref name);
                break;
            // Add other cases here as necessary
        }

        excelPwd = ClsPDScoreCardEncryption.Decrypt("", _appSettings?.WorkBookPwd ?? string.Empty);

        if (result == "OK")
        {
            // Open the Excel file using EPPlus
            FileInfo file = new FileInfo(filepath);

            // EPPlus requires a license declaration in commercial use cases
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; 

            using (var package = new ExcelPackage(file, excelPwd))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[0]; // Assuming you want the first sheet

                if (templatename == "Public Finance")
                {
                    worksheet.Cells["D6"].Value = id;
                    worksheet.Cells["D4"].Value = name;
                    worksheet.Cells["D13"].Value = GICS_sector;
                    worksheet.Cells["D15"].Value = GICS_industry_group;

                    if (id != 0)
                    {
                        worksheet.Cells["D11"].Value = DateTime.Now;
                        worksheet.Cells["D9"].Value = userID;
                    }
                }
                else
                {
                    foreach (var nameItem in workbook.Names)
                    {
                        wbNames.Add(nameItem.Name);
                    }

                    if (wbNames.Contains("EntityID"))
                    {
                        workbook.Names["EntityID"].Value = id;
                    }

                    if (wbNames.Contains("EntityName"))
                    {
                        workbook.Names["EntityName"].Value = name;
                    }

                    if (wbNames.Contains("GICSSector") && !string.IsNullOrEmpty(GICS_sector))
                    {
                        workbook.Names["GICSSector"].Value = GICS_sector;
                    }

                    if (wbNames.Contains("GICSIndustryGroup") && !string.IsNullOrEmpty(GICS_industry_group))
                    {
                        workbook.Names["GICSIndustryGroup"].Value = GICS_industry_group;
                    }

                    if (wbNames.Contains("BarclaysName") && !string.IsNullOrEmpty(lehmans_sector))
                    {
                        workbook.Names["BarclaysName"].Value = lehmans_sector;
                    }

                    if (id != 0)
                    {
                        if (wbNames.Contains("DateOfAnalysis"))
                        {
                            workbook.Names["DateOfAnalysis"].Value = DateTime.Now;
                        }

                        if (wbNames.Contains("DateofAnalysis"))
                        {
                            workbook.Names["DateofAnalysis"].Value = DateTime.Now;
                        }

                        if (wbNames.Contains("Analyst"))
                        {
                            workbook.Names["Analyst"].Value = userID;
                        }

                        if (wbNames.Contains("Ccy"))
                        {
                            workbook.Names["Ccy"].Value = userID;
                        }

                        if (wbNames.Contains("AnalystName"))
                        {
                            workbook.Names["AnalystName"].Value = userID;
                        }
                    }
                }

                // Save the changes
                package.Save();
            }
        }

        return "";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}
