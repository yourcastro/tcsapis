Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Configuration.ConfigurationManager
Imports CreditRiskPortalWSLib
Imports IT.INV.Service.Base
Imports IT.CIBU.SecurityFunctions

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class ScorecardProcessService
    Inherits System.Web.Services.WebService


    Public Const CONFIG_APPNAME As String = "Application_Name"
    Public Const CONFIG_DB_USER As String = "SQLUserName"
    Public Const CONFIG_DB_PWD As String = "SQLPassword"
    Public Const CONFIG_EXCEL_PWD As String = "WorkBookPwd"
    Public Const CONFIG_SESSION_MODE_DB As String = "Store_Session_In_DB"
    Public Const CONFIG_DB_DATASTORE_CONNECTION As String = "CR_DB_Datastore"
    Public Const CONFIG_DB_OPERATIONAL_CONNECTION As String = "CR_DB_Operational"
    Public Const CONFIG_DB_SCORECARD_CONNECTION As String = "CR_DB_Scorecard"
    Dim listProcess As New List(Of Integer)
    Public Enum CVErrEnum As Int32
        ErrDiv0 = -2146826281
        ErrNA = -2146826246
        ErrName = -2146826259
        ErrNull = -2146826288
        ErrNum = -2146826252
        ErrRef = -2146826265
        ErrValue = -2146826273
    End Enum
    Public Function IsCVErr(ByVal obj As Object) As Boolean
        If IsNumeric(obj) Then
            Select Case CType(obj, Integer)
                Case CVErrEnum.ErrDiv0, CVErrEnum.ErrNA, CVErrEnum.ErrName, CVErrEnum.ErrNull, CVErrEnum.ErrNum, CVErrEnum.ErrRef, CVErrEnum.ErrValue
                    Return True
                Case Else
                    Return False
            End Select
        End If
        Return False
    End Function
    <WebMethod()>
    Public Function GetUpdatedScoreCardFile(ByVal filepath As String, ByVal templatename As String, ByVal userID As String) As String
        Dim index = 0
        Dim id As Long = 0
        Dim name As String = ""
        Dim GICS_sector As String = ""
        Dim GICS_industry_group As String = ""
        Dim result As String = ""
        Dim lehmans_sector As String = ""
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim excelNames As Excel.Names = Nothing
        Dim objCryptography As New Cryptography
        Dim wbNames As New List(Of String)()
        Dim fileName As String = Path.GetFileName(filepath)
        Dim initCellValue = "<populated from EDS>"
        Dim excelPwd As String = String.Empty
        Dim myExcelProcessId As Integer = 0
        Dim strProcessName As String = String.Empty

        Try
            Select Case templatename
                Case "Financial Institutions - Commercial Banks"
                    result = GetScoredCardForCommerialBank(fileName, id, name)
                Case "Generic Large Corporate"
                    result = GetScoredcardForGenericLargeCorporate(fileName, id, name, GICS_sector, GICS_industry_group, lehmans_sector)
                Case "Financial Institutions - Insurance CIQ"
                    result = GetScoredCardForInsuranceCIQ(fileName, id, name)
                Case "Financial Institutions - Insurance Non-CIQ"
                    result = GetScoredCardForInsuranceNonCIQ(fileName, id, name, GICS_sector, GICS_industry_group)
                Case "Lease Finance"
                    result = GetScoredcardForLeaseFinance(fileName, id, name, GICS_sector, GICS_industry_group)
                Case "Mortgage Bonds"
                    result = GetScoredcardForMortgageBonds(fileName, id, name, GICS_sector, GICS_industry_group)
                Case "Financial Institutions - Non-Bank - CIQ"
                    result = GetScoredCardForNonBankFinancialCIQ(fileName, id, name)
                Case "Financial Institutions - Non-Bank - Non-CIQ"
                    result = GetScoredCardForNonBankFinancialNonCIQ(fileName, id, name, GICS_sector, GICS_industry_group)
                Case "Oil and Gas – Exploration and Production CIQ"
                    result = GetScoredCardForOilGasExplorationProductionCIQ(fileName, id, name, lehmans_sector)
                Case "Oil and Gas – Exploration and Production Non-CIQ"
                    result = GetScoredCardForOilGasExplorationProductionNonCIQ(fileName, id, name, GICS_sector, GICS_industry_group, lehmans_sector)
                Case "Oil and Gas – Midstream CIQ"
                    result = GetScoredCardForOilGasMidstreamCIQ(fileName, id, name, lehmans_sector)
                Case "Oil and Gas – Midstream Non-CIQ"
                    result = GetScoredCardForOilGasMidstreamNonCIQ(fileName, id, name, GICS_sector, GICS_industry_group, lehmans_sector)
                Case "Oil and Gas – Oil Field Services CIQ"
                    result = GetScoredCardForOilGasOilFieldServicesCIQ(fileName, id, name, lehmans_sector)
                Case "Oil and Gas – Oil Field Services Non-CIQ"
                    result = GetScoredCardForOilGasOilFieldServicesNonCIQ(fileName, id, name, GICS_sector, GICS_industry_group, lehmans_sector)
                Case "Oil and Gas – Refining and Marketing CIQ"
                    result = GetScoredCardForOilGasRefiningMarketingCIQ(fileName, id, name, lehmans_sector)
                Case "Oil and Gas – Refining and Marketing Non-CIQ"
                    result = GetScoredCardForOilGasRefiningMarketingNonCIQ(fileName, id, name, GICS_sector, GICS_industry_group, lehmans_sector)
                Case "Project Finance"
                    result = GetScoredCardForProjFinance(fileName, id, name)
                Case "Public Finance"
                    result = GetScoredCardForPublicFinancNonUS(fileName, id, name, GICS_sector, GICS_industry_group)
                Case "Real Estate Investment"
                    result = GetScoredCardForRealEstateInvestment(fileName, id, name)
                Case "Small Medium Enterprise"
                    result = GetScoredcardForSmallMediumEnterprises(fileName, id, name, GICS_sector, GICS_industry_group, lehmans_sector)
                Case "University, School or Hospital"
                    result = GetScoredCardForUniversitySchoolHospital(fileName, id, name, GICS_sector, GICS_industry_group)
                Case "Utilities"
                    result = GetScoredCardForUtilities(fileName, id, name)

            End Select

            excelPwd = ClsPDScoreCardEncryption.Decrypt("", AppSettings.Get(CONFIG_EXCEL_PWD))

            If result = "OK" Then

                CheckForExistingExcellProcesses()

                xlApp = New Excel.Application

                myExcelProcessId = GetExcelProcessID()
                xlApp.DisplayAlerts = False
                xlWorkBook = xlApp.Workbooks.Open(filepath, Password:=excelPwd, [ReadOnly]:=False, IgnoreReadOnlyRecommended:=True, UpdateLinks:=False)
                xlApp.Visible = False

                If templatename = "Public Finance" Then
                    xlWorkBook.Sheets("Template Assessment Sheet").Range("D6").Value = id
                    xlWorkBook.Sheets("Template Assessment Sheet").Range("D4").Value = name
                    xlWorkBook.Sheets("Template Assessment Sheet").Range("D13").Value = GICS_sector
                    xlWorkBook.Sheets("Template Assessment Sheet").Range("D15").Value = GICS_industry_group
                    If id <> 0 Then
                        xlWorkBook.Sheets("Template Assessment Sheet").Range("D11").Value = DateTime.Now
                        xlWorkBook.Sheets("Template Assessment Sheet").Range("D9").Value = userID

                    End If

                Else
                    excelNames = xlWorkBook.Names

                    For Each n In excelNames
                        wbNames.Add(n.Name.ToString())
                    Next n

                    If wbNames.Contains("EntityID") Then 'RangeNameExists(xlWorkBook, "EntityID") Then
                        xlWorkBook.Names.Item("EntityID").RefersToRange.Value = id
                    End If

                    If wbNames.Contains("EntityName") Then
                        'If Trim(xlWorkBook.Names.Item("EntityName").Value) = initCellValue Or Trim(xlWorkBook.Names.Item("EntityName").Value) = "" Then
                        xlWorkBook.Names.Item("EntityName").RefersToRange.Value = name
                        'End If
                    End If

                    If wbNames.Contains("GICSSector") And String.IsNullOrEmpty(GICS_sector) = False Then
                        'If Trim(xlWorkBook.Names.Item("GICSSector").Value) = initCellValue Or Trim(xlWorkBook.Names.Item("GICSSector").Value) = "" Then
                        xlWorkBook.Names.Item("GICSSector").RefersToRange.Value = GICS_sector
                        ' End If

                    End If

                    If wbNames.Contains("GICSIndustryGroup") And String.IsNullOrEmpty(GICS_industry_group) = False Then
                        ' If Trim(xlWorkBook.Names.Item("GICSIndustryGroup").Value) = initCellValue Or Trim(xlWorkBook.Names.Item("GICSIndustryGroup").Value) = "" Then
                        xlWorkBook.Names.Item("GICSIndustryGroup").RefersToRange.Value = GICS_industry_group
                        'End If
                    End If

                    If wbNames.Contains("BarclaysName") And String.IsNullOrEmpty(lehmans_sector) = False Then
                        'If Trim(xlWorkBook.Names.Item("BarclaysName").Value) = initCellValue Or Trim(xlWorkBook.Names.Item("BarclaysName").Value) = "" Then
                        xlWorkBook.Names.Item("BarclaysName").RefersToRange.Value = lehmans_sector
                        'End If
                    End If


                    If id <> 0 Then

                        If wbNames.Contains("DateOfAnalysis") Then
                            xlWorkBook.Names.Item("DateOfAnalysis").RefersToRange.Value = DateTime.Now
                        End If

                        If wbNames.Contains("DateofAnalysis") Then
                            xlWorkBook.Names.Item("DateofAnalysis").RefersToRange.Value = DateTime.Now
                        End If

                        If wbNames.Contains("Analyst") Then
                            xlWorkBook.Names.Item("Analyst").RefersToRange.Value = userID
                        End If
                        If wbNames.Contains("Ccy") Then
                            xlWorkBook.Names.Item("Ccy").RefersToRange.Value = userID
                        End If
                        If wbNames.Contains("AnalystName") Then
                            xlWorkBook.Names.Item("AnalystName").RefersToRange.Value = userID
                        End If
                    End If
                End If
                xlWorkBook.Close(SaveChanges:=True)
            End If

            Return ""
        Catch ex As Exception
            If xlApp IsNot Nothing And xlWorkBook IsNot Nothing Then
                xlWorkBook.Close(SaveChanges:=False)
            End If
            Return ex.Message
        Finally
            If excelNames IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelNames)
                excelNames = Nothing
            End If
            If xlWorkBook IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If
            If xlApp IsNot Nothing Then
                Dim wbs As Excel.Workbooks = xlApp.Workbooks
                wbs.Close()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wbs)
                wbs = Nothing

                xlApp.Quit()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If
            System.Threading.Thread.Sleep(200)
            GC.Collect()
            GC.WaitForPendingFinalizers()

            Try
                Dim aProcess As Process = Nothing
                strProcessName = Process.GetProcessById(myExcelProcessId).ProcessName
                aProcess = Process.GetProcessById(myExcelProcessId)
                If aProcess IsNot Nothing And UCase(strProcessName) = "EXCEL" Then
                    aProcess.Kill()
                End If
            Catch ex As Exception
            End Try
        End Try
    End Function
    Private Function CheckForExistingExcellProcesses()
        Dim AllProcesses() As Process = Process.GetProcessesByName("EXCEL")
        For Each excelProcess As Process In AllProcesses
            listProcess.Add(excelProcess.Id)
        Next
    End Function

    Private Function GetExcelProcessID()
        Dim AllProcesses() As Process = Process.GetProcessesByName("EXCEL")
        For Each excelProcess As Process In AllProcesses
            If listProcess.Contains(excelProcess.Id) = False Then
                Return excelProcess.Id
            End If
        Next
        AllProcesses = Nothing
    End Function

    <WebMethod()>
    Public Function SaveScoreCardFileData(ByVal EntityScorecardID As Integer, ByVal FilePath As String, ByVal TemplateName As String, ByVal userID As String) As String

        Dim xmlstring As String = ""
        Dim fileName As String = Path.GetFileName(FilePath)
        Dim result As String = ""
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim TranslationTypeId As String
        Dim Sessionkey As String = CreateSessionId()
        Dim objCryptography As New Cryptography
        Dim excelPwd As String = String.Empty
        Dim myExcelProcessId As Integer = 0
        Dim strProcessName As String = String.Empty

        Try
            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(AppSettings.Get(CONFIG_APPNAME))
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION)
                .ConnectionUser = AppSettings.Get(CONFIG_DB_USER)
                .ConnectionPwd = AppSettings.Get(CONFIG_DB_PWD)
                .TableNames(0) = "TranslationType"
                .FilterConditions(0) = "Code = '" + TemplateName + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION)
                .AuditConnectionUsers(0) = AppSettings.Get(CONFIG_DB_USER)
                .AuditConnectionPwds(0) = AppSettings.Get(CONFIG_DB_PWD)

            End With

            GetData(objArguments, Sessionkey)

            Dim tb As Data.DataTable = ds.Tables(0)
            Dim row As DataRow = tb.Rows(0)
            TranslationTypeId = row.Item("TranslationTypeId")

            excelPwd = ClsPDScoreCardEncryption.Decrypt("", AppSettings.Get(CONFIG_EXCEL_PWD))

            xlApp = New Excel.Application
            myExcelProcessId = GetExcelProcessID()
            xlApp.DisplayAlerts = False
            xlWorkBook = xlApp.Workbooks.Open(FilePath, Password:=excelPwd, [ReadOnly]:=True, IgnoreReadOnlyRecommended:=True, UpdateLinks:=False)
            xlApp.Visible = False

            xmlstring = GetScoreCardXML(xlWorkBook, TranslationTypeId, TemplateName)
            result = SaveScoreCardData(EntityScorecardID, fileName, userID, xmlstring)

            xlWorkBook.Close(SaveChanges:=False)

            Return ""
        Catch ex As Exception
            If xlApp IsNot Nothing And xlWorkBook IsNot Nothing Then
                xlWorkBook.Close(SaveChanges:=False)
            End If
            Return ex.Message
        Finally
            If xlWorkBook IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If
            If xlApp IsNot Nothing Then
                Dim wbs As Excel.Workbooks = xlApp.Workbooks
                wbs.Close()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wbs)
                wbs = Nothing

                xlApp.Quit()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If

            System.Threading.Thread.Sleep(200)
            GC.Collect()
            GC.WaitForPendingFinalizers()

            Try
                Dim aProcess As Process = Nothing
                strProcessName = Process.GetProcessById(myExcelProcessId).ProcessName
                aProcess = Process.GetProcessById(myExcelProcessId)
                If aProcess IsNot Nothing And UCase(strProcessName) = "EXCEL" Then
                    aProcess.Kill()
                End If
            Catch ex As Exception
            End Try
        End Try

    End Function


    Public Function GetScoredCardForCommerialBank(ByVal filename As String, ByRef ID As Integer,
        ByRef name As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForCommerialBank(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetScoredcardForGenericLargeCorporate(ByVal filename As String, ByRef ID As Integer,
  ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredcardForGenericLargeCorporate(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForInsuranceCIQ(ByVal filename As String, ByRef ID As Integer,
       ByRef name As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForInsuranceCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForInsuranceNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForInsuranceNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetScoredcardForLeaseFinance(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForLeaseFinance(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredcardForMortgageBonds(ByVal filename As String, ByRef ID As Integer,
        ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForMortgageBonds(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForNonBankFinancialCIQ(ByVal filename As String, ByRef ID As Integer,
       ByRef name As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForNonBankFinancialCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForNonBankFinancialNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForNonBankFinancialNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasExplorationProductionCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasExplorationProductionCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasExplorationProductionNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasExplorationProductionNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasMidstreamCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasMidstreamCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasMidstreamNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasMidstreamNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasOilFieldServicesCIQ(ByVal filename As String, ByRef ID As Integer,
ByRef name As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasOilFieldServicesCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasOilFieldServicesNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasOilFieldServicesNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasRefiningMarketingCIQ(ByVal filename As String, ByRef ID As Integer,
ByRef name As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasRefiningMarketingCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForOilGasRefiningMarketingNonCIQ(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForOilGasRefiningMarketingNonCIQ(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForUtilities(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForUtilities(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForProjFinance(ByVal filename As String, ByRef ID As Integer,
    ByRef name As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForProjFinance(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForPublicFinancNonUS(ByVal filename As String, ByRef ID As Integer,
        ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_IG As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForPublicFinancNonUS(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_IG)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForRealEstateInvestment(ByVal filename As String, ByRef ID As Integer,
ByRef name As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForRealEstateInvestment(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredcardForSmallMediumEnterprises(ByVal filename As String, ByRef ID As Integer,
   ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String
        Try
            Return ClsPDScoreCardFunctions.GetScoredcardForSmallMediumEnterprises(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function GetScoredCardForUniversitySchoolHospital(ByVal filename As String, ByRef ID As Integer,
ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            Return ClsPDScoreCardFunctions.GetScoredCardForUniversitySchoolHospital(AppSettings.Get(CONFIG_APPNAME),
                            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                            AppSettings.Get(CONFIG_DB_USER),
                            AppSettings.Get(CONFIG_DB_PWD),
                            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                            CreateSessionId(),
                            filename, ID, name, GICS_Sector, GICS_Industry_Group)
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function SaveScoreCardData(ByVal fileID As Integer, ByVal filename As String, ByVal userID As String, ByVal xml As String) As String
        Try
            Dim proc As New ClsPDScoreCardProcesor(fileID, filename, userID,
                AppSettings.Get(CONFIG_APPNAME),
                AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
                AppSettings.Get(CONFIG_DB_USER),
                AppSettings.Get(CONFIG_DB_PWD),
                AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
                CreateSessionId(), xml)
            proc.WriteDB()

            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Sub GetData(ByRef Arguments As sGenericTableRequestArguments, ByVal SessionKey As String, Optional ByVal GetSecurity As Boolean = False)
        Dim iCount As Integer
        Dim strErrors As String = ""
        Dim objServices As New sGenericTable
        Dim objResultBody As New ResultBody

        If Not GetSecurity AndAlso Arguments.TableNames.Length > 0 Then
            ReDim Arguments.GetSecurityInfo(Arguments.TableNames.Length)
        End If

        objResultBody = ServiceFunctions.Execute(objServices, "GetTables", Arguments, SessionKey)

        If objResultBody.ResultStatus.ReturnCode = "Error" Then
            For iCount = 0 To objResultBody.ResultStatus.MessageCount - 1
                strErrors &= objResultBody.ResultStatus.ResultMessage(iCount).MessageDescription & ControlChars.CrLf
            Next
            Throw New ApplicationException(strErrors)
        End If
    End Sub

    Private Function GetScoreCardXML(ByVal xlWorkBook As Excel.Workbook, ByVal TranslationTypeId As String, ByVal TemplateName As String) As String

        Dim strXML As String = ""
        Dim ds As New DataSet
        Dim dt As Date
        Dim nameRange As Object
        Dim result As String = ""
        Dim worksheet As Object
        Dim objArguments As New sGenericTableRequestArguments(AppSettings.Get(CONFIG_APPNAME))
        With objArguments
            ReDim .TableNames(0)
            ReDim .FilterConditions(0)
            .dsData = ds
            .ConnectionDatabase = AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION)
            .ConnectionUser = AppSettings.Get(CONFIG_DB_USER)
            .ConnectionPwd = AppSettings.Get(CONFIG_DB_PWD)
            .TableNames(0) = "Translation"
            .FilterConditions(0) = "TranslationTypeId = '" + TranslationTypeId + "'"

            ReDim .AuditConnectionDatabases(0)
            ReDim .AuditConnectionUsers(0)
            ReDim .AuditConnectionPwds(0)
            .AuditConnectionDatabases(0) = AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION)
            .AuditConnectionUsers(0) = AppSettings.Get(CONFIG_DB_USER)
            .AuditConnectionPwds(0) = AppSettings.Get(CONFIG_DB_PWD)
        End With

        GetData(objArguments, CreateSessionId())

        strXML = "<?xml version=""1.0"" encoding=""utf-8""?>" & Chr(13) & Chr(10)
        strXML = strXML & "<PDScoreCord>" & Chr(13) & Chr(10)
        strXML = strXML & "<Data>" & Chr(13) & Chr(10)

        Dim tb As Data.DataTable = ds.Tables(0)

        If TemplateName = "Public Finance" Then
            For i = 0 To tb.Rows.Count - 1

                If i <= 10 Then
                    worksheet = xlWorkBook.Sheets("Summary")
                Else
                    worksheet = xlWorkBook.Sheets("Template Assessment Sheet")
                End If

                If IsCVErr(worksheet.Range(tb.Rows.Item(i).Item("ValueOut1").ToString()).Value) Then
                    strXML = strXML & "<" & tb.Rows.Item(i).Item("ValueIn1").ToString() & " value=" & """""/>" & Chr(13) & Chr(10)
                Else
                    nameRange = worksheet.Range(tb.Rows.Item(i).Item("ValueOut1").ToString()).Value
                    If TypeOf (nameRange) Is Double Then
                        If nameRange.ToString.Contains(".") Then
                            If nameRange.ToString.Split(".")(1).Length > 15 Then
                                result = String.Format("{0:0.##############E+00}", nameRange)
                                nameRange = result
                            End If
                        End If
                    End If
                    If TypeOf (nameRange) Is Date Then
                        dt = nameRange
                        nameRange = dt.ToShortDateString
                    End If
                    strXML = strXML & "<" & tb.Rows.Item(i).Item("ValueIn1").ToString() & " value=""" & GetRangeValue(nameRange) & """/>" & Chr(13) & Chr(10)
                End If
            Next

        Else
            For i = 0 To tb.Rows.Count - 1

                If tb.Rows.Item(i).Item("ValueOut1").Equals("") Then
                    strXML = strXML & "<" & tb.Rows.Item(i).Item("ValueIn1").ToString() & " value=" & """""/>" & Chr(13) & Chr(10)
                ElseIf (RangeNameExists(xlWorkBook, tb.Rows.Item(i).Item("ValueOut1").ToString())) Then
                    If IsCVErr(xlWorkBook.Names.Item(tb.Rows.Item(i).Item("ValueOut1").ToString()).RefersToRange.Value) Then
                        strXML = strXML & "<" & tb.Rows.Item(i).Item("ValueIn1").ToString() & " value=" & """""/>" & Chr(13) & Chr(10)
                    Else
                        nameRange = xlWorkBook.Names.Item(tb.Rows.Item(i).Item("ValueOut1").ToString()).RefersToRange.Value
                        If TypeOf (nameRange) Is Double Then
                            If nameRange.ToString.Contains(".") Then
                                If nameRange.ToString.Split(".")(1).Length > 15 Then
                                    result = String.Format("{0:0.##############E+00}", nameRange)
                                    nameRange = result
                                End If
                            End If
                        End If
                        If TypeOf (nameRange) Is Date Then
                            dt = nameRange
                            nameRange = dt.ToShortDateString
                        End If
                        strXML = strXML & "<" & tb.Rows.Item(i).Item("ValueIn1").ToString() & " value=""" & GetRangeValue(nameRange) & """/>" & Chr(13) & Chr(10)
                    End If
                End If
            Next
        End If

        strXML = strXML & "</Data>" & Chr(13) & Chr(10)
        strXML = strXML & "</PDScoreCord>" & Chr(13) & Chr(10)
        strXML = strXML.Replace("<populated from EDS>", "")

        Return strXML
    End Function
    Private Function GetRangeValue(r As Object) As String
        If (r Is Nothing) Then
            Return ""
        End If
        Dim arrayResult As Array
        Dim myNumString As String = ""
        Dim result As String = ""

        If r.GetType.IsArray Then
            arrayResult = r
            For i = 1 To arrayResult.Length()

                If (arrayResult(1, i) IsNot Nothing) Then
                    myNumString = arrayResult(1, i).ToString
                End If
            Next
        Else
            myNumString = r.ToString()
        End If

        myNumString = myNumString.Replace("&", "&amp;")
        myNumString = myNumString.Replace("<", "&lt;")
        myNumString = myNumString.Replace(">", "&gt;")
        myNumString = myNumString.Replace("""", "&quot;")
        myNumString = myNumString.Replace("'", "&apos;")
        result = result + myNumString

        Return result
        Exit Function
    End Function


    Public Shared Function CreateSessionId() As String
        Return Guid.NewGuid.ToString
    End Function

    Private Function RangeNameExists(ActiveWorkbook As Excel.Workbook, nname As String) As Boolean
        Dim n As Excel.Name
        RangeNameExists = False
        For Each n In ActiveWorkbook.Names
            If n.Name = nname Then
                RangeNameExists = True
                Exit Function
            End If
        Next n
    End Function

    Private Sub releaseObject(ByVal obj As Object)                          'Closes an object using the garbage collector
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
