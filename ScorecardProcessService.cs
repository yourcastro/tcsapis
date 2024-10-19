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

    
