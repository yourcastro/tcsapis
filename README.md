Imports IT.INV.Service.Base
Imports System.Data

Public Class ClsPDScoreCardFunctions
    Public Shared Function CheckAccess(ByVal userID As String, ByVal filename As String, ByVal appName As String, ByVal strConnDB As String, _
        ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
        ByVal sessionKey As String, ByRef errorMsg As String) As Boolean

        errorMsg = ""
        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                Return True
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            Dim currentUser As String
            Dim currentStatus As String

            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_and_factor_v"
                .FilterConditions(0) = "party_entity_scorecard_file_nm = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                errorMsg = "File " + filename + " has no record saved in the database."
                Return False
            Else
                Dim row As DataRow = tb.Rows(0)
                currentUser = CType(row.Item("last_update_process_id"), String)
                currentStatus = CType(row.Item("scorecard_status_cd"), String)
                If currentUser.ToUpper = userID.ToUpper And currentStatus.ToUpper = "O" Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            errorMsg = ex.Message
            Return False
        End Try
    End Function

    Public Shared Function GetScoredCardForCommerialBank(ByVal appName As String, ByVal strConnDB As String, _
        ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
        ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
        ByRef name As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_commerial_bank"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForPublicFinancNonUS(ByVal appName As String, ByVal strConnDB As String, _
        ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
        ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
        ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_IG As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                GICS_Sector = ""
                GICS_IG = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_public_finance_nonus"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_IG = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Shared Function GetScoredCardForGenericCorporateCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_generic_corp_ciq"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForGenericCorporateNonCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_generic_corp_non_ciq"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForUtilities(ByVal appName As String, ByVal strConnDB As String,
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String,
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, ByRef name As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_Utilities"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForOilGasExplorationProductionCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_OilGasExplorationProductionCIQ"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForOilGasMidstreamCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_OilGasMidstreamCIQ"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForOilGasOilFieldServicesCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_OilGasOilFieldServicesCIQ"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForOilGasRefiningMarketingCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_OilGasRefiningMarketingCIQ"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForOilGasExplorationProductionNonCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_OilGasExplorationProductionNonCIQ"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForOilGasMidstreamNonCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_OilGasMidstreamNonCIQ"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForOilGasOilFieldServicesNonCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_OilGasOilFieldServicesNonCIQ"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForOilGasRefiningMarketingNonCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_OilGasRefiningMarketingNonCIQ"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForCommercialMortgage(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Property_Address As String, _
    ByRef Loan_Number As String, ByRef Borrower_Name As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Property_Address = ""
                Loan_Number = ""
                Borrower_Name = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_CommercialMortgage_v"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
                Property_Address = CType(row.Item("Property_Address"), String)
                Loan_Number = CType(row.Item("Loan_Number"), String)
                Borrower_Name = CType(row.Item("Borrower_Name"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function


    Public Shared Function GetScoredCardForProjFinance(ByVal appName As String, ByVal strConnDB As String,
ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String,
ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, ByRef name As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_ProjectFinance"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForNonBankFinancialCIQ(ByVal appName As String, ByVal strConnDB As String, _
       ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
       ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
       ByRef name As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_NonBankFinancial_ciq"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForNonBankFinancialNonCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_NonBankFinancial_non_ciq"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForInsuranceCIQ(ByVal appName As String, ByVal strConnDB As String, _
       ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
       ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
       ByRef name As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_Insurance_ciq"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Shared Function GetScoredCardForInsuranceNonCIQ(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_Insurance_non_ciq"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForRealEstateInvestment(ByVal appName As String, ByVal strConnDB As String, _
      ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
      ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
      ByRef name As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_RealEstateInvestments"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Shared Function GetScoredCardForUniversitySchoolHospital(ByVal appName As String, ByVal strConnDB As String, _
ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_UniversitySchoolHospital"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Shared Function GetScoredCardForMortgageBonds(ByVal appName As String, ByVal strConnDB As String, _
        ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
        ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
        ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_MortgageBonds"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Shared Function GetScoredCardForLeaseFinance(ByVal appName As String, ByVal strConnDB As String, _
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_LeaseFinance"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Shared Function IsATemplateFile(ByVal appName As String, ByVal strConnDB As String, _
        ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
        ByVal sessionKey As String, ByVal filename As String) As Boolean

        Try
            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "inv_scorecard_template_t"
                .FilterConditions(0) = "scorecard_template_file_nm = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Shared Function GetScoredcardForGenericLargeCorporate(ByVal appName As String, ByVal strConnDB As String, _
ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer, _
ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_GenericLargeCorporate"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
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

    Public Shared Sub UpdateData(ByRef Arguments As sGenericTableRequestArguments, ByVal SessionKey As String, Optional ByVal GetSecurity As Boolean = False)
        Dim iCount As Integer
        Dim strErrors As String = ""
        Dim objServices As New sGenericTable
        Dim objResultBody As New ResultBody

        If Not GetSecurity AndAlso Arguments.TableNames.Length > 0 Then
            ReDim Arguments.GetSecurityInfo(Arguments.TableNames.Length)
        End If

        objResultBody = ServiceFunctions.Execute(objServices, "UpdateTables", Arguments, SessionKey, ServiceFunctions.TransactionLevel.DTSTransaction)

        If objResultBody.ResultStatus.ReturnCode = "Error" Then
            For iCount = 0 To objResultBody.ResultStatus.MessageCount - 1
                strErrors &= objResultBody.ResultStatus.ResultMessage(iCount).MessageDescription & ControlChars.CrLf
            Next
            Throw New ApplicationException(strErrors)
        End If
    End Sub


    Public Shared Function GetScoredcardForSmallMediumEnterprises(ByVal appName As String, ByVal strConnDB As String,
    ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String,
    ByVal sessionKey As String, ByVal filename As String, ByRef ID As Integer,
    ByRef name As String, ByRef GICS_Sector As String, ByRef GICS_Industry_Group As String, ByRef Lehmans_Sector As String) As String

        Try
            If IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename) Then
                ID = 0
                name = ""
                Lehmans_Sector = ""
                GICS_Sector = ""
                GICS_Industry_Group = ""
                Return "OK"
            End If

            Dim ds As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(appName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = ds
                .ConnectionDatabase = strConnDB
                .ConnectionUser = user
                .ConnectionPwd = password
                .TableNames(0) = "cr_interface_PD_scorecard_predata_SmallMediumEnterprises"
                .FilterConditions(0) = "Filename = '" + filename + "'"

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = strConnAuditDB
                .AuditConnectionUsers(0) = user
                .AuditConnectionPwds(0) = password
            End With

            GetData(objArguments, sessionKey)
            Dim tb As DataTable = ds.Tables(0)
            If tb.Rows.Count = 0 Then
                Return "Can not find PD scorecard information for " + filename
            Else
                Dim row As DataRow = tb.Rows(0)
                ID = CType(row.Item("Entity_ID"), Integer)
                name = CType(row.Item("Entity_Name"), String)
                Lehmans_Sector = CType(row.Item("Lehmans_Sector"), String)
                GICS_Sector = CType(row.Item("GICS_Sector"), String)
                GICS_Industry_Group = CType(row.Item("GICS_Industry_Group"), String)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
End Class
