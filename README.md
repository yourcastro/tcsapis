Imports System.Xml
Imports IT.INV.Service.Base


Public Class ClsPDScoreCardProcesor
    Private mFileID As Integer
    Private mFilename As String
    Private mUserID As String
    Private mAppName As String

    Private mStrConnDB As String
    Private mStrDBUser As String
    Private mStrDBPwd As String
    Private mStrConnAuditDB As String
    Private mStrSessionKey As String

    Private mXML As String


    Sub New(ByVal fileID As Integer, ByVal filename As String, ByVal userID As String, ByVal appName As String, ByVal strConnDB As String, _
        ByVal user As String, ByVal password As String, ByVal strConnAuditDB As String, _
        ByVal sessionKey As String, ByVal xml As String)

        mFileID = fileID
        mFilename = filename
        mUserID = userID
        mAppName = appName

        mStrConnDB = strConnDB
        mStrDBUser = user
        mStrDBPwd = password
        mStrConnAuditDB = strConnAuditDB
        mStrSessionKey = sessionKey

        mXML = xml
    End Sub

    'Public Sub LoadXML(ByVal xml As String)
    '    Try
    '        mXML = My.Computer.FileSystem.ReadAllText("C:\Documents and Settings\ak68\My Documents\Visual Studio 2005\CreditRiskPortalWebService\Bin\Test.xml")
    '        Dim str As String

    '        Dim warningMsg As String = ""
    '        Dim i As Int32
    '        Dim xmlDoc As New XmlDocument()
    '        xmlDoc.LoadXml(mXML)
    '        Dim nodes As XmlNodeList = xmlDoc.GetElementsByTagName("Data")
    '        Dim node As XmlNode
    '        If nodes.Count > 0 Then
    '            Dim nodeData As XmlNode = nodes.Item(0)
    '            If nodeData.HasChildNodes Then
    '                For i = 0 To nodeData.ChildNodes.Count - 1
    '                    node = nodeData.ChildNodes(i)
    '                    str = node.Name
    '                    str = node.Attributes("value").InnerText
    '                Next i
    '            End If

    '        End If
    '        i = 0

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Sub




    Public Sub WriteDB()
        Dim msg As String = "OK"
        Try
            If ClsPDScoreCardFunctions.IsATemplateFile(mAppName, mStrConnDB, mStrDBUser, mStrDBPwd, mStrConnAuditDB, mStrSessionKey, mFilename) Then
                Return
            End If

            Dim m_dsData As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(mAppName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = m_dsData
                .ConnectionDatabase = mStrConnDB
                .ConnectionUser = mStrDBUser
                .ConnectionPwd = mStrDBPwd

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = mStrConnAuditDB
                .AuditConnectionUsers(0) = mStrDBUser
                .AuditConnectionPwds(0) = mStrDBPwd

                .TableNames(0) = "inv_party_entity_scorecard_factors_t"
                .FilterConditions(0) = "party_entity_scorecard_file_nm = '" + mFilename + "'"
            End With

            Dim seesionKey As String = mStrSessionKey
            ClsPDScoreCardFunctions.GetData(objArguments, seesionKey)
            Dim tb As DataTable = m_dsData.Tables(0)
            'If tb.Rows.Count = 0 Then
            '    Dim row As DataRow = tb.NewRow()
            '    row.Item("party_entity_scorecard_factors_id") = mFileID
            '    row.Item("party_entity_scorecard_file_nm") = mFilename
            '    row.Item("create_dt_tm") = DateTime.Now
            '    row.Item("create_process_id") = mUserID
            '    SetRowValues(row)
            '    tb.Rows.Add(row)
            'Else
            '    Dim row As DataRow = tb.Rows(0)
            '    SetRowValues(row)
            'End If
            'ClsPDScoreCardFunctions.UpdateData(objArguments, seesionKey)
            If tb.Rows.Count > 0 Then
                Dim row As DataRow = tb.Rows(0)
                SetRowValues(row)
                ClsPDScoreCardFunctions.UpdateData(objArguments, seesionKey)
                WriteScoreCardTable(row("party_entity_scorecard_factors_id"))
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub WriteScoreCardTable(ByVal ScorecardFactorsID As Integer)
        Dim msg As String = "OK"
        Try
            Dim m_dsData As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(mAppName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = m_dsData
                .ConnectionDatabase = mStrConnDB
                .ConnectionUser = mStrDBUser
                .ConnectionPwd = mStrDBPwd

                ReDim .AuditConnectionDatabases(0)
                ReDim .AuditConnectionUsers(0)
                ReDim .AuditConnectionPwds(0)
                .AuditConnectionDatabases(0) = mStrConnAuditDB
                .AuditConnectionUsers(0) = mStrDBUser
                .AuditConnectionPwds(0) = mStrDBPwd

                .TableNames(0) = "inv_party_entity_scorecard_t"
                .FilterConditions(0) = "party_entity_scorecard_factors_id = " + ScorecardFactorsID.ToString
            End With

            Dim seesionKey As String = mStrSessionKey
            ClsPDScoreCardFunctions.GetData(objArguments, seesionKey)
            Dim tb As DataTable = m_dsData.Tables(0)
            If tb.Rows.Count > 0 Then
                Dim row As DataRow = tb.Rows(0)
                row.Item("last_update_process_id") = mUserID
                row.Item("last_update_dt_tm") = DateTime.Now
            End If
            ClsPDScoreCardFunctions.UpdateData(objArguments, seesionKey)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub SetRowValues(ByVal row As DataRow)
        Try
            Dim i As Int32
            Dim xmlDoc As New XmlDocument()
            xmlDoc.LoadXml(mXML)
            Dim nodes As XmlNodeList = xmlDoc.GetElementsByTagName("Data")
            Dim node As XmlNode
            If nodes.Count > 0 Then
                Dim nodeData As XmlNode = nodes.Item(0)
                If nodeData.HasChildNodes Then
                    For i = 0 To nodeData.ChildNodes.Count - 1
                        node = nodeData.ChildNodes(i)
                        row.Item(node.Name) = node.Attributes("value").InnerText
                    Next i
                End If

            End If

            'Dim cardItem As ClsScoreCardItem
            'Dim i As Integer

            'For i = 0 To mItems.Count - 1
            'cardItem = CType(mItems(i), ClsScoreCardItem)
            'row.Item(cardItem.mField) = cardItem.mValue
            'Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Function TestDBAccess() As String
        Dim msg As String = "OK"
        Try
            Dim m_dsData As New DataSet
            Dim objArguments As New sGenericTableRequestArguments(mAppName)
            With objArguments
                ReDim .TableNames(0)
                ReDim .FilterConditions(0)
                .dsData = m_dsData
                .ConnectionDatabase = mStrConnDB
                .ConnectionUser = mStrDBUser
                .ConnectionPwd = mStrDBPwd
                '.ConnectionDatabase = "Server=SQLRG1D.ca.sunlife\SLAV34D;Database=VRRPartyEntityDatastore;"
                '.ConnectionUser = "RRWebUser"
                '.ConnectionPwd = "JASUR/v70o8="
                .TableNames(0) = "inv_pdscorecard_file_t"

