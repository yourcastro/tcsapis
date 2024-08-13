Imports System.Configuration.ConfigurationManager
Imports it.INV.Service.Base
Imports it.INV.Presentation.Web.Base
Imports it.INV.Presentation.Web.Base.Constants
Imports it.INV.Presentation.Web.Base.CommonFunctions
Imports it.INV.Presentation.Web.Shared.PresentationFunctions
Imports System.Data
Imports it.INV.Presentation.Web.Shared

Partial Class SearchEntities
    Inherits System.Web.UI.Page

    Private Const TABLE_NAME_ENTITY As String = "inv_party_entity_t"
    Private Const TABLE_SP_NAME_ENTITY As String = "inv_party_entity_search_sp"
    Private Const TABLE_LIST_COUNTRY As String = "cr_interface_display_country_v"
    Private Const TABLE_LIST_REG_CLS As String = "regulatory_class_filtered_sp"
    Private Const TABLE_LIST_REG_SUB_CLS As String = "regulatory_sub_class_filtered_sp"
    Private Const TABLE_LIST_STATUS As String = "cr_interface_display_status_v"
    Private Const TABLE_LIST_BLOOMBERG As String = "Classifications_filtered_Bloomberg_sp"
    Private Const TABLE_LIST_GICS As String = "Classifications_filtered_GICS_sp"
    Private Const TABLE_LIST_LEHMAN As String = "Classifications_filtered_Lehman_sp"
    Private Const TABLE_LIST_MSCI As String = "Classifications_filtered_MSCI_sp"
    Private Const CONFIG_NUM_ROWS_PER_PAGE As String = "SearchRowsPerPage"
    Private Const CONFIG_DISPLAY_DETAIL_PAGES As String = "SearchDetailPages"
    Private Const SESSION_SEARCH_FILTER_ENT As String = "FilterEnt"
    Private Const SESSION_SEARCH_FILTER_LN As String = "FilterLN"
    Private Const SESSION_SEARCH_FILTER_STATUS As String = "FilterStatus"
    Private Const SESSION_SEARCH_FILTER_REGCLS As String = "FilterRegCls"
    Private Const SESSION_SEARCH_FILTER_REGSUBCLS As String = "FilterRegSubCls"
    Private Const SESSION_SEARCH_FILTER_DOMCNTRY As String = "FilterDomCntry"
    Private Const SESSION_SEARCH_FILTER_ULTCNTRY As String = "FilterUltCntry"
    Private Const SESSION_SEARCH_FILTER_BLOOMBERG As String = "FilterBloomberg"
    Private Const SESSION_SEARCH_FILTER_GICS As String = "FilterGICS"
    Private Const SESSION_SEARCH_FILTER_LEHMAN As String = "FilterLehman"
    Private Const SESSION_SEARCH_FILTER_MSCI As String = "FilterMSCI"
    Private Const SESSION_SORT_ORDER As String = "SortOrder"
    Private Const SESSION_CURR_PAGE As String = "SECurrPage"
    Private Const DG_CELL_EDIT As Integer = 0
    Private Const DG_CELL_SELECT As Integer = 1
    Private Const DG_CELL_STATUS As Integer = 6
    Private Const DG_CELL_VIEW As Integer = 2

    Private Const SESSION_PD_CURR_PAGE As String = "PDCurrPage"
    Private Const SESSION_PD_SEARCH_FILTER_ENTITY_ID As String = "FilterPDEntityID"
    Private Const SESSION_PD_SEARCH_FILTER_LEGAL_Name As String = "FilterPDLegalName"
    Private Const SESSION_PD_SEARCH_FILTER_ENTITY_STATUS As String = "FilterPDEntityStatus"
    Private Const SESSION_PD_SEARCH_FILTER_DOMICILE_COUNTRY As String = "FilterPDDomicileCountry"
    Private Const SESSION_PD_SEARCH_FILTER_GICS As String = "FilterPDGICS"
    Private Const SESSION_PD_SEARCH_FILTER_LEHMAN As String = "FilterPDLehman"
    Private Const SESSION_PD_SEARCH_FILTER_SCORECARD_TYPE As String = "FilterPDType"
    Private Const SESSION_PD_SEARCH_FILTER_SCORECARD_STATUS As String = "FilterPDStatus"
    Private Const SESSION_PD_SEARCH_FILTER_PD_RATING As String = "FilterPDRating"
    Private Const SESSION_PD_SEARCH_FILTER_START_EFFECTIVE_DATE As String = "FilterPDStartEffDate"
    Private Const SESSION_PD_SEARCH_FILTER_APPROVER As String = "FilterPDApprover"
    Private Const SESSION_PD_SEARCH_FILTER_APPROVAL_DATE As String = "FilterPDApprovalDate"
    Private Const SESSION_PD_SEARCH_FILTER_ANALYST As String = "FilterPDAnalyst"
    Private Const SESSION_PD_SEARCH_FILTER_ASSET_CLASS As String = "FilterPDAssetClass"
    Private Const SESSION_SEARCH_FILTER_ROLE_TYPE As String = "FilterPDRoleType"

    Private m_dsData As New DataSet
    Private m_dsListDataClass As New DataSet
    Private m_dsListDataRegClass As New DataSet
    Private m_objSessionItems As SessionItems

#Region "Protected Virtual Properties"
    Protected Property TotalRows() As Integer
        Get
            Dim i As Integer = 0
            If Not ViewState("TotalRows") Is Nothing Then
                i = ViewState("TotalRows")
            End If
            Return i
        End Get
        Set(ByVal value As Integer)
            ViewState("TotalRows") = value
        End Set
    End Property

    Protected ReadOnly Property NumPages() As Integer
        Get
            Return Decimal.Floor(TotalRows / AppSettings.Get(CONFIG_NUM_ROWS_PER_PAGE)) + 1
        End Get
    End Property
#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckSession(Request, Response)

        Try
            m_objSessionItems = New SessionItems
            m_objSessionItems.SessionKey = Request.Cookies(SESSION_COOKIE_CR).Value
            LoadSession(m_objSessionItems, Session, Request.Cookies(SESSION_COOKIE_CR).Value, Response)
            mhTop.MenuXmlData = m_objSessionItems.MenuXMLData

            Dim blnFirstLoad As Boolean = m_objSessionItems.ItemValues(SESSION_CURR_PAGE) = ""
            btnNew.Visible = CType(m_objSessionItems.ItemValues(ACCESS_ENTITY_RW), Boolean)

            If Page.IsPostBack Then
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_ENT, txtEntity.Text)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_LN, txtLegalNm.Text)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_STATUS, ddlStatus.SelectedValue)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_REGCLS, ddlRegCls.SelectedValue)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_REGSUBCLS, ddlRegSubCls.SelectedItem.Text)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_DOMCNTRY, ddlDomCntry.SelectedValue)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_ULTCNTRY, ddlUltCntry.SelectedValue)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_BLOOMBERG, ddlBloomberg.SelectedValue)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_GICS, ddlGICS.SelectedValue)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_LEHMAN, ddlLehmans.SelectedValue)
                m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_MSCI, ddlMSCI.SelectedValue)
                StoreSession(m_objSessionItems, Session, m_objSessionItems.SessionKey)
            Else
                If Not CType(m_objSessionItems.ItemValues(ACCESS_ENTITY_R), Boolean) Then
                    Response.Redirect("Home.aspx", False)
                End If

                btnFilter.Attributes.Add("onclick", "WaitCursor();")
                btnNew.Attributes.Add("onclick", "WaitCursor();")

                If m_objSessionItems.ItemValues(SESSION_SORT_ORDER) = "" Then
                    m_objSessionItems.ItemValues.Set(SESSION_SORT_ORDER, "EntityID")
                End If
                If m_objSessionItems.ItemValues(SESSION_CURR_PAGE) = "" Then
                    m_objSessionItems.ItemValues.Set(SESSION_CURR_PAGE, "0")
                End If
                GridView1.PageIndex = m_objSessionItems.ItemValues(SESSION_CURR_PAGE)
                StoreSession(m_objSessionItems, Session, m_objSessionItems.SessionKey)
                SetupStaticLists()
                SetupRegulatoryClassLists(True, True, "", "")
                SetupClassLists(True, True, True, True, "", "", "", "")

                If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_REGCLS) <> "" Then
                    ddlRegCls.SelectedValue = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_REGCLS)
                End If
                If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_REGSUBCLS) <> "" Then
                    ddlRegSubCls.SelectedValue = ddlRegSubCls.Items.FindByText(m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_REGSUBCLS)).Value
                End If

                If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_BLOOMBERG) <> "All" Then
                    ddlBloomberg.SelectedValue = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_BLOOMBERG)
                End If
                If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_GICS) <> "All" Then
                    ddlGICS.SelectedValue = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_GICS)
                End If
                If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_LEHMAN) <> "All" Then
                    ddlLehmans.SelectedValue = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_LEHMAN)
                End If

                If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_MSCI) <> "All" Then
                    ddlMSCI.SelectedValue = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_MSCI)
                End If

                If Not blnFirstLoad Then
                    GetDisplayData(True)
                    SetupGrid()
                    lblNoRecs.Visible = GridView1.Rows.Count = 0
                    DivBorders.Visible = True
                End If
            End If

            'Set up the navigation controls (needs to be run on every page load since controls are added)
            SetupNav(GridView1.PageIndex)
            If Not blnFirstLoad Then
                FormatNav(GridView1.PageIndex)
            End If
            If txtEntity.Text <> "" AndAlso txtLegalNm.Text = "" Then
                txtEntity.Focus()
            Else
                txtLegalNm.Focus()
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            GridView1.PageIndex = 0
            GetDisplayData(True)
            SetupNav(GridView1.PageIndex)
            FormatNav(GridView1.PageIndex)
            SetupGrid()
            lblNoRecs.Visible = GridView1.Rows.Count = 0
            DivBorders.Visible = True
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Response.Redirect("AddEntity.aspx?0=" & Encrypt("ID=-1"), False)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Try
            Dim iEntityID As Integer
            Select Case e.CommandName
                Case "OpenEdit"
                    iEntityID = GridView1.DataKeys(e.CommandArgument).Value
                    Response.Redirect("AddEntity.aspx?0=" & Encrypt("ID=" & iEntityID), False)
                Case "ParentRelationship"
                    iEntityID = GridView1.DataKeys(e.CommandArgument).Value
                    Response.Redirect("ChildParent.aspx?0=" & Encrypt("ID=" & iEntityID), False)
                Case "SortID"
                    GridView1.Sort("EntityID", SortDirection.Ascending)
                Case "SortNm"
                    GridView1.Sort("LegalName", SortDirection.Ascending)
                Case "ScorecardView"
                    iEntityID = GridView1.DataKeys(e.CommandArgument).Value
                    OpenPDScorecardMaintenancePage(iEntityID)
                Case "Page"
                    'do nothing
                Case Else
                    Throw New Exception("Unhandled Command")
            End Select
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Try
            Dim row As GridViewRow = e.Row
            Dim tb As DataTable = GridView1.DataSource
            If row.RowIndex >= 0 Then
                If tb.Rows(row.RowIndex).Item("ScorecardNumber") = 0 Then
                    row.Cells(DG_CELL_VIEW).Controls(0).Visible = False
                End If
            End If
            'Dim str As String = row.Cells(5).Text
            'If str = "0" Then
            '    row.Cells(2).Controls(0).Visible = False
            'End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub GridView1_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles GridView1.Sorting
        Try
            Dim iOldIndex As Integer = GridView1.PageIndex
            GridView1.PageIndex = 0
            m_objSessionItems.ItemValues(SESSION_SORT_ORDER) = e.SortExpression
            StoreSession(m_objSessionItems, Session, m_objSessionItems.SessionKey)
            GetDisplayData()
            FormatNav(GridView1.PageIndex, iOldIndex)
            SetupGrid()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    'This is the event handler called by any of the pagination links
    Protected Sub lnkNav_ChangePage(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim iOldPage As Integer = GridView1.PageIndex
            Dim strCaption As String = CType(sender, LinkButton).Text

            If strCaption = "<" Then
                GridView1.PageIndex = GridView1.PageIndex - 1
                If FindControl("lnkNav" & GridView1.PageIndex + 1) Is Nothing Then
                    iOldPage = -1
                    SetupNav(GridView1.PageIndex)
                End If
            ElseIf strCaption = ">" Then
                GridView1.PageIndex = GridView1.PageIndex + 1
                If FindControl("lnkNav" & GridView1.PageIndex + 1) Is Nothing Then
                    iOldPage = -1
                    SetupNav(GridView1.PageIndex)
                End If
            ElseIf strCaption.Contains("-") Then
                iOldPage = -1
                GridView1.PageIndex = strCaption.Substring(0, strCaption.IndexOf("-")) - 1
                SetupNav(GridView1.PageIndex)
            Else
                GridView1.PageIndex = strCaption - 1
            End If

            FormatNav(GridView1.PageIndex, iOldPage)
            GetDisplayData()
            SetupGrid()
            m_objSessionItems.ItemValues(SESSION_CURR_PAGE) = GridView1.PageIndex
            StoreSession(m_objSessionItems, Session, m_objSessionItems.SessionKey)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    'Does any setup required after page is changed. Can be called on initialize too.
    Protected Sub FormatNav(ByVal CurrentPage As Integer, Optional ByVal OldPage As Integer = -1)
        If NumPages > 1 Then
            If OldPage <> -1 Then
                CType(FindControl("lnkNav" & OldPage + 1), System.Web.UI.WebControls.LinkButton).CssClass = "NavNS"
            End If

            If CurrentPage = 0 Then
                CType(FindControl("lnkNavP"), System.Web.UI.WebControls.LinkButton).Enabled = False
            Else
                CType(FindControl("lnkNavP"), System.Web.UI.WebControls.LinkButton).Enabled = True
            End If
            If CurrentPage = NumPages - 1 Then
                CType(FindControl("lnkNavN"), System.Web.UI.WebControls.LinkButton).Enabled = False
            Else
                CType(FindControl("lnkNavN"), System.Web.UI.WebControls.LinkButton).Enabled = True
            End If

            CType(FindControl("lnkNav" & CurrentPage + 1), System.Web.UI.WebControls.LinkButton).CssClass = "NavS"
        End If
    End Sub

    'Steps to create navigation controls (needs to be called on load of each postback)
    Protected Sub SetupNav(ByVal CurrentPage As Integer)
        Dim lbl As Label
        Dim lnkNav As LinkButton
        Dim iCount, iCountPg, iStart, iEnd As Integer
        Dim iDetailPages As Integer = AppSettings.Get(CONFIG_DISPLAY_DETAIL_PAGES)
        Dim iNumGroups As Integer = Decimal.Ceiling(NumPages / iDetailPages)
        Dim iRowsPerPg As Integer = AppSettings.Get(CONFIG_NUM_ROWS_PER_PAGE)

        'Clear old controls in case we're calling more than once
        For iCount = 0 To pnlNav.Controls.Count - 1
            pnlNav.Controls.RemoveAt(0)
        Next

        If NumPages > 1 Then
            'Add the "prev" link
            lbl = New Label
            lbl.Text = "&nbsp;"
            lnkNav = New LinkButton
            lnkNav.Text = "<"
            lnkNav.ID = "lnkNavP"
            lnkNav.EnableViewState = False
            lnkNav.Attributes.Add("onclick", "WaitCursor();")
            AddHandler lnkNav.Click, AddressOf lnkNav_ChangePage
            pnlNav.Controls.Add(lnkNav)
            pnlNav.Controls.Add(lbl)

            'Add the "next" link
            lbl = New Label
            lbl.Text = "&nbsp;"
            lnkNav = New LinkButton
            lnkNav.Text = ">"
            lnkNav.ID = "lnkNavN"
            lnkNav.EnableViewState = False
            lnkNav.Attributes.Add("onclick", "WaitCursor();")
            AddHandler lnkNav.Click, AddressOf lnkNav_ChangePage
            pnlNav.Controls.Add(lnkNav)
            pnlNav.Controls.Add(lbl)

            If NumPages > iDetailPages Then
                'Add numbered buttons for each page, buttons outside the current range will be grouped ranges.
                For iCountPg = 0 To iNumGroups - 1
                    iStart = iCountPg * iDetailPages
                    iEnd = iStart + iDetailPages - 1
                    If iEnd > NumPages - 1 Then
                        iEnd = NumPages - 1
                    End If

                    If CurrentPage >= iStart AndAlso CurrentPage <= iEnd Then
                        'This is the current page group, add detailed page numbers
                        For iCount = iStart To iEnd
                            lbl = New Label
                            lbl.Text = "&nbsp;"
                            lnkNav = New LinkButton
                            lnkNav.Text = iCount + 1
                            lnkNav.ID = "lnkNav" & iCount + 1
                            lnkNav.EnableViewState = False
                            lnkNav.Attributes.Add("onclick", "WaitCursor();")
                            AddHandler lnkNav.Click, AddressOf lnkNav_ChangePage
                            pnlNav.Controls.Add(lnkNav)
                            pnlNav.Controls.Add(lbl)
                        Next
                    Else
                        'Not a selected page group, just add a single link for this page group
                        lbl = New Label
                        lbl.Text = "&nbsp;"
                        lnkNav = New LinkButton
                        lnkNav.Text = iStart + 1 & "-" & iEnd + 1
                        lnkNav.ID = "lnkNavG" & iCountPg + 1
                        lnkNav.EnableViewState = False
                        lnkNav.Attributes.Add("onclick", "WaitCursor();")
                        AddHandler lnkNav.Click, AddressOf lnkNav_ChangePage
                        pnlNav.Controls.Add(lnkNav)
                        pnlNav.Controls.Add(lbl)
                    End If
                Next
            Else
                'Add numbered buttons for each page
                For iCount = 0 To NumPages - 1
                    lbl = New Label
                    lbl.Text = "&nbsp;"
                    lnkNav = New LinkButton
                    lnkNav.Text = iCount + 1
                    lnkNav.ID = "lnkNav" & iCount + 1
                    lnkNav.EnableViewState = False
                    lnkNav.Attributes.Add("onclick", "WaitCursor();")
                    AddHandler lnkNav.Click, AddressOf lnkNav_ChangePage
                    pnlNav.Controls.Add(lnkNav)
                    pnlNav.Controls.Add(lbl)
                Next
            End If
        End If
    End Sub

    Private Sub SetupStaticLists()
        GetListData()
        SetupDDL(ddlDomCntry, m_dsData.Tables(TABLE_LIST_COUNTRY), False, True)
        SetupDDL(ddlUltCntry, m_dsData.Tables(TABLE_LIST_COUNTRY), False, True)
        SetupDDL(ddlStatus, m_dsData.Tables(TABLE_LIST_STATUS), False, True)

        If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_ENT) <> "" Then
            txtEntity.Text = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_ENT)
        End If
        If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_LN) <> "" Then
            txtLegalNm.Text = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_LN)
        End If
        If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_STATUS) <> "" Then
            ddlStatus.SelectedValue = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_STATUS)
        Else
            ddlStatus.SelectedValue = STATUS_ACTIVE
        End If
        If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_DOMCNTRY) <> "All" Then
            ddlDomCntry.SelectedValue = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_DOMCNTRY)
        End If
        If m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_ULTCNTRY) <> "All" Then
            ddlUltCntry.SelectedValue = m_objSessionItems.ItemValues(SESSION_SEARCH_FILTER_ULTCNTRY)
        End If
    End Sub

    Private Sub SetupGrid()
        Dim iCount As Integer

        GridView1.DataSource = m_dsData.Tables(TABLE_NAME_ENTITY)
        GridView1.DataBind()

        If GridView1.Rows.Count > 0 Then
            If m_objSessionItems.ItemValues(SESSION_SORT_ORDER) = "EntityID" Then
                CType(GridView1.HeaderRow.FindControl("lnkSortID"), LinkButton).CssClass = "HeaderS"
                CType(GridView1.HeaderRow.FindControl("lnkSortNm"), LinkButton).CssClass = "HeaderNS"
            Else
                CType(GridView1.HeaderRow.FindControl("lnkSortID"), LinkButton).CssClass = "HeaderNS"
                CType(GridView1.HeaderRow.FindControl("lnkSortNm"), LinkButton).CssClass = "HeaderS"
            End If
            CType(GridView1.HeaderRow.FindControl("lnkSortID"), LinkButton).Attributes.Add("onclick", "WaitCursor();")
            CType(GridView1.HeaderRow.FindControl("lnkSortNm"), LinkButton).Attributes.Add("onclick", "WaitCursor();")
        End If

        For iCount = 0 To GridView1.Rows.Count - 1
            If GetStatusKeyFromDesc(GridView1.Rows(iCount).Cells(DG_CELL_STATUS).Text, m_objSessionItems.SessionKey, Cache) <> STATUS_ACTIVE Then
                CType(GridView1.Rows(iCount).Cells(DG_CELL_SELECT).Controls(0), Button).Visible = False
            Else
                CType(GridView1.Rows(iCount).Cells(DG_CELL_SELECT).Controls(0), Button).Attributes.Add("onclick", "WaitCursor();")
            End If
            CType(GridView1.Rows(iCount).Cells(DG_CELL_EDIT).Controls(0), Button).Attributes.Add("onclick", "WaitCursor();")
        Next
    End Sub

    Private Sub HandleError(ByVal ex As Exception)
        DisplayMessage(ex, msgError)
    End Sub

    Private Sub GetDisplayData(Optional ByVal blnGetTotals As Boolean = False)
        Dim objArguments As New sGenericTableRequestArguments(AppSettings.Get(CONFIG_APPNAME))

        With objArguments
            ReDim .StoredProc(0)
            ReDim .TableNames(0)
            ReDim .SPParameters(0, 15)

            .dsData = m_dsData
            .ConnectionDatabase = AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION)
            .ConnectionUser = AppSettings.Get(CONFIG_DB_USER)
            .ConnectionPwd = AppSettings.Get(CONFIG_DB_PWD)

            .TableNames(0) = TABLE_NAME_ENTITY
            .StoredProc(0) = TABLE_SP_NAME_ENTITY
            .SPParameters(0, 0) = New SqlClient.SqlParameter("pageIndex", GridView1.PageIndex)
            .SPParameters(0, 1) = New SqlClient.SqlParameter("rowsPerPage", AppSettings.Get(CONFIG_NUM_ROWS_PER_PAGE))
            .SPParameters(0, 2) = New SqlClient.SqlParameter("SortKey", m_objSessionItems.ItemValues(SESSION_SORT_ORDER))
            .SPParameters(0, 3) = New SqlClient.SqlParameter("GetTotal", System.Math.Abs(CType(blnGetTotals, Integer)))
            .SPParameters(0, 4) = New SqlClient.SqlParameter("EntityID", txtEntity.Text.Trim().TrimStart("0").Replace("*", "%") & IIf(txtEntity.Text.Trim() = "", "", "%"))
            .SPParameters(0, 5) = New SqlClient.SqlParameter("LegalNm", txtLegalNm.Text.Replace("*", "%") & IIf(txtLegalNm.Text = "", "", "%"))
            .SPParameters(0, 6) = New SqlClient.SqlParameter("DomCountryCode", IIf(ddlDomCntry.SelectedItem.Text = "", "NULL", IIf(ddlDomCntry.SelectedItem.Text = "All", "", ddlDomCntry.SelectedValue)))
            .SPParameters(0, 7) = New SqlClient.SqlParameter("UltCountryCode", IIf(ddlUltCntry.SelectedItem.Text = "", "NULL", IIf(ddlUltCntry.SelectedItem.Text = "All", "", ddlUltCntry.SelectedValue)))
            .SPParameters(0, 8) = New SqlClient.SqlParameter("RegulatoryClass", IIf(ddlRegCls.SelectedValue = "All", "", ddlRegCls.SelectedValue))
            .SPParameters(0, 9) = New SqlClient.SqlParameter("RegulatorySubClass", IIf(ddlRegSubCls.SelectedItem.Text = "All", "", ddlRegSubCls.SelectedItem.Text))
            .SPParameters(0, 10) = New SqlClient.SqlParameter("Bloomberg", IIf(ddlBloomberg.SelectedItem.Text = "", "NULL", IIf(ddlBloomberg.SelectedItem.Text = "All", "", ddlBloomberg.SelectedValue)))
            .SPParameters(0, 11) = New SqlClient.SqlParameter("GICS", IIf(ddlGICS.SelectedItem.Text = "", "NULL", IIf(ddlGICS.SelectedItem.Text = "All", "", ddlGICS.SelectedValue)))
            .SPParameters(0, 12) = New SqlClient.SqlParameter("Lehman", IIf(ddlLehmans.SelectedItem.Text = "", "NULL", IIf(ddlLehmans.SelectedItem.Text = "All", "", ddlLehmans.SelectedValue)))
            .SPParameters(0, 13) = New SqlClient.SqlParameter("MSCI", IIf(ddlMSCI.SelectedItem.Text = "", "NULL", IIf(ddlMSCI.SelectedItem.Text = "All", "", ddlMSCI.SelectedValue)))
            .SPParameters(0, 14) = New SqlClient.SqlParameter("Status", IIf(ddlStatus.SelectedValue = "All", "", ddlStatus.SelectedValue))
            .SPParameters(0, 15) = New SqlClient.SqlParameter("totalRows", SqlDbType.Int, 4)
            .SPParameters(0, 15).Direction = ParameterDirection.Output
        End With
        GetData(objArguments, m_objSessionItems.SessionKey)

        If m_dsData.Tables(TABLE_NAME_ENTITY) Is Nothing Then
            Throw New ApplicationException("Problem retrieving Entity data.")
        End If
        If blnGetTotals Then
            TotalRows = objArguments.SPParameters(0, 15).Value
        End If
    End Sub

    Private Sub GetListData()
        Dim objArguments As New sGenericTableRequestArguments(AppSettings.Get(CONFIG_APPNAME))

        With objArguments
            ReDim .TableNames(1)
            ReDim .OrderBys(1)

            .dsData = m_dsData
            .ConnectionDatabase = AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION)
            .ConnectionUser = AppSettings.Get(CONFIG_DB_USER)
            .ConnectionPwd = AppSettings.Get(CONFIG_DB_PWD)

            .TableNames(0) = TABLE_LIST_COUNTRY
            .TableNames(1) = TABLE_LIST_STATUS
            '.TableNames(1) = TABLE_LIST_REG_CLS
            ' .TableNames(2) = TABLE_LIST_REG_SUB_CLS
            '.TableNames(4) = TABLE_LIST_BLOOMBERG
            '.TableNames(5) = TABLE_LIST_GICS
            '.TableNames(6) = TABLE_LIST_LEHMAN
            ' 12-Sep-2011 - Krishna - Change the order of the country list - Start
            .OrderBys(0) = "SortOrder, DisplayVal"
            ' 12-Sep-2011 - Krishna - Change the order of the country list - End
            .OrderBys(1) = "DisplayVal"
            '.OrderBys(2) = "DisplayVal"
            '.OrderBys(3) = "DisplayVal"
            '.OrderBys(4) = "DisplayVal"
            '.OrderBys(5) = "DisplayVal"
            '.OrderBys(6) = "DisplayVal"
        End With
        GetData(objArguments, m_objSessionItems.SessionKey)
    End Sub

    Protected Sub ddlRegCls_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles ddlRegCls.SelectedIndexChanged

        If ddlRegCls.SelectedValue <> "" Then
            'filter  Reg Sub Cls             
            SetupRegulatoryClassLists(False, True, ddlRegCls.SelectedValue, ddlRegSubCls.SelectedItem.Text)
        End If

    End Sub

    Protected Sub ddlRegSubCls_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles ddlRegSubCls.SelectedIndexChanged

        If ddlRegSubCls.SelectedValue <> "" Then
            'filter  Reg Sub Cls
            SetupRegulatoryClassLists(True, False, ddlRegCls.SelectedValue, ddlRegSubCls.SelectedItem.Text)
        End If

    End Sub

    Protected Sub ddlLehmans_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles ddlLehmans.SelectedIndexChanged

        If ddlLehmans.SelectedValue <> "" AndAlso ddlLehmans.SelectedValue <> CLASS_NONRATED_LEHMANS Then
            'filter GICS & Bloomberg for the current Lehman selection
            SetupClassLists(True, True, False, True, ddlBloomberg.SelectedValue, ddlGICS.SelectedValue, ddlLehmans.SelectedValue, ddlMSCI.SelectedValue, True)
        End If

    End Sub

    Protected Sub ddlMSCI_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles ddlMSCI.SelectedIndexChanged

        If ddlMSCI.SelectedValue <> "" AndAlso ddlMSCI.SelectedValue <> CLASS_NONRATED_MSCI Then
            'filter GICS & Bloomberg for the current MSCI selection
            SetupClassLists(True, True, True, False, ddlBloomberg.SelectedValue, ddlGICS.SelectedValue, ddlLehmans.SelectedValue, ddlMSCI.SelectedValue, True)
        End If

    End Sub

    Protected Sub ddlGICS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles ddlGICS.SelectedIndexChanged

        If ddlGICS.SelectedValue <> "" AndAlso ddlGICS.SelectedValue <> CLASS_NONRATED_GICS Then
            'filter Bloomberg & Lehman for the current GICS selection
            SetupClassLists(True, False, True, True, ddlBloomberg.SelectedValue, ddlGICS.SelectedValue, ddlLehmans.SelectedValue, ddlMSCI.SelectedValue, True)
        End If

    End Sub

    Protected Sub ddlBloomberg_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles ddlBloomberg.SelectedIndexChanged

        If ddlBloomberg.SelectedValue <> "" AndAlso ddlBloomberg.SelectedValue <> CLASS_NONRATED_BLOOMBERG Then
            'filter GICS & Lehman for the current Bloomberg selection
            SetupClassLists(False, True, True, True, ddlBloomberg.SelectedValue, ddlGICS.SelectedValue, ddlLehmans.SelectedValue, ddlMSCI.SelectedValue, True)
        End If

    End Sub

    Protected Sub btnRefreshClass_Click(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles btnRefreshClass.Click
        Try

            SetupClassLists(True, True, True, True, "", "", "", "")
            ddlGICS.SelectedValue = "All"
            ddlBloomberg.SelectedValue = "All"
            ddlLehmans.SelectedValue = "All"
            ddlMSCI.SelectedValue = "All"

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub SetupClassLists(ByVal RefreshB As Boolean, ByVal RefreshG As Boolean, ByVal RefreshL As Boolean, ByVal RefreshM As Boolean, ByVal FilterB As String, ByVal FilterG As String, ByVal FilterL As String, ByVal FilterM As String, Optional ByVal ddlSelected As Boolean = False)
        Dim strOldVal As String

        GetClassListData(RefreshB, RefreshG, RefreshL, RefreshM, IIf(FilterB = "All", "", FilterB), IIf(FilterG = "All", "", FilterG), IIf(FilterL = "All", "", FilterL), IIf(FilterM = "All", "", FilterM))
        If RefreshB Then
            strOldVal = ddlBloomberg.SelectedValue
            SetupDDL(ddlBloomberg, m_dsListDataClass.Tables(TABLE_LIST_BLOOMBERG), False, True)
            If strOldVal <> "All" AndAlso Not ddlBloomberg.Items.FindByValue(strOldVal) Is Nothing Then
                ddlBloomberg.SelectedValue = strOldVal
            ElseIf ddlSelected = True Then
                Select Case ddlBloomberg.Items.Count
                    Case 3
                        Dim i As Integer = 0
                        For i = 0 To ddlBloomberg.Items.Count - 1
                            If ddlBloomberg.Items(i).Value <> CLASS_NONRATED_BLOOMBERG And ddlBloomberg.Items(i).Value <> "All" Then
                                ddlBloomberg.Items(i).Selected = True
                            End If
                        Next
                    Case Else
                        ddlBloomberg.SelectedValue = "All"
                End Select
            End If
        End If
        If RefreshG Then
            strOldVal = ddlGICS.SelectedValue
            SetupDDL(ddlGICS, m_dsListDataClass.Tables(TABLE_LIST_GICS), False, True)
            If strOldVal <> "All" AndAlso Not ddlGICS.Items.FindByValue(strOldVal) Is Nothing Then
                ddlGICS.SelectedValue = strOldVal
            ElseIf ddlSelected = True Then
                Select Case ddlGICS.Items.Count
                    Case 3
                        Dim i As Integer = 0
                        For i = 0 To ddlGICS.Items.Count - 1
                            If ddlGICS.Items(i).Value <> CLASS_NONRATED_GICS And ddlGICS.Items(i).Value <> "All" Then
                                ddlGICS.Items(i).Selected = True
                            End If
                        Next
                    Case Else
                        ddlGICS.SelectedValue = "All"
                End Select
            End If
        End If
        If RefreshL Then
            strOldVal = ddlLehmans.SelectedValue
            SetupDDL(ddlLehmans, m_dsListDataClass.Tables(TABLE_LIST_LEHMAN), False, True)
            If strOldVal <> "All" AndAlso Not ddlLehmans.Items.FindByValue(strOldVal) Is Nothing Then
                ddlLehmans.SelectedValue = strOldVal
            ElseIf ddlSelected = True Then
                Select Case ddlLehmans.Items.Count
                    Case 3
                        Dim i As Integer = 0
                        For i = 0 To ddlLehmans.Items.Count - 1
                            If ddlLehmans.Items(i).Value <> CLASS_NONRATED_LEHMANS And ddlLehmans.Items(i).Value <> "All" Then
                                ddlLehmans.Items(i).Selected = True
                            End If
                        Next
                    Case Else
                        ddlLehmans.SelectedValue = "All"
                End Select
            End If
        End If
        If RefreshM Then
            strOldVal = ddlMSCI.SelectedValue
            SetupDDL(ddlMSCI, m_dsListDataClass.Tables(TABLE_LIST_MSCI), False, True)
            If strOldVal <> "All" AndAlso Not ddlMSCI.Items.FindByValue(strOldVal) Is Nothing Then
                ddlMSCI.SelectedValue = strOldVal
            ElseIf ddlSelected = True Then
                Select Case ddlMSCI.Items.Count
                    Case 3
                        Dim i As Integer = 0
                        For i = 0 To ddlMSCI.Items.Count - 1
                            If ddlMSCI.Items(i).Value <> CLASS_NONRATED_MSCI And ddlMSCI.Items(i).Value <> "All" Then
                                ddlMSCI.Items(i).Selected = True
                            End If
                        Next
                    Case Else
                        ddlMSCI.SelectedValue = "All"
                End Select
            End If
        End If
    End Sub



    '------------------------------------------------------------------------------------
    ' Get data for Classification lists
    '------------------------------------------------------------------------------------
    Private Sub GetClassListData(ByVal GetB As Boolean, ByVal GetG As Boolean, ByVal GetL As Boolean, ByVal GetM As Boolean, ByVal FilterB As String, ByVal FilterG As String, ByVal FilterL As String, ByVal FilterM As String)
        Dim iTables As Integer = Math.Abs(GetB + GetG + GetL + GetM) - 1
        Dim iCurrTbl As Integer = 0
        Dim objArguments As New sGenericTableRequestArguments(AppSettings.Get(CONFIG_APPNAME))

        With objArguments
            ReDim .TableNames(iTables)
            ReDim .StoredProc(iTables)
            ReDim .SPParameters(iTables, 3)

            .dsData = m_dsListDataClass
            .ConnectionDatabase = AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION)
            .ConnectionUser = AppSettings.Get(CONFIG_DB_USER)
            .ConnectionPwd = AppSettings.Get(CONFIG_DB_PWD)

            If GetB Then
                .TableNames(iCurrTbl) = TABLE_LIST_BLOOMBERG
                .StoredProc(iCurrTbl) = TABLE_LIST_BLOOMBERG
                .SPParameters(iCurrTbl, 0) = New SqlClient.SqlParameter("FilterGICS", FilterG)
                .SPParameters(iCurrTbl, 1) = New SqlClient.SqlParameter("FilterLehman", FilterL)
                .SPParameters(iCurrTbl, 2) = New SqlClient.SqlParameter("FilterMSCI", FilterM)
                iCurrTbl += 1
            End If
            If GetG Then
                .TableNames(iCurrTbl) = TABLE_LIST_GICS
                .StoredProc(iCurrTbl) = TABLE_LIST_GICS
                .SPParameters(iCurrTbl, 0) = New SqlClient.SqlParameter("FilterBloomberg", FilterB)
                .SPParameters(iCurrTbl, 1) = New SqlClient.SqlParameter("FilterLehman", FilterL)
                .SPParameters(iCurrTbl, 2) = New SqlClient.SqlParameter("FilterMSCI", FilterM)
                iCurrTbl += 1
            End If
            If GetL Then
                .TableNames(iCurrTbl) = TABLE_LIST_LEHMAN
                .StoredProc(iCurrTbl) = TABLE_LIST_LEHMAN
                .SPParameters(iCurrTbl, 0) = New SqlClient.SqlParameter("FilterGICS", FilterG)
                .SPParameters(iCurrTbl, 1) = New SqlClient.SqlParameter("FilterBloomberg", FilterB)
                .SPParameters(iCurrTbl, 2) = New SqlClient.SqlParameter("FilterMSCI", FilterM)
                iCurrTbl += 1
            End If
            If GetM Then
                .TableNames(iCurrTbl) = TABLE_LIST_MSCI
                .StoredProc(iCurrTbl) = TABLE_LIST_MSCI
                .SPParameters(iCurrTbl, 0) = New SqlClient.SqlParameter("FilterGICS", FilterG)
                .SPParameters(iCurrTbl, 1) = New SqlClient.SqlParameter("FilterBloomberg", FilterB)
                .SPParameters(iCurrTbl, 2) = New SqlClient.SqlParameter("FilterLehman", FilterL)
                .SPParameters(iCurrTbl, 3) = New SqlClient.SqlParameter("Country", "")
                iCurrTbl += 1
            End If
        End With
        GetData(objArguments, m_objSessionItems.SessionKey)
    End Sub

    Private Sub SetupRegulatoryClassLists(ByVal RefreshR As Boolean, ByVal RefreshS As Boolean, ByVal FilterR As String, ByVal FilterS As String)
        Dim strOldVal As String

        GetRegulatoryClassListData(RefreshR, RefreshS, IIf(FilterR = "All", "", FilterR), IIf(FilterS = "All", "", FilterS))
        If RefreshR Then
            strOldVal = ddlRegCls.SelectedValue
            SetupDDL(ddlRegCls, m_dsListDataRegClass.Tables(TABLE_LIST_REG_CLS), False, True)
            If strOldVal <> "" AndAlso Not ddlRegCls.Items.FindByValue(strOldVal) Is Nothing Then
                ddlRegCls.SelectedValue = strOldVal
            End If
        End If
        If RefreshS Then

            strOldVal = ddlRegSubCls.SelectedValue
            If strOldVal <> "" Then
                strOldVal = ddlRegSubCls.SelectedItem.Text
            End If

            Dim dt As New DataTable

            dt = SelectDistinct(m_dsListDataRegClass.Tables(TABLE_LIST_REG_SUB_CLS), "DisplayVal")

            SetupDDL(ddlRegSubCls, dt, False, True)
            If strOldVal <> "" AndAlso Not ddlRegSubCls.Items.FindByText(strOldVal) Is Nothing Then
                ddlRegSubCls.SelectedValue = ddlRegSubCls.Items.FindByText(strOldVal).Value
            End If
        End If
    End Sub


    '------------------------------------------------------------------------------------
    ' Get data for Regulatory and Sub Regulatory Class lists
    '------------------------------------------------------------------------------------
    Private Sub GetRegulatoryClassListData(ByVal GetR As Boolean, ByVal GetS As Boolean, ByVal FilterR As String, ByVal FilterS As String)
        Dim iTables As Integer = Math.Abs(GetR + GetS) - 1
        Dim iCurrTbl As Integer = 0
        Dim objArguments As New sGenericTableRequestArguments(AppSettings.Get(CONFIG_APPNAME))

        With objArguments
            ReDim .TableNames(iTables)
            ReDim .StoredProc(iTables)
            ReDim .SPParameters(iTables, 0)

            .dsData = m_dsListDataRegClass
            .ConnectionDatabase = AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION)
            .ConnectionUser = AppSettings.Get(CONFIG_DB_USER)
            .ConnectionPwd = AppSettings.Get(CONFIG_DB_PWD)

            If GetR Then
                .TableNames(iCurrTbl) = TABLE_LIST_REG_CLS
                .StoredProc(iCurrTbl) = TABLE_LIST_REG_CLS
                .SPParameters(iCurrTbl, 0) = New SqlClient.SqlParameter("FilterRegSubClass", FilterS)
                iCurrTbl += 1
            End If
            If GetS Then
                .TableNames(iCurrTbl) = TABLE_LIST_REG_SUB_CLS
                .StoredProc(iCurrTbl) = TABLE_LIST_REG_SUB_CLS
                .SPParameters(iCurrTbl, 0) = New SqlClient.SqlParameter("FilterRegClass", FilterR)
                iCurrTbl += 1
            End If
        End With
        GetData(objArguments, m_objSessionItems.SessionKey)
    End Sub

    Protected Sub btnRefreshRegClass_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            SetupRegulatoryClassLists(True, True, "", "")
            ddlRegCls.SelectedValue = "All"
            ddlRegSubCls.SelectedValue = "All"

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
    Private Sub OpenPDScorecardMaintenancePage(ByVal id As Integer)
        m_objSessionItems.ItemValues.Set(SESSION_PD_CURR_PAGE, 0)
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_ENTITY_ID, id.ToString("000000"))
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_LEGAL_Name, "")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_ENTITY_STATUS, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_DOMICILE_COUNTRY, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_GICS, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_LEHMAN, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_SCORECARD_TYPE, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_SCORECARD_STATUS, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_PD_RATING, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_START_EFFECTIVE_DATE, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_APPROVER, "")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_APPROVAL_DATE, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_ANALYST, "All")
        m_objSessionItems.ItemValues.Set(SESSION_PD_SEARCH_FILTER_ASSET_CLASS, "All")
        m_objSessionItems.ItemValues.Set(SESSION_SEARCH_FILTER_ROLE_TYPE, "All")

        StoreSession(m_objSessionItems, Session, m_objSessionItems.SessionKey)
        Response.Redirect("PDScorecardMaint.aspx", False)

    End Sub


End Class
