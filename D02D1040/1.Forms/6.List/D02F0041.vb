Imports System
'#-------------------------------------------------------------------------------------
'# Created Date: 22/10/2007 11:52:49 AM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 22/10/2007 11:52:49 AM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Public Class D02F0041
	Dim report As D99C2003

    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

	Private _formIDPermission As String = "D02F0041"
	Public WriteOnly Property FormIDPermission() As String
		Set(ByVal Value As String)
			       _formIDPermission = Value
		   End Set
	End Property


#Region "Const of tdbg"
    Private Const COL_ACodeID As Integer = 0          ' Mã khoản mục
    Private Const COL_Description As Integer = 1      ' Diễn giải
    Private Const COL_Disabled As Integer = 2         ' Không sử dụng 
    Private Const COL_TypeCodeID As Integer = 3       ' TypeCodeID
    Private Const COL_Type As Integer = 4             ' Type
    Private Const COL_CreateUserID As Integer = 5     ' CreateUserID
    Private Const COL_CreateDate As Integer = 6       ' CreateDate
    Private Const COL_LastModifyUserID As Integer = 7 ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 8   ' LastModifyDate
#End Region

    Dim dtGrid, dtCaptionCols As DataTable
    Private sKey As String
    Private sTypeCodeID As String
    Dim bRefreshFilter As Boolean
    Dim sFilter As New System.Text.StringBuilder()

    Private Sub D02F0040_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
        End If
    End Sub

    Private Sub D02F0040_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1)
        Loadlanguage()
        ResetColorGrid(tdbg)
        gbEnabledUseFind = False
        LoadTDBCombo()
        tdbcTypeCodeID.Text = "%"
        InputbyUnicode(Me, gbUnicode)
        SetResolutionForm(Me, ContextMenuStrip1)
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_muc_ma_phan_tich_-_D02F0041") & UnicodeCaption(gbUnicode) 'Danh móc mº ph¡n tÛch - D02F0041
        '================================================================ 
        lblTypeCodeID.Text = rl3("Loai_phan_tich") 'Loại phân tích
        '================================================================ 
        tdbcTypeCodeID.Columns("TypeCodeID").Caption = rl3("Ma") 'Mã
        tdbcTypeCodeID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        tdbg.Columns("ACodeID").Caption = rl3("Ma_khoan_muc") 'Mã khoản mục
        tdbg.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("Disabled").Caption = rl3("KSD") 'KSD
        '================================================================ 
        chkShowDisabled.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
    End Sub

    'Private Sub LoadTDBGrid(ByVal sTypeCodeID As String, Optional ByVal bFlag As Boolean = False)
    '    Dim sSQL As String = ""
    '    sSQL = "Select * From D02T0041 Where Type = 'A' And TypeCodeID like " & SQLString(sTypeCodeID) & " Order by ACodeID "
    '    dtGrid = ReturnDataTable(sSQL)
    '    LoadDataSource(tdbg, dtGrid)

    '    If .bSaved Then
    '        dtGrid.DefaultView.Sort = "ACodeID"
    '        tdbg.Bookmark = dtGrid.DefaultView.Find(sKey)
    '    End If
    'End Sub

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String = ""
        sSQL = "Select      ACodeID, Description" & UnicodeJoin(gbUnicode) & " As Description, " & vbCrLf
        sSQL &= "           Disabled, TypeCodeID, Type," & vbCrLf
        sSQL &= "           CreateUserID, CreateDate, LastModifyUserID, LastModifyDate" & vbCrLf
        sSQL &= "From       D02T0041 WITH(NOLOCK)" & vbCrLf
        sSQL &= "Where      Type = 'A' And TypeCodeID like " & SQLString(tdbcTypeCodeID.Text) & vbCrLf
        sSQL &= "Order by   ACodeID " & vbCrLf

        dtGrid = ReturnDataTable(sSQL)

        gbEnabledUseFind = dtGrid.Rows.Count > 0

        If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFilter = New System.Text.StringBuilder("")
            sFind = ""
        End If

        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("ACodeID = " & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0))
        End If

        If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
    End Sub

    Private Sub ReLoadTDBGrid(Optional ByVal bUseFilterBar As Boolean = False)
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString

        If Not chkShowDisabled.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "Disabled =0"
        End If
        dtGrid.DefaultView.RowFilter = strFind

        CheckMenu(_formIDPermission, TableToolStrip, tdbg.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
        FooterTotalGrid(tdbg, COL_ACodeID)
    End Sub

    Private Sub LoadTDBCombo()
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, gbUnicode)
        Dim sSQL As String = ""
        'Load tdbcTypeCodeID
        sSQL = "Select 0 as DisplayOrder,'%' As TypeCodeID, " & sLanguage & " Description  " & vbCrLf
        sSQL &= "Union All" & vbCrLf
        sSQL &= "Select 1 as DisplayOrder,TypeCodeID, " & IIf(geLanguage = EnumLanguage.Vietnamese, "VieTypeCodeName" & sUnicode, "EngTypeCodeName" & sUnicode).ToString & " As Description " & vbCrLf
        sSQL &= " From D02T0040 WITH(NOLOCK) Where Type = 'A' And Disabled = 0 Order By DisplayOrder,TypeCodeID "
        LoadDataSource(tdbcTypeCodeID, sSQL, gbUnicode)
    End Sub

#Region "Events tdbcTypeCodeID with txtTypeCodeName"

    Private Sub tdbcTypeCodeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.Close
        If tdbcTypeCodeID.FindStringExact(tdbcTypeCodeID.Text) = -1 Then
            tdbcTypeCodeID.Text = ""
            txtTypeCodeName.Text = ""
        End If
    End Sub

    Private Sub tdbcTypeCodeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.SelectedValueChanged
        txtTypeCodeName.Text = tdbcTypeCodeID.Columns(1).Value.ToString
        LoadTDBGrid(True)
    End Sub

    Private Sub tdbcTypeCodeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcTypeCodeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcTypeCodeID.Text = ""
            txtTypeCodeName.Text = ""
        End If
    End Sub

#End Region

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
	Dim gbEnabledUseFind As Boolean = False
    'Cần sửa Tìm kiếm như sau:
	'Bỏ sự kiện Finder_FindClick.
	'Sửa tham số Me.Name -> Me
	'Phải tạo biến properties có tên chính xác strNewFind và strNewServer
	'Sửa gdtCaptionExcel thành dtCaptionCols: biến toàn cục trong form
	'Nếu có F12 dùng D09U1111 thì Sửa dtCaptionCols thành ResetTableByGrid(usrOption, dtCaptionCols.DefaultView.ToTable)
    Private sFind As String = ""
	Public WriteOnly Property strNewFind() As String
		Set(ByVal Value As String)
			sFind = Value
			ReLoadTDBGrid()'Làm giống sự kiện Finder_FindClick. Ví dụ đối với form Báo cáo thường gọi btnPrint_Click(Nothing, Nothing): sFind = "
		End Set
	End Property

    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me, "0", gbUnicode)
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Menu bar"

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim f As New D02F0042
        With f
            .TypeCodeID = tdbcTypeCodeID.Text
            .ACodeID = ""
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            sKey = .ACodeID
            .Dispose()
        End With
        If f.SavedOK Then
            LoadTDBGrid(True, sKey)
        End If
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim f As New D02F0042
        With f
            .TypeCodeID = tdbg.Columns(COL_TypeCodeID).Text
            .ACodeID = tdbg.Columns(COL_ACodeID).Text
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        If tdbg.RowCount <= 0 Then Exit Sub
        Dim f As New D02F0042
        With f
            .TypeCodeID = tdbg.Columns(COL_TypeCodeID).Text
            .ACodeID = tdbg.Columns(COL_ACodeID).Text
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            sKey = .ACodeID
            .Dispose()
            If .SavedOK = True Then
                LoadTDBGrid(False, sKey)
            End If
        End With
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        Dim iBookmark As Integer
        Dim bResult As Boolean
        If D99C0008.MsgAskDelete = Windows.Forms.DialogResult.No Then Exit Sub

        Dim sSQL As String = ""
        If IsExistKey("D02T0001", "ACode" & Microsoft.VisualBasic.Right(tdbg.Columns(COL_TypeCodeID).Text, 2) & "ID", tdbg.Columns(COL_ACodeID).Text) Then
            D99C0008.MsgCanNotDelete()
            Exit Sub
        Else
            If Not IsDBNull(tdbg.Bookmark) Then iBookmark = tdbg.Bookmark
            sSQL = "Delete D02T0041 Where ACodeID=" & SQLString(tdbg.Columns(COL_ACodeID).Text) & " And TypeCodeID=" & SQLString(tdbg.Columns(COL_TypeCodeID).Text)
            bResult = ExecuteSQL(sSQL)

            If bResult Then
                DeleteOK()
                LoadTDBGrid()
                If Not IsDBNull(iBookmark) Then tdbg.Bookmark = iBookmark
            Else
                DeleteNotOK()
            End If
        End If
    End Sub

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(Me,tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbPrint.Click, tsmPrint.Click, mnsPrint.Click
        'Dim report As New D99C1003
        'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
		If Not AllowNewD99C2003(report, Me) Then Exit Sub
		'************************************
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = "D02R0041"
        Dim sSubReportName As String = "D02R0000"
        Dim sReportCaption As String = ""
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        sReportCaption = rl3("Danh_muc_ma_phan_tich") & " - " & sReportName
        sPathReport = UnicodeGetReportPath(gbUnicode, D02Options.ReportLanguage, "") & sReportName & ".rpt"

        sSQL = "Select * From D02T0041 WITH(NOLOCK) Where Type = 'A' And TypeCodeID like " & SQLString(tdbcTypeCodeID.Text) & " Order By ACodeID " & vbCrLf
        sSQLSub = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
        UnicodeSubReport(sSubReportName, sSQLSub, , gbUnicode)

        With report
            .OpenConnection(conn)
            If tdbcTypeCodeID.Text = "%" Then
                .AddParameter("txtTitleReport", rl3("DANH_MUC_MA_PHAN_TICHV"))
            Else
                .AddParameter("txtTitleReport", rl3("DANH_MUC_MA_PHAN_TICHV") & " " & tdbcTypeCodeID.Text)
            End If
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(dtGrid.DefaultView.ToTable)
            .PrintReport(sPathReport, sReportCaption)
        End With
    End Sub

    Private Sub chkShowDisabled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub
#End Region

#Region "Grid"

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If tdbg.FilterActive Then Exit Sub

        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg, e)
    End Sub

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            FilterChangeGrid(tdbg, sFilter)
            ReLoadTDBGrid()
        Catch ex As Exception
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

#End Region

End Class