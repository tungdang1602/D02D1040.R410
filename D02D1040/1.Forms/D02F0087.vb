Imports System
Public Class D02F0087

    Private _formIDPermission As String = "D02F0087"
    Public WriteOnly Property FormIDPermission() As String
        Set(ByVal Value As String)
            _formIDPermission = Value
        End Set
    End Property

    Private _bChoose As Boolean = False
    Public Property bChoose() As Boolean
        Get
            Return _bChoose
        End Get
        Set(ByVal Value As Boolean)
            _bChoose = Value
        End Set
    End Property

    Private _sAssetID As String = ""
    Public Property sAssetID() As String
        Get
            Return _sAssetID
        End Get
        Set(ByVal Value As String)
            _sAssetID = Value
        End Set
    End Property

    Private _sMethodID As String = ""
    Public Property sMethodID() As String
        Get
            Return _sMethodID
        End Get
        Set(ByVal Value As String)
            _sMethodID = Value
        End Set
    End Property

    Dim dtSel As DataTable
    Private sKeyString As String = ""
    Private sLastKey As String = ""

    'Private _sSQLInsertD91T1001 As String = ""
    'Public ReadOnly Property sSQLInsertD91T1001 As String
    '    Get
    '        Return _sSQLInsertD91T1001
    '    End Get
    'End Property

    'Private _sSQLUpdateD91T1001 As String = ""
    'Public ReadOnly Property sSQLUpdateD91T1001 As String
    '    Get
    '        Return _sSQLUpdateD91T1001
    '    End Get
    'End Property

    Private _sSQLD91T1001_SaveLastKey As String = "" '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
    Public ReadOnly Property sSQLD91T1001_SaveLastKey As String
        Get
            Return _sSQLD91T1001_SaveLastKey
        End Get
    End Property

    ' NGOC HUY- 107676  -Bổ sung filter lọc theo dạng mới
    Dim oFilterCombo As Lemon3.Controls.FilterCombo
    Private Sub D02F0087_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        LoadInfoGeneral() 'Load System/ Option /... in DxxD9940
        LoadLanguage()
        ' NGOCHUY- 107434  -Bổ sung filter lọc theo dạng mới
        oFilterCombo = New Lemon3.Controls.FilterCombo
        oFilterCombo.CheckD91 = True
        oFilterCombo.UseFilterCombo(tdbcDescription01, tdbcDescription02, tdbcDescription03, tdbcDescription04, tdbcDescription05)
        ''lưu ý cấn đặt trên hàm loadTDBCombo
        LoadTDBCombo()
        InputbyUnicode(Me, gbUnicode)
        SetBackColorObligatory()
        'Nếu form có nút Lọc thì mở ra
        'CheckMenu(Me.Name, TableToolStrip, tdbg.RowCount, gbEnabledUseFind, True/False, ContextMenuStrip1)
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub LoadLanguage()
        '================================================================ 
        Me.Text = rl3("Tao_ma_tu_dongF") & " - " & Me.Name & UnicodeCaption(gbUnicode) 'TÁo mº tø ¢èng
        '================================================================ 
        lblIGEMethodID.Text = rl3("Phuong_phap") 'Phương pháp
        lblGrp1.Text = rl3("Tieu_thuc_tao_ma") 'Tiêu thức tạo mã
        lblInventoryID.Text = rL3("Ma_TSCD") 'Mã TSCĐ
        lblGrp2.Text = rl3("Ma_TSCD_duoc_tao") 'Mã TSCĐ được tạo
        '================================================================ 
        btnCreateID.Text = rl3("Tao_ma") 'Tạo mã
        btnChoose.Text = rl3("Chon") 'Chọn
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        '================================================================ 
        tdbcIGEMethodID.Columns("IGEMethodID").Caption = rl3("Ma") 'Mã
        tdbcIGEMethodID.Columns("IGEMethodName").Caption = rl3("Ten") 'Tên
        tdbcDescription01.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription01.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription06.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription06.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription11.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription11.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription12.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription12.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription07.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription07.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription02.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription02.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription14.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription14.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription09.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription09.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription04.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription04.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription13.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription13.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription08.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription08.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription03.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription03.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription15.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription15.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription10.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription10.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcDescription05.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcDescription05.Columns("SelectionName").Caption = rL3("Ten") 'Tên

        '================================================================ 
        lblDescription01.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription02.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription03.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription04.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription05.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription06.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription07.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription08.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription09.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription10.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription11.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription12.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription13.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription14.Text = rL3("Tieu_thuc") 'Tiêu thức
        lblDescription15.Text = rL3("Tieu_thuc") 'Tiêu thức

    End Sub



    Private Sub D02F0087_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt Then
        ElseIf e.Control Then
        Else
            Select Case e.KeyCode
                Case Keys.Enter
                    UseEnterAsTab(Me, True)
            End Select
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        sSQL = "--Combo Phương pháp" & vbCrLf & _
         "SELECT      IGEMethodID,IGEMethodNameU as IGEMethodName, Defaults" & vbCrLf & _
         "FROM        D91T0045" & vbCrLf & _
         "WHERE    ModuleID  = '02' AND FormID  = 'D02F0070'" & vbCrLf & _
         "AND Disabled  = 0" & vbCrLf & _
         "ORDER BY    IGEMethodID"
        Dim dtIGE As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(tdbcIGEMethodID, dtIGE, gbUnicode)
        If dtIGE.Rows.Count > 0 Then
            Dim dr() As DataRow = dtIGE.Select("Defaults=1")
            If dr.Length > 0 Then
                tdbcIGEMethodID.SelectedValue = dr(0)("IGEMethodID")
            Else
                tdbcIGEMethodID.SelectedIndex = 0
            End If
        End If


    End Sub

#Region "Events tdbcIGEMethodID with txtIGEMethodName"

    Private Sub tdbcIGEMethodID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcIGEMethodID.SelectedValueChanged
        If tdbcIGEMethodID.SelectedValue Is Nothing Then
            txtIGEMethodName.Text = ""
        Else
            txtIGEMethodName.Text = tdbcIGEMethodID.Columns(1).Value.ToString
        End If
        pnl1.ResetText()
        txtInventoryID.Text = ""
        LoadSelectionID()
    End Sub

    Private Sub tdbcIGEMethodID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcIGEMethodID.LostFocus
        If tdbcIGEMethodID.FindStringExact(tdbcIGEMethodID.Text) = -1 Then
            tdbcIGEMethodID.Text = ""
        End If
    End Sub

#End Region

    Private Sub AddNewAssetS(tdbc As C1.Win.C1List.C1Combo, iIndexTab As Integer)
        If ReturnPermission("D02F3000") <= 1 Then
            D99C0008.MsgL3(rL3("Ban_khong_co_quyen_them_moi"), L3MessageBoxIcon.Information)
            tdbc.Text = ""
            Exit Sub
        End If
        Dim sValue As String = ""

        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "IndexTab", iIndexTab)
        SetProperties(arrPro, "FormIDPermission", "D02F3000")
        Dim frm As Form = CallFormShowDialog("D02D1240", "D02F3001", arrPro)
        If frm Is Nothing Then Exit Sub 'TH form đã gọi rồi thì không gọi nữa
        If L3Bool(GetProperties(frm, "SavedOk")) Then
            sValue = GetProperties(frm, "AssetID").ToString
        End If

        If sValue = "" Then
            tdbc.Text = ""
            'tdbc.SelectedValue = ""
        Else
            LoadSelCombo(tdbc, tdbc.Tag.ToString)
            tdbc.Text = sValue
        End If
    End Sub

#Region "Events tdbcDescription01"
    Private Sub tdbcDescription01_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription01.Validated
        oFilterCombo.FilterCombo(tdbcDescription01, e)
        If tdbcDescription01.FindStringExact(tdbcDescription01.Text) = -1 Or tdbcDescription01.Text = "" Then tdbcDescription01.Text = ""

        If tdbcDescription01.Text = "+" Then '16/7/2020, Đặng Ngọc Tài:id 141815-SVI_Cho phép thêm mới mã phân loại tại màn hình tạo mã TSCD
            AddNewAssetS(tdbcDescription01, 0)
        End If
    End Sub

#End Region

#Region "Events tdbcDescription02"

    Private Sub tdbcDescription02_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription02.Validated
        oFilterCombo.FilterCombo(tdbcDescription02, e)
        If tdbcDescription02.FindStringExact(tdbcDescription02.Text) = -1 Or tdbcDescription02.Text = "" Then tdbcDescription02.Text = ""

        If tdbcDescription02.Text = "+" Then '16/7/2020, Đặng Ngọc Tài:id 141815-SVI_Cho phép thêm mới mã phân loại tại màn hình tạo mã TSCD
            AddNewAssetS(tdbcDescription02, 1)
        End If
    End Sub

#End Region

#Region "Events tdbcDescription03"

    Private Sub tdbcDescription03_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription03.Validated
        oFilterCombo.FilterCombo(tdbcDescription03, e)
        If tdbcDescription03.FindStringExact(tdbcDescription03.Text) = -1 Or tdbcDescription03.Text = "" Then tdbcDescription03.Text = ""

        If tdbcDescription03.Text = "+" Then '16/7/2020, Đặng Ngọc Tài:id 141815-SVI_Cho phép thêm mới mã phân loại tại màn hình tạo mã TSCD
            AddNewAssetS(tdbcDescription03, 2)
        End If
    End Sub

#End Region

#Region "Events tdbcDescription04"

    Private Sub tdbcDescription04_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription04.Validated
        oFilterCombo.FilterCombo(tdbcDescription04, e)
        If tdbcDescription04.FindStringExact(tdbcDescription04.Text) = -1 Or tdbcDescription04.Text = "" Then tdbcDescription04.Text = ""

        If tdbcDescription04.Text = "+" Then '16/7/2020, Đặng Ngọc Tài:id 141815-SVI_Cho phép thêm mới mã phân loại tại màn hình tạo mã TSCD
            AddNewAssetS(tdbcDescription04, 3)
        End If
    End Sub

#End Region

#Region "Events tdbcDescription05"
    Private Sub tdbcDescription05_Validated(sender As Object, e As EventArgs) Handles tdbcDescription05.Validated
        oFilterCombo.FilterCombo(tdbcDescription05, e)
        If tdbcDescription05.FindStringExact(tdbcDescription05.Text) = -1 Or tdbcDescription05.Text = "" Then tdbcDescription05.Text = ""

        If tdbcDescription05.Text = "+" Then '16/7/2020, Đặng Ngọc Tài:id 141815-SVI_Cho phép thêm mới mã phân loại tại màn hình tạo mã TSCD
            AddNewAssetS(tdbcDescription05, 4)
        End If
    End Sub

#End Region

#Region "Events tdbcDescription06"

    Private Sub tdbcDescription06_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription06.LostFocus
        If tdbcDescription06.FindStringExact(tdbcDescription06.Text) = -1 Then tdbcDescription06.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription07"

    Private Sub tdbcDescription07_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription07.LostFocus
        If tdbcDescription07.FindStringExact(tdbcDescription07.Text) = -1 Then tdbcDescription07.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription08"

    Private Sub tdbcDescription08_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription08.LostFocus
        If tdbcDescription08.FindStringExact(tdbcDescription08.Text) = -1 Then tdbcDescription08.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription09"

    Private Sub tdbcDescription09_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription09.LostFocus
        If tdbcDescription09.FindStringExact(tdbcDescription09.Text) = -1 Then tdbcDescription09.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription10"

    Private Sub tdbcDescription10_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription10.LostFocus
        If tdbcDescription10.FindStringExact(tdbcDescription10.Text) = -1 Then tdbcDescription10.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription11"

    Private Sub tdbcDescription11_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription11.LostFocus
        If tdbcDescription11.FindStringExact(tdbcDescription11.Text) = -1 Then tdbcDescription11.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription12"

    Private Sub tdbcDescription12_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription12.LostFocus
        If tdbcDescription12.FindStringExact(tdbcDescription12.Text) = -1 Then tdbcDescription12.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription13"

    Private Sub tdbcDescription13_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription13.LostFocus
        If tdbcDescription13.FindStringExact(tdbcDescription13.Text) = -1 Then tdbcDescription13.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription14"

    Private Sub tdbcDescription14_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription14.LostFocus
        If tdbcDescription14.FindStringExact(tdbcDescription14.Text) = -1 Then tdbcDescription14.Text = ""
    End Sub

#End Region

#Region "Events tdbcDescription15"

    Private Sub tdbcDescription15_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescription15.LostFocus
        If tdbcDescription15.FindStringExact(tdbcDescription15.Text) = -1 Then tdbcDescription15.Text = ""
    End Sub

#End Region
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0079
    '# Created User: KIM LONG
    '# Created Date: 29/11/2016 11:49:13
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0079() As String
        Dim sSQL As String = ""
        sSQL &= ("-- MethodID" & vbCrlf)
        sSQL &= "Exec D02P0079 "
        sSQL &= SQLString(ReturnValueC1Combo(tdbcIGEMethodID)) & COMMA 'IGEMethodID, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString("D02F0070") 'FormID, varchar[20], NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0078
    '# Created User: KIM LONG
    '# Created Date: 29/11/2016 12:51:55
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0078(ByVal sCode As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Combo Tieu thuc " & vbCrLf)
        sSQL &= "Exec D02P0078 "
        sSQL &= SQLString("D02F0070") & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(sCode) & COMMA 'Code, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD91T1000
    '# Created User: KIM LONG
    '# Created Date: 29/11/2016 01:08:30
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD91T1000() As String
        Dim sSQL As String = ""
        sSQL &= ("-- DELETE 	D91T1000 	 " & vbCrLf)
        sSQL &= "Delete From D91T1000"
        sSQL &= " Where "
        sSQL &= "UserID = " & SQLString(gsUserID) & " And "
        sSQL &= "HostID = " & SQLString(My.Computer.Name) & " And "
        sSQL &= "FormID = " & SQLString("D02F0070") & " And "
        sSQL &= "ModuleID  = " & SQLString("02")
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T1000
    '# Created User: KIM LONG
    '# Created Date: 29/11/2016 01:10:20
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T1000(ByVal sCode As String, ByVal sID As String, ByVal sName As String) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- --Insert du lieu tung combo tieu thuc tren luoi " & vbCrLf)
        sSQL.Append("Insert Into D91T1000(")
        sSQL.Append("UserID, HostID, ModuleID, FormID, Code, " & vbCrLf)
        sSQL.Append("SelectionID, SelectionNameU")
        sSQL.Append(") Values(" & vbCrLf)
        sSQL.Append(SQLString(gsUserID) & COMMA) 'UserID, varchar[50], NOT NULL
        sSQL.Append(SQLString(My.Computer.Name) & COMMA) 'HostID, varchar[50], NOT NULL
        sSQL.Append(SQLString("02") & COMMA) 'ModuleID, varchar[50], NOT NULL
        sSQL.Append(SQLString("D02F0070") & COMMA) 'FormID, varchar[50], NOT NULL
        sSQL.Append(SQLString(sCode) & COMMA & vbCrLf) 'Code, varchar[500], NOT NULL
        sSQL.Append(SQLString(sID) & COMMA) 'SelectionID, varchar[50], NOT NULL
        sSQL.Append(SQLStringUnicode(sName, gbUnicode, True)) 'SelectionNameU, nvarchar[1000], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P1000
    '# Created User: KIM LONG
    '# Created Date: 29/11/2016 01:13:54
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD91P1000() As String
        Dim sSQL As String = ""
        sSQL &= ("-- D91P1000   " & vbCrlf)
        sSQL &= "Exec D91P1000 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString("D02F0070") & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString("02") & COMMA 'ModuleID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcIGEMethodID)) & COMMA 'IGEMethodID, varchar[50], NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'AutoCreateName, tinyint, NOT NULL
        sSQL &= SQLNumber(50) & COMMA 'Length, tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'IsD07F0011, tinyint, NOT NULL
        sSQL &= SQLNumber(0) 'NewLastKey, int, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T1001
    '# Created User: KIM LONG
    '# Created Date: 29/11/2016 02:15:57
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T1001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Insert" & vbCrLf)
        sSQL.Append("Insert Into D91T1001(")
        sSQL.Append("KeyString, LastKey, ModuleID, FormID")
        sSQL.Append(") Values(" & vbCrlf)
        sSQL.Append(SQLString(sKeyString) & COMMA) 'KeyString, varchar[250], NOT NULL
        sSQL.Append(SQLNumber(sLastKey) & COMMA) 'LastKey, int, NOT NULL
        sSQL.Append(SQLString("02") & COMMA) 'ModuleID, varchar[20], NOT NULL
        sSQL.Append(SQLString("D02F0070")) 'FormID, varchar[20], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD91T1001
    '# Created User: KIM LONG
    '# Created Date: 29/11/2016 02:17:00
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD91T1001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Update" & vbCrlf)
        sSQL.Append("Update D91T1001 Set ")
        sSQL.Append(" LastKey = " & SQLNumber(sLastKey) & vbCrLf) 'int, NOT NULL
        sSQL.Append(" Where KeyString = " & SQLString(sKeyString))
        sSQL.Append(" and ModuleID = " & SQLString("02"))
        sSQL.Append(" and FormID  = " & SQLString("D02F0070"))
        Return sSQL
    End Function

    Private Sub LoadSelectionID()
        ReadOnlyAll(pnl1)
        Dim sSQL As String = SQLStoreD02P0079()
        dtSel = ReturnDataTable(sSQL)
        If dtSel.Rows.Count > 0 Then
            For i As Integer = 0 To dtSel.Rows.Count - 1
                ' Điều kiện TranMonth.....
                If L3String(dtSel.Rows(i)("Code")) <> "{TranMonth(MM)}" And L3String(dtSel.Rows(i)("Code")) <> "{TranYear(YYYY)}" And L3String(dtSel.Rows(i)("Code")) <> "{TranYear(YY)}" And L3String(dtSel.Rows(i)("Code")) <> "{Month(MM)}" And L3String(dtSel.Rows(i)("Code")) <> "{Year(YYYY)}" And L3String(dtSel.Rows(i)("Code")) <> "{Year(YY)}" Then
                    Dim lbl() As Control = Me.Controls.Find("lblDescription" & (i + 1).ToString("00"), True)
                    Dim tdbc() As Control = Me.Controls.Find("tdbcDescription" & (i + 1).ToString("00"), True)
                    If lbl.Length > 0 Then lbl(0).Text = L3String(dtSel.Rows(i)("Description"))
                    If tdbc.Length > 0 Then
                        CType(tdbc(0), C1.Win.C1List.C1Combo).Tag = L3String(dtSel.Rows(i)("Code"))
                        LoadSelCombo(CType(tdbc(0), C1.Win.C1List.C1Combo), L3String(dtSel.Rows(i)("Code")))
                    End If
                    'CType(tdbc(0), C1.Win.C1List.C1Combo).ReadOnly = False
                    UnReadOnlyControl(True, CType(tdbc(0), C1.Win.C1List.C1Combo))
                End If
                'Dim lbl() As Control = Me.Controls.Find("lblDescription" & (i + 1).ToString("00"), True)
                'Dim tdbc() As Control = Me.Controls.Find("tdbcDescription" & (i + 1).ToString("00"), True)
                'If lbl.Length > 0 Then lbl(0).Text = L3String(dtSel.Rows(i)("Description"))
                'If tdbc.Length > 0 Then
                '    CType(tdbc(0), C1.Win.C1List.C1Combo).Tag = L3String(dtSel.Rows(i)("Code"))
                '    'If L3String(dtSel.Rows(i)("Code")) = "{PL1}" Then

                '    LoadSelCombo(CType(tdbc(0), C1.Win.C1List.C1Combo), L3String(dtSel.Rows(i)("Code")))
                '    'End If
                'End If
                ''CType(tdbc(0), C1.Win.C1List.C1Combo).ReadOnly = False
                'UnReadOnlyControl(True, CType(tdbc(0), C1.Win.C1List.C1Combo))

            Next
        End If
    End Sub

    Private Sub LoadSelCombo(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sCode As String)
        LoadDataSource(tdbc, SQLStoreD02P0078(sCode), gbUnicode)
    End Sub

    Private Function AllowCreate() As Boolean
        If tdbcIGEMethodID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(lblIGEMethodID.Text)
            tdbcIGEMethodID.Focus()
            Return False
        End If

        If Not tdbcDescription01.ReadOnly Then
            If tdbcDescription01.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription01.Text)
                tdbcDescription01.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription02.ReadOnly Then
            If tdbcDescription02.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription02.Text)
                tdbcDescription02.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription03.ReadOnly Then
            If tdbcDescription03.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription03.Text)
                tdbcDescription03.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription04.ReadOnly Then
            If tdbcDescription04.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription04.Text)
                tdbcDescription04.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription05.ReadOnly Then
            If tdbcDescription05.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription05.Text)
                tdbcDescription05.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription06.ReadOnly Then
            If tdbcDescription06.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription06.Text)
                tdbcDescription06.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription07.ReadOnly Then
            If tdbcDescription07.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription07.Text)
                tdbcDescription07.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription08.ReadOnly Then
            If tdbcDescription08.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription08.Text)
                tdbcDescription08.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription09.ReadOnly Then
            If tdbcDescription09.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription09.Text)
                tdbcDescription09.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription10.ReadOnly Then
            If tdbcDescription10.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription10.Text)
                tdbcDescription10.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription11.ReadOnly Then
            If tdbcDescription11.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription11.Text)
                tdbcDescription11.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription12.ReadOnly Then
            If tdbcDescription12.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription12.Text)
                tdbcDescription12.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription13.ReadOnly Then
            If tdbcDescription13.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription13.Text)
                tdbcDescription13.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription14.ReadOnly Then
            If tdbcDescription14.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription14.Text)
                tdbcDescription14.Focus()
                Return False
            End If
        End If

        If Not tdbcDescription15.ReadOnly Then
            If tdbcDescription15.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescription15.Text)
                tdbcDescription15.Focus()
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub SetBackColorObligatory()
        tdbcDescription01.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcIGEMethodID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription02.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription03.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription04.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription05.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription06.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription07.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription08.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription09.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription10.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription11.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription12.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription13.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription14.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescription15.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub




    Private Sub btnCreateID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateID.Click
        If Not AllowCreate() Then Exit Sub
        Dim sSQL As New StringBuilder
        Dim dtCreateID As New DataTable
        sSQL.Append(SQLDeleteD91T1000() & vbCrLf)
        For i As Integer = 0 To dtSel.Rows.Count - 1
            Dim ctrl() As Control = Me.Controls.Find("tdbcDescription" & (i + 1).ToString("00"), True)
            Dim tdbc As C1.Win.C1List.C1Combo = CType(ctrl(0), C1.Win.C1List.C1Combo)

            sSQL.Append(SQLInsertD91T1000(L3String(dtSel.Rows(i)("Code")) _
                        , ReturnValueC1Combo(tdbc) _
                        , ReturnValueC1Combo(tdbc, "SelectionName")).ToString & vbCrLf)
        Next
        sSQL.Append(SQLStoreD91P1000() & vbCrLf)
        If Not CheckStorebyTrans(sSQL.ToString, dtCreateID, False) Then Exit Sub
        If dtCreateID.Rows.Count > 0 Then
            sKeyString = L3String(dtCreateID.Rows(0)("KeyString"))
            sLastKey = L3String(dtCreateID.Rows(0)("LastKey"))
            txtInventoryID.Text = L3String(dtCreateID.Rows(0)("ID"))
        End If

    End Sub

    Private Function AllowChoose() As Boolean
        If txtInventoryID.Text.Trim = "" Then
            D99C0008.Msg(rL3("Ma_tai_san_chua_duoc_tao"))
            txtInventoryID.Focus()
            Return False
        End If
        If txtInventoryID.Text.Trim.Length > 20 Then
            D99C0008.Msg(rL3("Ma_tai_san_khong_duoc_vuot_qua_20_ky_tu"))
            txtInventoryID.Focus()
            Return False
        End If
        Return True
    End Function


    Private Sub btnChoose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChoose.Click
        If Not AllowChoose() Then Exit Sub
        'Dim sSQL As New StringBuilder
        '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
        If sLastKey = "1" Then
            _sSQLD91T1001_SaveLastKey = SQLInsertD91T1001.ToString '   sSQL.Append(SQLInsertD91T1001())
        Else
            _sSQLD91T1001_SaveLastKey = SQLUpdateD91T1001.ToString '  sSQL.Append(SQLUpdateD91T1001())
        End If
        'Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        'If bRunSQL Then
        _bChoose = True
        sAssetID = txtInventoryID.Text
        sMethodID = ReturnValueC1Combo(tdbcIGEMethodID)
        Me.Close()
        'End If
    End Sub
    
  
End Class