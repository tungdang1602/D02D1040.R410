'#-------------------------------------------------------------------------------------
'# Created Date: 24/10/2007 3:25:42 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 24/10/2007 3:25:42 PM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------

Imports System.Text
Imports System

Public Class D02F0042


    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

    Private _formName As String = ""
    Public WriteOnly Property FormName As String
        Set(ByVal Value As String)
            _formName = Value
        End Set
    End Property
    
    Private _aCodeID As String = ""
    Public Property ACodeID() As String
        Get
            Return _aCodeID
        End Get
        Set(ByVal value As String)
            _aCodeID = value
        End Set
    End Property

    Private _typeCodeID As String = ""
    Public Property TypeCodeID() As String
        Get
            Return _typeCodeID
        End Get
        Set(ByVal value As String)
            _typeCodeID = value
        End Set
    End Property

    Dim bLoadFormState As Boolean = False
	Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
	bLoadFormState = True
	LoadInfoGeneral()
            _FormState = value
            LoadTDBCombo()
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnSave.Enabled = True
                    btnNext.Enabled = False
                    LoadAddNew()
                Case EnumFormState.FormEdit
                    btnSave.Enabled = True
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
                Case EnumFormState.FormView
                    btnSave.Enabled = False
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
            End Select
        End Set
    End Property

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F0042_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F0042_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
	If bLoadFormState = False Then FormState = _formState
        Loadlanguage()
        SetBackColorObligatory()
        InputbyUnicode(Me, gbUnicode)
        CheckIdTextBox(txtACodeID)

        If _formName = "D02F1031" Then '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
            ReadOnlyControl(tdbcTypeCodeID)
        End If

        SetResolutionForm(Me)
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_ma_phan_tich_-_D02F0042") & UnicodeCaption(gbUnicode) 'CËp nhËt mº ph¡n tÛch - D02F0042
        '================================================================ 
        lblTypeCodeID.Text = rl3("Ma_loai_phan_tich") 'Mã loại phân tích
        lblACodeID.Text = rl3("Ma_khoan_muc") 'Mã khoản mục
        lblDescription.Text = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnNext.Text = rl3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        chkDisabled.Text = rl3("Khong_su_dung") 'Không sử dụng
        '================================================================ 
        tdbcTypeCodeID.Columns("TypeCodeID").Caption = rl3("Ma") 'Mã
        tdbcTypeCodeID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcTypeCodeID
        sSQL = " Select TypeCodeID," & IIf(geLanguage = EnumLanguage.Vietnamese, "VieTypeCodeName" & UnicodeJoin(gbUnicode), "EngTypeCodeName" & UnicodeJoin(gbUnicode)).ToString & "  As Description, MaxLength , CopyToD19" & vbCrLf
        sSQL &= " From D02T0040 WITH(NOLOCK) Where Type ='A' Order By TypeCodeID "
        LoadDataSource(tdbcTypeCodeID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadMaster()
        Dim sSQL As String = ""
        sSQL = "Select * From D02T0041 WITH(NOLOCK) Where ACodeID = " & SQLString(_aCodeID) & " And TypeCodeID = " & SQLString(_typeCodeID) & " And Type = 'A'"
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                tdbcTypeCodeID.SelectedValue = .Item("TypeCodeID").ToString
                txtACodeID.Text = .Item("ACodeID").ToString
                chkDisabled.Checked = CBool(.Item("Disabled"))
                txtDescription.Text = .Item("Description" & UnicodeJoin(gbUnicode)).ToString
            End With
        End If
    End Sub

    Private Sub LoadAddNew()
        LoadMaster()
        If _typeCodeID <> "%" Then
            tdbcTypeCodeID.SelectedValue = _typeCodeID
        End If
    End Sub

    Private Sub LoadEdit()
        tdbcTypeCodeID.Enabled = False
        txtACodeID.Enabled = False
        LoadMaster()
        txtDescription.Focus()
    End Sub

    Private Sub SetBackColorObligatory()
        tdbcTypeCodeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtACodeID.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

#Region "Events tdbcTypeCodeID with txtACodeID"

    Private Sub tdbcTypeCodeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.Close
        If tdbcTypeCodeID.FindStringExact(tdbcTypeCodeID.Text) = -1 Then
            tdbcTypeCodeID.Text = ""
            txtTypeCodeName.Text = ""
        End If
    End Sub

    Private Sub tdbcTypeCodeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.SelectedValueChanged
        txtTypeCodeName.Text = tdbcTypeCodeID.Columns("Description").Value.ToString
        If tdbcTypeCodeID.Columns("MaxLength").Value.ToString <> "" Then
            txtACodeID.MaxLength = CInt(tdbcTypeCodeID.Columns("MaxLength").Value)
        End If
    End Sub

    Private Sub tdbcTypeCodeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcTypeCodeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcTypeCodeID.Text = ""
            txtTypeCodeName.Text = ""
        End If
    End Sub

#End Region

    Private Function AllowSave() As Boolean
        Dim sArrField As String() = {"ACodeID", "TypeCodeID"}
        Dim sArrValue As String() = {txtACodeID.Text, tdbcTypeCodeID.Text}
        If tdbcTypeCodeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Ma_loai_phan_tich"))
            tdbcTypeCodeID.Focus()
            Return False
        End If
        If txtACodeID.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ma_khoan_muc"))
            txtACodeID.Focus()
            Return False
        End If
        If _FormState = EnumFormState.FormAdd Then
            If IsExistKey("D02T0041", sArrField, sArrValue) Then
                D99C0008.MsgDuplicatePKey()
                tdbcTypeCodeID.Focus()
                Return False
            End If
        End If

        
        Return True
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        'Kiểm tra trước khi lưu
        If Not AllowSave() Then Exit Sub
        btnSave.Enabled = False
        btnClose.Enabled = False
        _savedOK = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                sSQL.Append(SQLInsertD02T0041().ToString & vbCrLf)
                Dim sSQL1 As String = "SELECT TOP 1 1 FROM D19T1801 WITH(NOLOCK) WHERE TypeCodeID = " & SQLString("C" & ReturnValueC1Combo(tdbcTypeCodeID).Substring(1, 2)) & " And CCodeID = " & SQLString(txtACodeID.Text)
                If L3Bool(ReturnScalar(sSQL1)) Then
                    D99C0008.Msg(rL3("Ma_nay_da_ton_tai_o_Module_CPTT"))
                Else
                    If L3Bool(ReturnValueC1Combo(tdbcTypeCodeID, "CopyToD19")) Then
                        Dim sTypeCodeID As String = "C" & ReturnValueC1Combo(tdbcTypeCodeID).Substring(1, 2)
                        sSQL.Append(SQLInsertD19T1801(sTypeCodeID))
                    End If
                End If
            Case EnumFormState.FormEdit
                sSQL.Append(SQLUpdateD02T0041.ToString & vbCrLf)
                If L3Bool(ReturnValueC1Combo(tdbcTypeCodeID, "CopyToD19")) Then
                    Dim sTypeCodeID As String = "C" & ReturnValueC1Combo(tdbcTypeCodeID).Substring(1, 2)
                    sSQL.Append(SQLUpdateD19T1801(sTypeCodeID))
                End If
        End Select
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _savedOK = True
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _aCodeID = txtACodeID.Text
                    btnNext.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit
                    btnSave.Enabled = True
                    btnClose.Focus()
            End Select
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD19T1801
    '# Created User: HUỲNH KHANH
    '# Created Date: 28/01/2015 02:40:50
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD19T1801(ByVal sTypeCodeID As String) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Luu bang D19T1801 neu copytoD19 bang 1" & vbCrLf)
        sSQL.Append("Insert Into D19T1801(")
        sSQL.Append("CCodeID, TypeCodeID, Disabled, CreateDate, " & vbCrLf)
        sSQL.Append("CreateUserID, LastModifyDate, LastModifyUserID, DescriptionU")
        sSQL.Append(") Values(" & vbCrLf)
        sSQL.Append(SQLString(txtACodeID.Text) & COMMA) 'CCodeID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(sTypeCodeID) & COMMA) 'TypeCodeID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append("GetDate()" & COMMA & vbCrLf) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append(SQLStringUnicode(txtDescription, True) & vbCrLf) 'DescriptionU, nvarchar[1000], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD19T1801
    '# Created User: HUỲNH KHANH
    '# Created Date: 28/01/2015 02:58:07
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD19T1801(ByVal sTypeCodeID As String) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Cap nhat bang D19T1801 neu copytoD19 bang 1" & vbCrLf)
        sSQL.Append("Update D19T1801 Set ")
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("DescriptionU = " & SQLStringUnicode(txtDescription, True)) 'nvarchar[1000], NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("CCodeID = " & SQLString(txtACodeID.Text) & " And ")
        sSQL.Append("TypeCodeID = " & SQLString(sTypeCodeID))

        Return sSQL
    End Function



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0041
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 26/10/2007 11:11:57
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0041() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T0041(")
        sSQL.Append("ACodeID, TypeCodeID, Type, DescriptionU, Disabled, ")
        sSQL.Append("CreateDate, CreateUserID, LastModifyDate, LastModifyUserID")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtACodeID.Text) & COMMA) 'ACodeID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcTypeCodeID.Text) & COMMA) 'TypeCodeID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString("A") & COMMA) 'Type, varchar[1], NULL
        sSQL.Append(SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'Description, varchar[250], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID)) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0041
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 26/10/2007 11:12:11
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0041() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0041 Set ")
        sSQL.Append("DescriptionU = " & SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
        sSQL.Append(" Where ")
        sSQL.Append("ACodeID = " & SQLString(txtACodeID.Text) & " And ")
        sSQL.Append("TypeCodeID = " & SQLString(tdbcTypeCodeID.Text))

        Return sSQL
    End Function

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        btnNext.Enabled = False
        btnSave.Enabled = True
        txtACodeID.Text = ""
        chkDisabled.Checked = False
        txtDescription.Text = ""
        txtACodeID.Focus()
    End Sub

End Class