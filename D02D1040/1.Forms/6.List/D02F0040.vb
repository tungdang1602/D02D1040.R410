'#-------------------------------------------------------------------------------------
'# Created Date: 24/10/2007 3:12:25 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 24/10/2007 3:12:25 PM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text

Public Class D02F0040

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F0040_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F0040_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadInfoGeneral()
        Loadlanguage()
        LoadForm()
        btnSave.Enabled = ReturnPermission(Me.Name) > EnumPermission.View
        InputbyUnicode(Me, gbUnicode)
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub LoadForm()
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, gbUnicode)

        Dim sSQL As String = ""
        sSQL = "Select CopyToD19, TypeCodeID, Disabled, MaxLength, " & IIf(geLanguage = EnumLanguage.Vietnamese, "VieTypeCodeName" & sUnicode, "EngTypeCodeName" & sUnicode).ToString & " As Description" & vbCrLf
        sSQL &= "From D02T0040 WITH(NOLOCK) Where Type='A' Order By TypeCodeID"
        Dim dtMain As DataTable = ReturnDataTable(sSQL)
        If dtMain.Rows.Count > 0 Then
            With dtMain
                txtTypeCodeID1.Text = .Rows(0).Item("TypeCodeID").ToString
                txtTypeCodeID2.Text = .Rows(1).Item("TypeCodeID").ToString
                txtTypeCodeID3.Text = .Rows(2).Item("TypeCodeID").ToString
                txtTypeCodeID4.Text = .Rows(3).Item("TypeCodeID").ToString
                txtTypeCodeID5.Text = .Rows(4).Item("TypeCodeID").ToString
                txtTypeCodeID6.Text = .Rows(5).Item("TypeCodeID").ToString
                txtTypeCodeID7.Text = .Rows(6).Item("TypeCodeID").ToString
                txtTypeCodeID8.Text = .Rows(7).Item("TypeCodeID").ToString
                txtTypeCodeID9.Text = .Rows(8).Item("TypeCodeID").ToString
                txtTypeCodeID10.Text = .Rows(9).Item("TypeCodeID").ToString

                txtDescription1.Text = .Rows(0).Item("Description").ToString
                txtDescription2.Text = .Rows(1).Item("Description").ToString
                txtDescription3.Text = .Rows(2).Item("Description").ToString
                txtDescription4.Text = .Rows(3).Item("Description").ToString
                txtDescription5.Text = .Rows(4).Item("Description").ToString
                txtDescription6.Text = .Rows(5).Item("Description").ToString
                txtDescription7.Text = .Rows(6).Item("Description").ToString
                txtDescription8.Text = .Rows(7).Item("Description").ToString
                txtDescription9.Text = .Rows(8).Item("Description").ToString
                txtDescription10.Text = .Rows(9).Item("Description").ToString

                chkDisabled1.Checked = L3Bool(.Rows(0).Item("Disabled"))
                chkDisabled2.Checked = L3Bool(.Rows(1).Item("Disabled"))
                chkDisabled3.Checked = L3Bool(.Rows(2).Item("Disabled"))
                chkDisabled4.Checked = L3Bool(.Rows(3).Item("Disabled"))
                chkDisabled5.Checked = L3Bool(.Rows(4).Item("Disabled"))
                chkDisabled6.Checked = L3Bool(.Rows(5).Item("Disabled"))
                chkDisabled7.Checked = L3Bool(.Rows(6).Item("Disabled"))
                chkDisabled8.Checked = L3Bool(.Rows(7).Item("Disabled"))
                chkDisabled9.Checked = L3Bool(.Rows(8).Item("Disabled"))
                chkDisabled10.Checked = L3Bool(.Rows(9).Item("Disabled"))

                txtMaxLength1.Text = .Rows(0).Item("MaxLength").ToString
                txtMaxLength2.Text = .Rows(1).Item("MaxLength").ToString
                txtMaxLength3.Text = .Rows(2).Item("MaxLength").ToString
                txtMaxLength4.Text = .Rows(3).Item("MaxLength").ToString
                txtMaxLength5.Text = .Rows(4).Item("MaxLength").ToString
                txtMaxLength6.Text = .Rows(5).Item("MaxLength").ToString
                txtMaxLength7.Text = .Rows(6).Item("MaxLength").ToString
                txtMaxLength8.Text = .Rows(7).Item("MaxLength").ToString
                txtMaxLength9.Text = .Rows(8).Item("MaxLength").ToString
                txtMaxLength10.Text = .Rows(9).Item("MaxLength").ToString

                chkCopyToD19_00.Checked = L3Bool(.Rows(0).Item("CopyToD19").ToString)
                chkCopyToD19_01.Checked = L3Bool(.Rows(1).Item("CopyToD19").ToString)
                chkCopyToD19_02.Checked = L3Bool(.Rows(2).Item("CopyToD19").ToString)
                chkCopyToD19_03.Checked = L3Bool(.Rows(3).Item("CopyToD19").ToString)
                chkCopyToD19_04.Checked = L3Bool(.Rows(4).Item("CopyToD19").ToString)
                chkCopyToD19_05.Checked = L3Bool(.Rows(5).Item("CopyToD19").ToString)
                chkCopyToD19_06.Checked = L3Bool(.Rows(6).Item("CopyToD19").ToString)
                chkCopyToD19_07.Checked = L3Bool(.Rows(7).Item("CopyToD19").ToString)
                chkCopyToD19_08.Checked = L3Bool(.Rows(8).Item("CopyToD19").ToString)
                chkCopyToD19_09.Checked = L3Bool(.Rows(9).Item("CopyToD19").ToString)
            End With
        End If
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowSave() Then Exit Sub
        btnSave.Enabled = False
        btnClose.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder

        sSQL.Append(SQLUpdateD02T0040(txtDescription1.Text, chkDisabled1.Checked, txtMaxLength1.Text, txtTypeCodeID1.Text, chkCopyToD19_00.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription2.Text, chkDisabled2.Checked, txtMaxLength2.Text, txtTypeCodeID2.Text, chkCopyToD19_01.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription3.Text, chkDisabled3.Checked, txtMaxLength3.Text, txtTypeCodeID3.Text, chkCopyToD19_02.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription4.Text, chkDisabled4.Checked, txtMaxLength4.Text, txtTypeCodeID4.Text, chkCopyToD19_03.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription5.Text, chkDisabled5.Checked, txtMaxLength5.Text, txtTypeCodeID5.Text, chkCopyToD19_04.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription6.Text, chkDisabled6.Checked, txtMaxLength6.Text, txtTypeCodeID6.Text, chkCopyToD19_05.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription7.Text, chkDisabled7.Checked, txtMaxLength7.Text, txtTypeCodeID7.Text, chkCopyToD19_06.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription8.Text, chkDisabled8.Checked, txtMaxLength8.Text, txtTypeCodeID8.Text, chkCopyToD19_07.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription9.Text, chkDisabled9.Checked, txtMaxLength9.Text, txtTypeCodeID9.Text, chkCopyToD19_08.Checked))
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLUpdateD02T0040(txtDescription10.Text, chkDisabled10.Checked, txtMaxLength10.Text, txtTypeCodeID10.Text, chkCopyToD19_09.Checked))
        sSQL.Append(vbCrLf)

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
            btnClose.Focus()

        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0040
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 24/10/2007 03:43:06
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0040(ByVal sDescription As String, ByVal bDisabled As Boolean, ByVal sMaxLength As String, ByVal sTypeCodeID As String, ByVal bCopyToD19 As Boolean) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0040 Set ")
        If geLanguage = EnumLanguage.Vietnamese Then
            sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(sDescription, gbUnicode, True) & COMMA) 'varchar[50], NULL
        ElseIf geLanguage = EnumLanguage.English Then
            sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(sDescription, gbUnicode, True) & COMMA) 'varchar[50], NULL
        End If
        sSQL.Append("Disabled = " & SQLNumber(bDisabled) & COMMA) 'bit, NOT NULL
        sSQL.Append("MaxLength = " & SQLNumber(sMaxLength) & COMMA) 'tinyint, NOT NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("CopyToD19 = " & SQLNumber(bCopyToD19)) 'tinyint, NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("TypeCodeID = " & SQLString(sTypeCodeID))

        Return sSQL
    End Function


    Private Sub txtMaxLength1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength1.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub


    Private Sub txtMaxLength2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength2.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtMaxLength3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength3.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtMaxLength4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength4.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtMaxLength5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength5.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtMaxLength6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength6.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtMaxLength7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength7.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtMaxLength8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength8.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtMaxLength9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength9.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtMaxLength10_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength10.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub
    Private Function AllowSave() As Boolean
        If txtMaxLength1.Text.Trim <> "" Then
            If CInt(txtMaxLength1.Text) < 1 Or L3Int(txtMaxLength1.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength1.Text = "1"
                txtMaxLength1.Focus()
                Return False
            End If
        End If

        If txtMaxLength2.Text.Trim <> "" Then
            If CInt(txtMaxLength2.Text) < 1 Or L3Int(txtMaxLength2.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength2.Text = "1"
                txtMaxLength2.Focus()
                Return False
            End If
        End If
        If txtMaxLength3.Text.Trim <> "" Then
            If CInt(txtMaxLength3.Text) < 1 Or L3Int(txtMaxLength3.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength3.Text = "1"
                txtMaxLength3.Focus()
                Return False
            End If
        End If
        If txtMaxLength4.Text.Trim <> "" Then
            If CInt(txtMaxLength4.Text) < 1 Or L3Int(txtMaxLength4.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength4.Text = "1"
                txtMaxLength4.Focus()
                Return False

            End If
        End If
        If txtMaxLength5.Text.Trim <> "" Then
            If CInt(txtMaxLength5.Text) < 1 Or L3Int(txtMaxLength5.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength5.Text = "1"
                txtMaxLength5.Focus()
                Return False
            End If
        End If
        If txtMaxLength6.Text.Trim <> "" Then
            If CInt(txtMaxLength6.Text) < 1 Or L3Int(txtMaxLength6.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength6.Text = "1"
                txtMaxLength6.Focus()
                Return False
            End If
        End If
        If txtMaxLength7.Text.Trim <> "" Then
            If CInt(txtMaxLength7.Text) < 1 Or L3Int(txtMaxLength7.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength7.Text = "1"
                txtMaxLength7.Focus()
                Return False
            End If
        End If
        If txtMaxLength8.Text.Trim <> "" Then
            If CInt(txtMaxLength8.Text) < 1 Or L3Int(txtMaxLength8.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength8.Text = "1"
                txtMaxLength8.Focus()
                Return False

            End If
        End If
        If txtMaxLength9.Text.Trim <> "" Then
            If CInt(txtMaxLength9.Text) < 1 Or L3Int(txtMaxLength9.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength9.Text = "1"
                txtMaxLength9.Focus()
                Return False
            End If
        End If
        If txtMaxLength10.Text.Trim <> "" Then
            If CInt(txtMaxLength10.Text) < 1 Or L3Int(txtMaxLength10.Text) > 50 Then
                D99C0008.MsgL3(rL3("Chieu_dai_khong_hop_leU"))
                txtMaxLength10.Text = "1"
                txtMaxLength10.Focus()
                Return False
            End If
        End If

        Return True
    End Function

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Dinh_nghia_ma_phan_tich_-_D02F0040") & UnicodeCaption(gbUnicode) '˜Ünh nghÚa mº ph¡n tÛch - D02F0040
        '================================================================ 
        lbl1.Text = "(<=50)" '(<=20)
        lblLength.Text = rL3("Chieu_dai") 'Chiều dài
        lblDisabled.Text = rL3("Khong_su_dung") 'Không sử dụng
        lblName.Text = rL3("Ten_loai_ma_phan_tich") 'Tên loại mã phân tích
        lblCode.Text = rL3("Ma") 'Mã
        lblCopyToD19.Text = rL3("Chuyen_Ma_phan_tich_sang_CPTT") 'Chuyển Mã phân tích sang CPTT

        '================================================================ 
        btnSave.Text = rL3("_Luu") '&Lưu
        btnClose.Text = rL3("Do_ng") 'Đó&ng
        '================================================================ 
    End Sub

End Class