Imports System.Collections
Imports System.IO
Imports System
Public Class D02F1031
    Private sPathImage As String = ""
#Region "Variables of IGE"
    Private _S1 As String = ""
    Private _S2 As String = ""
    Private _S3 As String = ""
    Private _OutputOrder As String = ""
    Private _OutputLength As Integer = 0
    Private _AssetSeperated As Boolean = False
    Private _Seperator As String = ""
    Private _TableName As String = "D02T0001"
    Private bAuto As Integer


#End Region

#Region "Const of tdbgDetail - Total of Columns: 12"
    Private Const COL_OrderNum As Integer = 0          ' STT
    Private Const COL_EquipmentID As Integer = 1       ' Mã thiết bị đính kèm
    Private Const COL_EquipmentName As Integer = 2     ' Tên thiết bị đính kèm
    Private Const COL_EquipmentQuantity As Integer = 3 ' Số lượng
    Private Const COL_UnitPrice As Integer = 4         ' Đơn giá
    Private Const COL_EquipmentValue As Integer = 5    ' Giá trị
    Private Const COL_TaxAmount As Integer = 6        ' Tiền thuế GTGT
    Private Const COL_AcceptanceTime As Integer = 7    ' Thời gian nghiệm thu
    Private Const COL_PurchaseDate As Integer = 8      ' Ngày mua
    Private Const COL_ObjectTypeID As Integer = 9      ' Loại phòng ban
    Private Const COL_ObjectID As Integer = 10         ' Mã phòng ban
    Private Const COL_Notes As Integer = 11            ' Ghi chú
#End Region

#Region "Property"
    Private mAssetID As String = ""
    Public WriteOnly Property AssetID() As String
        Set(ByVal value As String)
            mAssetID = value
        End Set
    End Property

    Private mAssetID_D02F1031 As String = ""
    Public ReadOnly Property AssetID_D02F1031() As String
        Get
            Return mAssetID_D02F1031
        End Get
    End Property

    Private _Completed As Boolean = False
    Public WriteOnly Property Completed() As Boolean
        Set(ByVal value As Boolean)
            _Completed = value
        End Set
    End Property

    Private _locationID As String
    Public WriteOnly Property LocationID() As String
        Set(ByVal Value As String)
            _locationID = Value
        End Set
    End Property

    Private _sAssetID As String
    Public WriteOnly Property sAssetID() As String
        Set(ByVal Value As String)
            _sAssetID = Value
        End Set
    End Property

    Private _sMethodID As String
    Public WriteOnly Property sMethodID() As String
        Set(ByVal Value As String)
            _sMethodID = Value
        End Set
    End Property

    Private _bFormD02F0087 As Boolean = False
    Public Property bFormD02F0087() As Boolean
        Get
            Return _bFormD02F0087
        End Get
        Set(ByVal Value As Boolean)
            _bFormD02F0087 = Value
        End Set
    End Property
#End Region

    Private dtObjectID As DataTable
    'Private dtSupplierID As DataTable
    'Private dtObjectID2 As DataTable
    Private sCreateUserID As String
    Private sCreateDate As String
    '   Dim dtSystem As DataTable
    'Bổ sung Field Unicode
    Dim sUnicode As String = ""
    Dim sLanguage As String = ""
    Dim sDefaultIGEMethodID As String = ""
    Dim clsFilterCombo As Lemon3.Controls.FilterCombo
    Private _isTools As Boolean
    Public WriteOnly Property IsTools() As Boolean
        Set(ByVal Value As Boolean)
            _isTools = Value
        End Set
    End Property
    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

    Private _sSQLD91T1001_SaveLastKey As String = "" '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
    Public WriteOnly Property sSQLD91T1001_SaveLastKey As String
        Set(value As String)
            _sSQLD91T1001_SaveLastKey = value
        End Set
    End Property

    Dim bLoadFormState As Boolean = False
    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            bLoadFormState = True
            LoadInfoGeneral()
            _FormState = value
            tdbgDetail_NumberFormat()
            UnicodeAllString(sUnicode, sLanguage, gbUnicode)
            VisibleIGEMethodID()

            'ID 78424 12/08/2015
            clsFilterCombo = New Lemon3.Controls.FilterCombo
            clsFilterCombo.CheckD91 = True 'Giá trị mặc định True: kiểm tra theo DxxFormat.LoadFormNotINV. Ngược lại luôn luôn Filter dạng mới (dùng cho Novaland)

            clsFilterCombo.AddPairObject(tdbcObjectTypeID, tdbcObjectID) 'Tab 1: Bộ phận tiếp nhận
            clsFilterCombo.AddPairObject(tdbcObjectTypeID2, tdbcObjectID2) 'Tab 1: Bộ phận quản lý
            clsFilterCombo.AddPairObject(tdbcSupplierOTID, tdbcSupplierID) 'Tab 1: Nhà cung cấp
            clsFilterCombo.AddPairObject(tdbcObjectTypeID6, tdbcObjectID6) 'Tab 6: Bộ phận tiếp nhận 
            clsFilterCombo.AddPairObject(tdbcManagementObTypeID6, tdbcManagementObID6) 'Tab 6: Bộ phận quản lý
            clsFilterCombo.AddPairObject(tdbcSupplierOTIDID6, tdbcSupplierIDID6) 'Tab 6: Nhà cung cấp

            clsFilterCombo.UseFilterComboObjectID()

            clsFilterCombo.UseFilterCombo(tdbcEmployeeID, tdbcLocationID, tdbcAssetAccountID, tdbcDepAccountID, tdbcAssetConditionName, tdbcAcode01ID, tdbcAcode02ID, tdbcAcode03ID, tdbcAcode04ID, tdbcAcode05ID, tdbcAcode06ID, tdbcAcode07ID, tdbcAcode08ID, tdbcAcode09ID, tdbcAcode10ID, tdbcReceiverID, tdbcLocationIDID6, tdbcUnitID, tdbcAccountID, tdbcMethodIDCCDC)
            clsFilterCombo.UseFilterCombo(tdbcAssetS1ID, tdbcAssetS2ID, tdbcAssetS3ID)

            LoadTDBCombo()
            LoadCaption()
            '    GetInfoSystemDefault()
            GetAutoAssetInfo()
            VisibleIsTools() 'Ẩn hiện check box công cụ dụng cụ
            Select Case _FormState
                Case EnumFormState.FormAdd
                    Reload()
                    LoadAddNew()
                    txtAssetID.Focus()
                Case EnumFormState.FormEdit
                    Reload()
                    LoadEdit()
                    LoadEdit_Data()
                    btnSave.Enabled = True
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                Case EnumFormState.FormView
                    Reload()
                    LoadEdit()
                    LoadEdit_Data()
                    btnSave.Enabled = False
                    btnNext.Enabled = False
                    tdbcAssetS1ID.Enabled = False
                    tdbcAssetS2ID.Enabled = False
                    tdbcAssetS3ID.Enabled = False
                    ReadOnlyControl(txtAssetID)
                    btnSetNewKey.Enabled = False
                    grp03.Enabled = False
                    btnSave.Enabled = False
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                Case EnumFormState.FormCopy
                    ' ''GetAutoAssetInfo()
                    Reload()
                    LoadEdit()
                    LoadEdit_Data()
                    btnSave.Enabled = True
                    btnNext.Enabled = False
            End Select
        End Set
    End Property

    Private Sub VisibleIsTools()
        chkIsTools.Visible = D02Systems.IsAssetIDForD02D43
    End Sub

    '    Private Sub GetInfoSystemDefault()
    '        Dim sSQL As String = ""
    '        sSQL = "Select DefAssetAccountID,DefDepreciationAccountID,MethodID,MethodEndID From D02T0000 WITH(NOLOCK)"
    '        dtSystem = ReturnDataTable(sSQL)
    '    End Sub

    Private Sub GetAutoAssetInfo()
        If D02Systems.AssetAuto = 0 Then
            tdbcAssetS1ID.Enabled = False
            tdbcAssetS2ID.Enabled = False
            tdbcAssetS3ID.Enabled = False
            UnReadOnlyControl(txtAssetID, True)
            btnSetNewKey.Enabled = False
            bAuto = 0

        ElseIf D02Systems.AssetAuto = 1 Then
            bAuto = 1 '31/7/2019, id 122577-Lỗi không sinh mã tài sản cố định
            ReadOnlyControl(txtAssetID)
            _OutputOrder = D02Systems.AssetOutputOrder
            _OutputLength = D02Systems.AssetOutputLength
            _AssetSeperated = D02Systems.AssetSeperated
            _Seperator = D02Systems.AssetSeperator
            If D02Systems.AssetS1Enabled Then
                tdbcAssetS1ID.Enabled = True
                tdbcAssetS1ID.Text = D02Systems.AssetS1Default
                tdbcAssetS1ID_SelectedValueChanged(Nothing, Nothing)
            Else
                tdbcAssetS1ID.Enabled = False
            End If
            If D02Systems.AssetS2Enabled Then
                tdbcAssetS2ID.Enabled = True
                tdbcAssetS2ID.Text = D02Systems.AssetS2Default
                tdbcAssetS2ID_SelectedValueChanged(Nothing, Nothing)
            Else
                tdbcAssetS2ID.Enabled = False
            End If
            If D02Systems.AssetS3Enabled Then
                tdbcAssetS3ID.Enabled = True
                tdbcAssetS3ID.Text = D02Systems.AssetS3Default
                tdbcAssetS3ID_SelectedValueChanged(Nothing, Nothing)
            Else
                tdbcAssetS3ID.Enabled = False
            End If
            btnSetNewKey.Enabled = True
            'bAuto = 1 '31/7/2019, id 122577-Lỗi không sinh mã tài sản cố định
        Else
            bAuto = 2
            'btnSetNewKey.Enabled = True
         
            btnSetNewKey.Enabled = False
            If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormCopy Then
                'ID 92099 29.11.2016
                tdbcIGEMethodID.SelectedValue = _sMethodID
                txtAssetID.Text = _sAssetID
                tdbcIGEMethodID.Enabled = Not _bFormD02F0087

                '24/3/2017, Trần Hoàng Anh: id 95287-Tìm nguyên nhân lỗi SQL khi tạo mã TS để fix triệt để
                'txtAssetID.Enabled = Not _bFormD02F0087
                ReadOnlyControl(txtAssetID)

                btnSetNewKey.Enabled = Not _bFormD02F0087
            Else
                btnSetNewKey.Enabled = True
            End If

        End If
    End Sub

    Private Sub LoadAddNew()
        c1datePurchaseDate.Value = Date.Today
        c1datePeriod.Value = giTranMonth.ToString() & "/" & giTranYear.ToString()
        c1dateDepPeriod.Value = giTranMonth.ToString() & "/" & giTranYear.ToString()
        c1dateTranDate.Value = giTranMonth.ToString() & "/" & giTranYear.ToString()
        grp03.Enabled = False
        btnNext.Enabled = False

        tdbcAssetAccountID.Text = D02Systems.DefAssetAccountID
        tdbcDepAccountID.Text = D02Systems.DefDepreciationAccountID
        '        Else
        '        tdbcAssetAccountID.Text = ""
        '        tdbcDepAccountID.Text = ""
        '        End If

        If D02Systems.MethodID <> "" Then '8/2/2022, Đặng Thân Yến Nhi:id 214826-[LAFOOCO] D02 - Mặc định các thông tin tài chính khi tạo mới Danh mục TSCĐ theo thiết lập hệ thống
            tdbcMethodID.SelectedValue = D02Systems.MethodID
        Else
            tdbcMethodID.SelectedIndex = 0
        End If

        If D02Systems.MethodEndID <> "" Then '8/2/2022, Đặng Thân Yến Nhi:id 214826-[LAFOOCO] D02 - Mặc định các thông tin tài chính khi tạo mới Danh mục TSCĐ theo thiết lập hệ thống
            tdbcMethodEndID.SelectedValue = D02Systems.MethodEndID
        Else
            tdbcMethodEndID.SelectedIndex = 0
        End If

        tdbcAssignmentTypeID.SelectedIndex = 0
        UnReadOnlyControl(True, tdbcIGEMethodID)
        tdbcUnitID.SelectedValue = "-1"
        tdbcAccountID.SelectedValue = "-1"
        tdbcMethodIDCCDC.SelectedValue = "-1"
        txtSetupVoucherID.Text = ""
        c1dateSetupDate.Value = ""
        tdbcObjectTypeID6.SelectedValue = ""
        tdbcObjectID6.SelectedValue = ""
        txtObjectName6.Text = ""

        tdbcManagementObTypeID6.SelectedValue = ""
        tdbcManagementObID6.SelectedValue = ""
        txtManagementObName.Text = ""
        cneOQuantity.Value = ""
        txtCQuantity.Text = ""
        If iAssetAuto = 2 Then
            tdbcIGEMethodID.Focus()
        Else
            tdbcAssetS1ID.Focus()
        End If
        txtSetupVoucherID.Enabled = True
        c1dateSetupDate.Enabled = True
        cneOQuantity.Enabled = True
        txtCQuantity.Enabled = False
        tdbcChargeObjType.SelectedIndex = 0 'ID : 252774
        tdbcChargeObjType.Enabled = True 'ID : 252774
    End Sub

    Dim bIsDepreciated As Boolean
    Private Sub LoadEdit()
        grp01.Enabled = True
        grp02.Enabled = True
        grp03.Enabled = False
        tdbcAssetAccountID.Enabled = True
        tdbcDepAccountID.Enabled = True

        If _Completed Then
            Dim sSQL As String = ""
            tdbcObjectTypeID.Enabled = False
            tdbcObjectID.Enabled = False
            tdbcEmployeeID.Enabled = False
            tdbcAssetAccountID.Enabled = False
            tdbcDepAccountID.Enabled = False
            tdbcLocationID.Enabled = False
            tdbcObjectTypeID2.Enabled = False
            tdbcObjectID2.Enabled = False
            grp01.Enabled = False

            txtSetupVoucherID.Enabled = Not _Completed
            c1dateSetupDate.Enabled = Not _Completed
            cneOQuantity.Enabled = Not _Completed
            tdbcChargeObjType.Enabled = Not _Completed 'ID : 252774
            tdbcObjectTypeID6.Enabled = Not _Completed
            tdbcObjectID6.Enabled = Not _Completed
            tdbcReceiverID.Enabled = Not _Completed
            tdbcManagementObTypeID6.Enabled = Not _Completed
            tdbcManagementObID6.Enabled = Not _Completed
            tdbcLocationIDID6.Enabled = Not _Completed 'ID : 252774

            sSQL = SQLStoreD02P0613()
            Dim dt As DataTable = ReturnDataTable(sSQL)
            bIsDepreciated = False
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("Status").ToString() = "1" Then
                    'MsgBox(dt.Rows(0)("Message").ToString(), MsgBoxStyle.Information, "Th¤ng bÀo")
                    grp02.Enabled = False
                    bIsDepreciated = True
                End If
            End If
        End If

        If _FormState <> EnumFormState.FormCopy Then
            tdbcAssetS1ID.Enabled = False
            tdbcAssetS2ID.Enabled = False
            tdbcAssetS3ID.Enabled = False
            ReadOnlyControl(txtAssetID)
            btnSetNewKey.Enabled = False
            btnNext.Enabled = False
            grp03.Enabled = False
            chkIsTools.Checked = False

            'Ẩn cách phương pháp tạo mã khi sửa
            tdbcAssetS1ID.Visible = False
            tdbcAssetS2ID.Visible = False
            tdbcAssetS3ID.Visible = False
            tdbcIGEMethodID.Visible = False
            txtAssetID.Location = New Point(75, 13)
            txtAssetID.Width = 432
            ReadOnlyControl(chkIsTools)
        Else
            If iAssetAuto = 2 Then 'Trường hợp hiện phương pháp tạo mã
                tdbcAssetS1ID.Visible = False
                tdbcAssetS2ID.Visible = False
                tdbcAssetS3ID.Visible = False
                tdbcIGEMethodID.Visible = True
                UnReadOnlyControl(True, tdbcIGEMethodID)
                ReadOnlyControl(txtAssetID)
            End If

            If bAuto = 1 Then
                '    '    tdbcAssetS1ID.Enabled = True
                '    '    tdbcAssetS2ID.Enabled = True
                '    '    tdbcAssetS3ID.Enabled = True
               ReadOnlyControl(txtAssetID)
                'btnSetNewKey.Enabled = True
            Else
                UnReadOnlyControl(txtAssetID, True)
                btnSetNewKey.Enabled = False
                '    '    tdbcAssetS1ID.Enabled = False
                '    '    tdbcAssetS2ID.Enabled = False
                '    '    tdbcAssetS3ID.Enabled = False
            End If
        End If
    End Sub

    Private Sub LoadEdit_Data()

        'Dim sSQL As New StringBuilder()
        'sSQL.Append("SELECT T01.FAString01" & UnicodeJoin(gbUnicode) & " as FAString01,")
        'sSQL.Append("T01.FAString02" & UnicodeJoin(gbUnicode) & " as FAString02,")
        'sSQL.Append("T01.FAString02" & UnicodeJoin(gbUnicode) & " as FAString02,")
        'sSQL.Append("T01.FAString03" & UnicodeJoin(gbUnicode) & " as FAString03,")
        'sSQL.Append("T01.FAString04" & UnicodeJoin(gbUnicode) & " as FAString04,")
        'sSQL.Append("T01.FAString05" & UnicodeJoin(gbUnicode) & " as FAString05,")
        'sSQL.Append("T01.FAString06" & UnicodeJoin(gbUnicode) & " as FAString06,")
        'sSQL.Append("T01.FAString07" & UnicodeJoin(gbUnicode) & " as FAString07,")
        'sSQL.Append("T01.FAString08" & UnicodeJoin(gbUnicode) & " as FAString08,")
        'sSQL.Append("T01.FAString09" & UnicodeJoin(gbUnicode) & " as FAString09,")
        'sSQL.Append("T01.FAString10" & UnicodeJoin(gbUnicode) & " as FAString10,")
        'sSQL.Append("T01.FANum01,T01.FANum02,T01.FANum03,T01.FANum04,T01.FANum05,T01.FANum06,T01.FANum07,T01.FANum08,T01.FANum09,T01.FANum10,")
        'sSQL.Append("T01.FADate01,T01.FADate02,T01.FADate03,T01.FADate04,T01.FADate05,T01.FADate06,T01.FADate07,T01.FADate08,T01.FADate09,T01.FADate10,")
        'sSQL.Append("T01.AssetID, T01.AssetName, T01.AssetNameU, T01.ShortName, T01.ShortNameU, T01.Disabled, T01.DeprTableID, ")
        'sSQL.Append("T01.AssetS1ID, T01.AssetS2ID, T01.AssetS3ID, " & vbCrLf)
        'sSQL.Append("ltrim(rtrim(str(T01.TranMonth)))+'/'+ltrim(rtrim(str(T01.TranYear))) as AssetPeriod, " & vbCrLf)
        'sSQL.Append(" N19.ObjectTypeID, N19.ObjectID, N19.EmployeeID as AssetUserID, N19.FullName, N19.FullNameU, " & vbCrLf)
        'sSQL.Append(" N19.ServiceLife, N19.DepreciatedPeriod, T01.AssetAccountID, T01.DepAccountID, " & vbCrLf)
        'sSQL.Append(" N19.CurrentCost as ConvertedAmount, CASE WHEN N19.DepreciationAmount = 0  " & vbCrLf)
        'sSQL.Append(" THEN N19.DepreciatedAmount ELSE N19.DepreciationAmount END AS DepreciationAmount , " & vbCrLf)
        'sSQL.Append(" N19.RemainAmount, T01.IsCompleted,T01.IsRevalued, T01.IsDisposed, T01.UnitName, T01.UnitNameU," & vbCrLf)
        'sSQL.Append(" T01.Index1, T01.Index2, T01.Index3, T01.Index4,T01.Index5, T01.Index6," & vbCrLf)
        'sSQL.Append(" T01.ACode01ID, T01.ACode02ID, T01.ACode03ID, T01.ACode04ID, " & vbCrLf)
        'sSQL.Append(" T01.ACode05ID, T01.ACode06ID, T01.ACode07ID, T01.ACode08ID, " & vbCrLf)
        'sSQL.Append(" T01.ACode09ID, T01.ACode10ID, T01.Percentage, N19.CurrentLTDDepreciation as AmountDepreciation," & vbCrLf)
        'sSQL.Append(" T01.MethodID, T01.MethodEndID, T01.CountryID, T01.EmployeeID, " & vbCrLf)
        'sSQL.Append(" T01.MadeYear, T01.SeriNo, T01.Specification, T01.SpecificationU, T01.Notes, T01.NotesU, " & vbCrLf)
        'sSQL.Append(" ltrim(rtrim(str(T01.UseMonth)))+'/'+ltrim(rtrim(str(T01.UseYear))) as UsePeriod, " & vbCrLf)
        'sSQL.Append(" ltrim(rtrim(str(T01.DepMonth)))+'/'+ltrim(rtrim(str(T01.DepYear))) as DepPeriod, " & vbCrLf)
        'sSQL.Append(" T01.AssetNo, T01.Version, T01.AssetTag, T01.AssetTagU, T01.Tool, T01.ToolU, AssignmentTypeID , " & vbCrLf)
        'sSQL.Append(" T01.Maintainable, T01.SupplierOTID, T01.SupplierID, T01.PurchaseDate, T01.CreateUserID, T01.DepDate, T01.CreateDate " & vbCrLf)
        'sSQL.Append(" FROM D02T0001 as T01 WITH(NOLOCK) INNER JOIN D02N0019(" & giTranMonth & "," & giTranYear & ") as N19 ON T01.AssetID = N19.AssetID AND ")
        'sSQL.Append("T01.DivisionID = N19.DivisionID WHERE T01.AssetID = " & SQLString(mAssetID) & "AND T01.DivisionID = " & SQLString(gsDivisionID) & vbCrLf)

        Dim sSQL As String = SQLStoreD02P1032()
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                If _FormState <> EnumFormState.FormCopy Then
                    tdbcAssetS1ID.Text = .Item("AssetS1ID").ToString()
                    tdbcAssetS2ID.Text = .Item("AssetS2ID").ToString()
                    tdbcAssetS3ID.Text = .Item("AssetS3ID").ToString()
                    txtAssetID.Text = .Item("AssetID").ToString()
                    txtAssetName.Text = .Item("AssetName").ToString()
                Else
                    tdbcAssetS1ID.Text = ""
                    tdbcAssetS2ID.Text = ""
                    tdbcAssetS3ID.Text = ""
                    txtAssetName.Text = ""
                End If

                txtNotes.Text = .Item("Notes").ToString()
                'LoadTDBCombo()
                tdbcObjectTypeID.Text = .Item("ObjectTypeID").ToString()
                tdbcObjectID.Text = .Item("ObjectID").ToString()
                tdbcObjectTypeID2.SelectedValue = .Item("ManagementObjTypeID").ToString()
                'LoadtdbcObjectID(tdbcObjectID2, ReturnValueC1Combo(tdbcObjectTypeID2))
                tdbcObjectID2.SelectedValue = .Item("ManagementObjID").ToString()
                tdbcEmployeeID.SelectedValue = .Item("AssetUserID").ToString()
                txtShortName.Text = .Item("ShortName").ToString()
                txtAssetTag.Text = .Item("AssetTag").ToString()
                tdbcSupplierOTID.Text = .Item("SupplierOTID").ToString()
                tdbcSupplierID.Text = .Item("SupplierID").ToString()
                c1datePurchaseDate.Value = .Item("PurchaseDate").ToString()
                chkMaintainable.Checked = CType(.Item("Maintainable"), Boolean)
                If _FormState <> EnumFormState.FormCopy Then
                    tdbcAssetAccountID.Text = .Item("AssetAccountID").ToString()
                    tdbcDepAccountID.Text = .Item("DepAccountID").ToString()
                    c1datePeriod.Value = .Item("UsePeriod").ToString() 'Format(.Item("UseMonth"), "0#") & "/" & .Item("UseYear").ToString()
                    c1dateDepPeriod.Value = .Item("DepPeriod").ToString()  'Format(.Item("DepMonth"), "0#") & "/" & .Item("DepYear").ToString()
                    c1dateTranDate.Value = .Item("AssetPeriod").ToString() 'Format(.Item("TranMonth"), "0#") & "/" & .Item("TranYear").ToString()
                    c1dateDepDate.Value = .Item("DepDate").ToString()
                    tdbcMethodID.Text = .Item("MethodID").ToString()
                    tdbcMethodEndID.Text = .Item("MethodEndID").ToString()
                    tdbcDeprTableID.Text = .Item("DeprTableID").ToString()
                    tdbcAssignmentTypeID.Text = .Item("AssignmentTypeID").ToString()
                    txtConvertedAmount.Text = .Item("ConvertedAmount").ToString()
                    txtRemainAmount.Text = .Item("RemainAmount").ToString()
                    txtDepreciatedPeriod.Text = .Item("DepreciatedPeriod").ToString()
                    txtDepreciationAmount.Text = .Item("DepreciationAmount").ToString()
                    txtAmountDepreciation.Text = .Item("AmountDepreciation").ToString()
                    txtServiceLife.Text = .Item("ServiceLife").ToString()
                    txtPercentage.Text = .Item("Percentage").ToString()
                Else

                    tdbcAssetAccountID.Text = D02Systems.DefAssetAccountID
                    tdbcDepAccountID.Text = D02Systems.DefDepreciationAccountID
                    '                    Else
                    '                    tdbcAssetAccountID.Text = ""
                    '                    tdbcDepAccountID.Text = ""
                    '                End If
                    c1datePeriod.Value = giTranMonth.ToString() & "/" & giTranYear.ToString()
                    c1dateDepPeriod.Value = giTranMonth.ToString() & "/" & giTranYear.ToString()
                    c1dateTranDate.Value = giTranMonth.ToString() & "/" & giTranYear.ToString()
                    tdbcMethodID.SelectedIndex = 0
                    tdbcMethodEndID.SelectedIndex = 0
                    tdbcDeprTableID.SelectedValue = ""
                    txtDeprTableName.Text = ""
                    tdbcAssignmentTypeID.SelectedIndex = 0
                    txtConvertedAmount.Text = ""
                    txtRemainAmount.Text = ""
                    txtDepreciatedPeriod.Text = ""
                    txtDepreciationAmount.Text = ""
                    txtAmountDepreciation.Text = ""
                    txtServiceLife.Text = ""
                    txtPercentage.Text = ""
                    grp03.Enabled = False
                    tdbcAssetS1ID.Focus()
                End If

                txtSpecification.Text = .Item("Specification").ToString()
                txtCountryID.Text = .Item("CountryID").ToString()
                txtMadeYear.Text = .Item("MadeYear").ToString()
                txtSeriNo.Text = .Item("SeriNo").ToString()
                txtVersion.Text = .Item("Version").ToString()
                txtAssetNo.Text = .Item("AssetNo").ToString()
                txtUnitName.Text = .Item("UnitName").ToString()
                txtTool.Text = .Item("Tool").ToString()
                txtIndex1.Text = .Item("Index1").ToString()
                txtIndex2.Text = .Item("Index2").ToString()
                txtIndex3.Text = .Item("Index3").ToString()
                txtIndex4.Text = .Item("Index4").ToString()
                txtIndex5.Text = .Item("Index5").ToString()
                txtIndex6.Text = .Item("Index6").ToString()
                tdbcAcode01ID.SelectedValue = .Item("Acode01ID").ToString()
                tdbcAcode02ID.SelectedValue = .Item("Acode02ID").ToString()
                tdbcAcode03ID.SelectedValue = .Item("Acode03ID").ToString()
                tdbcAcode04ID.SelectedValue = .Item("Acode04ID").ToString()
                tdbcAcode05ID.SelectedValue = .Item("Acode05ID").ToString()
                tdbcAcode06ID.SelectedValue = .Item("Acode06ID").ToString()
                tdbcAcode07ID.SelectedValue = .Item("Acode07ID").ToString()
                tdbcAcode08ID.SelectedValue = .Item("Acode08ID").ToString()
                tdbcAcode09ID.SelectedValue = .Item("Acode09ID").ToString()
                tdbcAcode10ID.SelectedValue = .Item("Acode10ID").ToString()
                sCreateUserID = .Item("CreateUserID").ToString()
                sCreateDate = .Item("CreateDate").ToString()
                txtFAString01.Text = .Item("FAString01").ToString()
                txtFAString02.Text = .Item("FAString02").ToString()
                txtFAString03.Text = .Item("FAString03").ToString()
                txtFAString04.Text = .Item("FAString04").ToString()
                txtFAString05.Text = .Item("FAString05").ToString()
                txtFAString06.Text = .Item("FAString06").ToString()
                txtFAString07.Text = .Item("FAString07").ToString()
                txtFAString08.Text = .Item("FAString08").ToString()
                txtFAString09.Text = .Item("FAString09").ToString()
                txtFAString10.Text = .Item("FAString10").ToString()
                txtFANum01.Text = Double.Parse((.Item("FANum01").ToString())).ToString
                txtFANum02.Text = Double.Parse((.Item("FANum02").ToString())).ToString
                txtFANum03.Text = Double.Parse((.Item("FANum03").ToString())).ToString
                txtFANum04.Text = Double.Parse((.Item("FANum04").ToString())).ToString
                txtFANum05.Text = Double.Parse((.Item("FANum05").ToString())).ToString
                txtFANum06.Text = Double.Parse((.Item("FANum06").ToString())).ToString
                txtFANum07.Text = Double.Parse((.Item("FANum07").ToString())).ToString
                txtFANum08.Text = Double.Parse((.Item("FANum08").ToString())).ToString
                txtFANum09.Text = Double.Parse((.Item("FANum09").ToString())).ToString
                txtFANum10.Text = Double.Parse((.Item("FANum10").ToString())).ToString
                c1dateFADate01.Value = .Item("FADate01").ToString()
                c1dateFADate02.Value = .Item("FADate02").ToString()
                c1dateFADate03.Value = .Item("FADate03").ToString()
                c1dateFADate04.Value = .Item("FaDate04").ToString()
                c1dateFADate05.Value = .Item("FADate05").ToString()
                c1dateFADate06.Value = .Item("FADate06").ToString()
                c1dateFADate07.Value = .Item("FADate07").ToString()
                c1dateFADate08.Value = .Item("FADate08").ToString()
                c1dateFADate09.Value = .Item("FADate09").ToString()
                c1dateFADate10.Value = .Item("FADate10").ToString()

                tdbcUnitID.SelectedValue = .Item("UnitID").ToString()
                tdbcAccountID.SelectedValue = .Item("AccountID").ToString()
                tdbcMethodIDCCDC.SelectedValue = .Item("MethodID").ToString()
                tdbcLocationID.SelectedValue = .Item("LocationID").ToString()
                c1dateMaintainDate.Value = SQLDateShow(.Item("MaintainDate"))
                tdbcAssetConditionName.SelectedValue = .Item("AssetConditionID")

                tdbcObjectTypeID6.SelectedValue = .Item("ObjectTypeID").ToString()

                tdbcObjectID6.SelectedValue = .Item("ObjectID").ToString()
                tdbcManagementObTypeID6.SelectedValue = .Item("ManagementObjTypeID").ToString()

                tdbcManagementObID6.SelectedValue = .Item("ManagementObjID").ToString()
                txtSetupVoucherID.Text = .Item("SetupVoucherID").ToString()
                c1dateSetupDate.Value = SQLDateShow(.Item("SetupDate"))
                cneOQuantity.Value = Number(.Item("OQuantity"), DxxFormat.D07_QuantityDecimals)
                txtCQuantity.Text = SQLNumber(.Item("CQuantity"), DxxFormat.D07_QuantityDecimals)

                tdbcReceiverID.SelectedValue = .Item("ToolReceiverID").ToString
                tdbcLocationIDID6.SelectedValue = .Item("ToolLocationID").ToString
                tdbcSupplierOTIDID6.SelectedValue = .Item("ToolSupplierOTID").ToString
                tdbcSupplierIDID6.SelectedValue = .Item("ToolSupplierID").ToString
                tdbcChargeObjType.SelectedValue = .Item("ChargeObjType").ToString 'ID : 252774
            End With
            chkIsTools.Checked = _isTools
            If chkIsTools.Checked = False Then chkIsTools_CheckedChanged(Nothing, Nothing) '5/9/2018, id 112182-Lỗi mất thông tin phòng ban và đơn vị khi truy vấn mã CCDC tại D02
        End If

        'Load LocationID
        tdbcLocationID.SelectedValue = _locationID
        LoadImage()

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1032
    '# Created User: HUỲNH KHANH
    '# Created Date: 10/10/2014 04:38:02
    '#---------------------------------------------------------------------------------------------------
    'Private Function SQLStoreD02P1032() As String
    '    Dim sSQL As String = ""
    '    sSQL &= ("-- Do nguon khi sua" & vbCrlf)
    '    sSQL &= "Exec D02P1032 "
    '    sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
    '    sSQL &= SQLString(mAssetID) & COMMA 'AssetID, varchar[20], NOT NULL
    '    sSQL &= SQLNumber(_isTools) & COMMA 'IsTools, tinyint, NOT NULL
    '    sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
    '    Return sSQL
    'End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1032
    '# Created User: HUỲNH KHANH
    '# Thảo yêu cầu bổ sung thêm  TranMonth, TranYear
    '# Created Date: 22/10/2014 04:10:27
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1032() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Đổ nguồn khi sửa" & vbCrlf)
        sSQL &= "Exec D02P1032 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(mAssetID) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= SQLNumber(_isTools) & COMMA 'IsTools, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function



    '    Private Function SQLGetInfoAddNew() As String
    '        Dim sSQL As String = ""
    '        sSQL = "Select AssetAuto,"
    '        sSQL &= "AssetS1Enabled, AssetS1Default,"
    '        sSQL &= "AssetS2Enabled, AssetS2Default,"
    '        sSQL &= "AssetS3Enabled, AssetS3Default,"
    '        sSQL &= "AssetOutputOrder, AssetOutputLength,"
    '        sSQL &= "AssetSeperated, AssetSeperator "
    '        sSQL &= "From D02T0000 WITH(NOLOCK)"
    '        Return sSQL
    '    End Function

    Private Sub LoadTdbcAsset1()
        ' 2/12/2013 id 61719 
        '   Dim dt1 As DataTable = ReturnTableEmployeeID()

        Dim sSQL As String = ""
        'Load tdbcAssetS1ID
        If gsLanguage = "84" Then
            sSQL = "Select     '<<' as AssetS1ID, " & IIf(gbUnicode, "N'<Thêm mới", "'<Theâm môùi").ToString() & ">' As AssetS1Name " & vbCrLf
            sSQL &= "Union" & vbCrLf
            sSQL &= "Select     AssetS1ID, AssetS1Name" & sUnicode & " As AssetS1Name" & vbCrLf
            sSQL &= "From       D02T1000 WITH(NOLOCK)" & vbCrLf
            sSQL &= "Where      Disabled = 0" & vbCrLf
            sSQL &= "Order by   AssetS1ID" & vbCrLf
        Else
            sSQL = "Select     '<<' as AssetS1ID, 'Add' as AssetS1Name" & vbCrLf
            sSQL &= "Union" & vbCrLf
            sSQL &= "Select     AssetS1ID, AssetS1Name" & sUnicode & " As AssetS1Name" & vbCrLf
            sSQL &= "From       D02T1000 WITH(NOLOCK) " & vbCrLf
            sSQL &= "Where      Disabled = 0" & vbCrLf
            sSQL &= "Order by   AssetS1ID" & vbCrLf
        End If
        LoadDataSource(tdbcAssetS1ID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTdbcAsset2()
        Dim sSQL As String = ""
        If gsLanguage = "84" Then
            sSQL = "Select '<<' as AssetS2ID, " & IIf(gbUnicode, "N'<Thêm mới", "'<Theâm môùi").ToString() & ">' as AssetS2Name Union Select AssetS2ID, AssetS2Name" & sUnicode & " As AssetS2Name From D02T2000 WITH(NOLOCK) Where Disabled=0 Order by AssetS2ID"
        Else
            sSQL = "Select '<<' as AssetS2ID, 'Add' as AssetS2Name Union Select AssetS2ID, AssetS2Name" & sUnicode & " As AssetS2Name From D02T2000 WITH(NOLOCK) Where Disabled=0 Order by AssetS2ID"
        End If
        LoadDataSource(tdbcAssetS2ID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTdbcAsset3()
        Dim sSQL As String = ""
        If gsLanguage = "84" Then
            sSQL = "Select '<<' as AssetS3ID, " & IIf(gbUnicode, "N'<Thêm mới", "'<Theâm môùi").ToString() & ">' as AssetS3Name Union Select AssetS3ID, AssetS3Name" & sUnicode & " As AssetS3Name From D02T3000 WITH(NOLOCK) Where Disabled=0 Order by AssetS3ID"
        Else
            sSQL = "Select '<<' as AssetS3ID, 'Add' as AssetS3Name Union Select AssetS3ID, AssetS3Name" & sUnicode & " As AssetS3Name From D02T3000 WITH(NOLOCK) Where Disabled=0 Order by AssetS3ID"
        End If
        LoadDataSource(tdbcAssetS3ID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcAssetS1ID
        LoadTdbcAsset1()
        'Load tdbcAssetS2ID
        LoadTdbcAsset2()
        'Load tdbcAssetS3ID
        LoadTdbcAsset3()
        'Load tdbcObjectTypeID
        '------------------Làm ID 75729 nên sẵn chuẩn hóa combo loại đối tượng và đối tượng
        Dim dtObjectTypeID As DataTable = ReturnTableObjectTypeID(gbUnicode)
        'LoadDataSource(tdbcObjectTypeID, dtObjectTypeID.Copy, gbUnicode)
        'LoadDataSource(tdbcSupplierOTID, dtObjectTypeID.Copy, gbUnicode)
        LoadDataSource(tdbdObjectTypeID, dtObjectTypeID.Copy, gbUnicode)
        LoadObjectTypeID(tdbcObjectTypeID, dtObjectTypeID, gbUnicode)
        LoadObjectTypeID(tdbcSupplierOTID, dtObjectTypeID.DefaultView.ToTable, gbUnicode)
        LoadObjectTypeID(tdbcObjectTypeID2, dtObjectTypeID.DefaultView.ToTable, gbUnicode)

        LoadObjectTypeID(tdbcObjectTypeID6, dtObjectTypeID.DefaultView.ToTable, gbUnicode)
        LoadObjectTypeID(tdbcManagementObTypeID6, dtObjectTypeID.DefaultView.ToTable, gbUnicode)
        LoadObjectTypeID(tdbcSupplierOTIDID6, dtObjectTypeID.DefaultView.ToTable, gbUnicode)
        '--------------------
        'Load tdbdOjectID
        'sSQL = "Select ObjectID, ObjectName" & sUnicode & " As ObjectName, ObjectTypeID From Object WITH(NOLOCK) Where Disabled = 0  order by ObjectID" ' and ObjectTypeID=" & SQLString(ID)
        'dtObjectID = ReturnDataTable(sSQL)

        Using obj As Lemon3.Data.LoadData.ObjectID = New Lemon3.Data.LoadData.ObjectID
            obj.LoadObjectID(tdbcObjectID, dtObjectID)
            obj.LoadObjectID(tdbcObjectID2, dtObjectID.DefaultView.ToTable)
            obj.LoadObjectID(tdbcSupplierID, dtObjectID.DefaultView.ToTable)

            obj.LoadObjectID(tdbcObjectID6, dtObjectID.DefaultView.ToTable)
            obj.LoadObjectID(tdbcManagementObID6, dtObjectID.DefaultView.ToTable)
            obj.LoadObjectID(tdbcSupplierIDID6, dtObjectID.DefaultView.ToTable)
        End Using

        'Load tdbcEmployeeID
        'LoadCboCreateBy(tdbcEmployeeID, gbUnicode)
        'LoadCboCreateBy(tdbcReceiverID, gbUnicode)
        Dim dtCreateBy As DataTable = ReturnTableCreateBy(gbUnicode)
        LoadDataSource(tdbcEmployeeID, dtCreateBy, gbUnicode)
        LoadDataSource(tdbcReceiverID, dtCreateBy.DefaultView.ToTable, gbUnicode)

        'Load tdbcAssetAccountID
        '12/6/2019, Lường Thị Huyền:id 120542-Tài khoản check KSD nhưng khi tạo mã TSCĐ vẫn hiển thị
        sSQL = "Select AccountID,  " & IIf(geLanguage = EnumLanguage.Vietnamese, "AccountName", "AccountName01").ToString & sUnicode & " As AccountName  From D90T0001 WITH(NOLOCK) Where GroupID='7' and OffAccount=0 and AccountStatus=0  and Disabled = 0 Order by AccountID"
        LoadDataSource(tdbcAssetAccountID, sSQL, gbUnicode)
        'Load tdbcDepAccountID
        '13/6/2019, Lường Thị Huyền:id 120542-Tài khoản check KSD nhưng khi tạo mã TSCĐ vẫn hiển thị
        sSQL = "Select AccountID,  " & IIf(geLanguage = EnumLanguage.Vietnamese, "AccountName", "AccountName01").ToString & sUnicode & " As AccountName From D90T0001 WITH(NOLOCK) Where GroupID='19' and OffAccount=0 and AccountStatus=0  and Disabled = 0  Order by AccountID"
        LoadDataSource(tdbcDepAccountID, sSQL, gbUnicode)
        'Load tdbcMethodID
        ' update 31/7/2013id 58504
        sSQL = "Select convert(varchar(20),IntCode) as MethodID, Description" & sUnicode & " As MethodName From D02T8000 WITH(NOLOCK) Where Type=0 and ModuleID='02' and Language=" & SQLString(gsLanguage) & " Order by IntCode"
        LoadDataSource(tdbcMethodID, sSQL, gbUnicode)
        'Load tdbcMethodEndID
        ' update 31/7/2013id 58504
        sSQL = "Select convert(varchar(20),IntCode) as MethodEndID, Description" & sUnicode & " As MethodEndName From D02T8000 WITH(NOLOCK) Where Type=1 and ModuleID='02' and Language=" & SQLString(gsLanguage) & " Order by IntCode"
        LoadDataSource(tdbcMethodEndID, sSQL, gbUnicode)
        'Load tdbcDeprTableID
        sSQL = "Select DeprTableID, DeprTableName" & sUnicode & " As DeprTableName From D02T0070 WITH(NOLOCK) Where Disabled=0 And DivisionID=" & SQLString(gsDivisionID)
        LoadDataSource(tdbcDeprTableID, sSQL, gbUnicode)
        'Load tdbcAssignmentTypeID
        ' update 31/7/2013id 58504
        sSQL = "Select AssignmentTypeID, AssignmentTypeName" & IIf(geLanguage = EnumLanguage.English, "01", "").ToString & sUnicode & " As AssignmentTypeName  From D02V0041"
        LoadDataSource(tdbcAssignmentTypeID, sSQL, gbUnicode)

        'Load tdbcAcode01ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A01' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode01ID, sSQL, gbUnicode)
        ''Load tdbcAcode02ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A02' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode02ID, sSQL, gbUnicode)
        ''Load tdbcAcode03ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A03' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode03ID, sSQL, gbUnicode)
        ''Load tdbcAcode04ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A04' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode04ID, sSQL, gbUnicode)
        ''Load tdbcAcode05ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A05' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode05ID, sSQL, gbUnicode)
        ''Load tdbcAcode06ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A06' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode06ID, sSQL, gbUnicode)
        ''Load tdbcAcode07ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A07' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode07ID, sSQL, gbUnicode)
        ''Load tdbcAcode08ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A08' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode08ID, sSQL, gbUnicode)
        ''Load tdbcAcode09ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A09' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode09ID, sSQL, gbUnicode)
        ''Load tdbcAcode10ID
        'sSQL = "Select ACodeID, Description" & sUnicode & " As Description From D02T0041 WITH(NOLOCK) Where TypeCodeID='A10' and Disabled=0 Order by AcodeID"
        'LoadDataSource(tdbcAcode10ID, sSQL, gbUnicode)

        LoadTDBCACodeID() '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định


        'Load Location
        sSQL = "-- Combo Vi tri " & vbCrLf
        sSQL &= " SELECT		 LookupID As LocationID, Description" & UnicodeJoin(gbUnicode) & " As LocationName"
        sSQL &= " FROM 		D91T0320 WITH(NOLOCK) "
        sSQL &= " WHERE 		LookupType = 'D02_Position' "
        sSQL &= " And (DAGroupID =  ''  Or DAGroupID "
        sSQL &= " IN (Select DAGroupID From lemonsys.dbo.D00V0080 Where UserID= " & SQLString(gsUserID) & " ) Or 'LEMONADMIN' = " & SQLString(gsUserID) & ")"
        sSQL &= " Order By		 LookupID"
        Dim dtLocation As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(tdbcLocationID, dtLocation, gbUnicode)
        LoadDataSource(tdbcLocationIDID6, dtLocation.DefaultView.ToTable, gbUnicode)

        'Incident 68268
        sSQL = "--Combo Phương pháp tạo mã tự động " & vbCrLf
        sSQL &= "SELECT  IGEMethodID, IGEMethodName" & UnicodeJoin(gbUnicode) & " As IGEMethodName , Defaults, FormID "
        sSQL &= "FROM D91T0045 "
        sSQL &= "WHERE 	ModuleID = " & SQLString("02") & "  And Disabled = 0 And FormID = 'D02F0070'  And (DivisionID = " & SQLString(gsDivisionID) & "  Or DivisionID = '' ) "
        sSQL &= "ORDER BY 	IGEMethodID"
        Dim dtTemp As DataTable = ReturnDataTable(sSQL)
        Dim dr() As DataRow = dtTemp.Select("Defaults = 1")
        LoadDataSource(tdbcIGEMethodID, dtTemp, gbUnicode)
        If iAssetAuto = 2 And dr.Length > 0 Then
            sDefaultIGEMethodID = dr(0).Item("IGEMethodID").ToString
            tdbcIGEMethodID.SelectedValue = sDefaultIGEMethodID
        End If

        'Load đơn vị tính
        sSQL = "-- Don vi tinh" & vbCrLf
        sSQL &= "SELECT  UnitID, UnitName" & UnicodeJoin(gbUnicode) & " AS UnitName "
        sSQL &= "FROM 	D07T0005 WITH(NOLOCK) "
        sSQL &= "WHERE Disabled = 0 "
        sSQL &= "ORDER BY UnitID"
        LoadDataSource(tdbcUnitID, sSQL, gbUnicode)

        'Load tài khoản tồn kho
        sSQL = "-- TK ton kho" & vbCrLf
        sSQL &= "SELECT  AccountID, AccountName" & UnicodeJoin(gbUnicode) & " AS  AccountName, "
        sSQL &= "OffAccount, GroupID "
        sSQL &= "FROM 		D90T0001 WITH(NOLOCK)  "
        sSQL &= "WHERE AccountStatus = 0 And Disabled = 0 "
        sSQL &= "ORDER BY 	AccountID "
        LoadDataSource(tdbcAccountID, sSQL, gbUnicode)

        'Load phương pháp tính giá
        sSQL = "-- Combo Phương pháp tính giá" & vbCrLf
        sSQL &= "SELECT	MethodID,MethodName, IsCostByLot, IsPricebyCQuantity "
        sSQL &= "FROM	(SELECT * FROM D07N4000 (" & SQLNumber(gbUnicode) & ", '84')) AS A "
        sSQL &= "ORDER BY	 MethodID"
        LoadDataSource(tdbcMethodIDCCDC, sSQL, gbUnicode)

        sSQL = "--do nguon combo tinh trang" & vbCrLf
        sSQL &= " SELECT		LookupID AS AssetConditionID, "
        sSQL &= " Description" & UnicodeJoin(gbUnicode) & " AS Description"
        sSQL &= " FROM		D91T0320 WITH(NOLOCK)"
        sSQL &= " WHERE		Disabled =0 AND LookupType = 'D02_AssetConditionID'"
        sSQL &= " AND  	(DAGroupID =  ''  Or DAGroupID  IN "
        sSQL &= " (SELECT 	DAGroupID "
        sSQL &= " FROM lemonsys.dbo.D00V0080 "
        sSQL &= " WHERE    UserID = " & SQLString(gsUserID) & ") Or 'LEMONADMIN' = " & SQLString(gsUserID) & ")"
        sSQL &= " ORDER  BY	LookupID"

        LoadDataSource(tdbcAssetConditionName, sSQL, gbUnicode)

        'Load Bộ phận chi phí 'ID : 252774
        sSQL = "-- Combo Bộ phận chi phí" & vbCrLf
        sSQL &= "SELECT	ID AS ChargeObjType,  "
        sSQL &= "CASE WHEN " & SQLString(gsLanguage) & " = '84' THEN Name84U" & vbCrLf
        sSQL &= "ELSE Name01U END AS ChargeObjTypeName" & vbCrLf
        sSQL &= "FROM		D02N5555 ('D02F1031', 'ChargeObjType', '', '', '', '')"
        LoadDataSource(tdbcChargeObjType, sSQL, gbUnicode)
    End Sub

    Dim iAssetAuto As Integer
    Private Sub VisibleIGEMethodID()
        '        Dim sSQL As String = "Select * From D02T0000"
        '        Dim dt As DataTable = ReturnDataTable(sSQL)
        '        If dt.Rows.Count > 0 Then
        iAssetAuto = D02Systems.AssetAuto ' L3Int(dt.Rows(0).Item("AssetAuto"))
        If iAssetAuto = 0 Then
            tdbcAssetS1ID.Visible = True
            tdbcAssetS2ID.Visible = True
            tdbcAssetS3ID.Visible = True
            tdbcAssetS1ID.Enabled = False
            tdbcAssetS2ID.Enabled = False
            tdbcAssetS3ID.Enabled = False
            UnReadOnlyControl(txtAssetID, True)
            tdbcIGEMethodID.Visible = False
        ElseIf iAssetAuto = 1 Then
            tdbcAssetS1ID.Visible = True
            tdbcAssetS2ID.Visible = True
            tdbcAssetS3ID.Visible = True
            tdbcAssetS1ID.Enabled = True
            tdbcAssetS2ID.Enabled = True
            tdbcAssetS3ID.Enabled = True
            ReadOnlyControl(txtAssetID)
            tdbcIGEMethodID.Visible = False

        Else
            tdbcAssetS1ID.Visible = False
            tdbcAssetS2ID.Visible = False
            tdbcAssetS3ID.Visible = False
            tdbcIGEMethodID.Visible = True

        End If
        '  End If


    End Sub


    Private Sub LoadTDBGrid()
        Dim sSQL As String = ""
        sSQL = "Select OrderNum, EquipmentID, EquipmentName" & sUnicode & " As EquipmentName, " & vbCrLf
        sSQL &= "EquipmentQuantity, UnitPrice, EquipmentValue, TaxAmount, AcceptanceTime, PurchaseDate, ObjectTypeID, ObjectID, Notes" & sUnicode & " As Notes " & vbCrLf
        sSQL &= "From D02T4001 WITH(NOLOCK) "
        sSQL &= "Where AssetID = " & SQLString(mAssetID) & " and DivisionID=" & SQLString(gsDivisionID) & " And IsTool =" & SQLNumber(chkIsTools.Checked) & vbCrLf
        sSQL &= "Order by OrderNum"
        LoadDataSource(tdbgDetail, sSQL, gbUnicode)
    End Sub

    Private Sub tdbgDetail_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbgDetail.Columns(COL_EquipmentQuantity).DataField, DxxFormat.D07_QuantityDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgDetail.Columns(COL_UnitPrice).DataField, DxxFormat.D07_UnitCostDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgDetail.Columns(COL_EquipmentValue).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgDetail.Columns(COL_TaxAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbgDetail, arr)
    End Sub


#Region "Events tdbcAssetS1ID"

    'Private Sub tdbcAssetS1ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS1ID.Close
    '    If tdbcAssetS1ID.FindStringExact(tdbcAssetS1ID.Text) = -1 Then
    '        tdbcAssetS1ID.Text = ""
    '        Exit Sub
    '    End If

    '    If tdbcAssetS1ID.Text = "<<" Then
    '        'Dim exe As New D02E1240(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString, gsDivisionID, giTranMonth, giTranYear)
    '        'exe.FormActive = D02E1240Form.D02F3001
    '        'exe.Key01ID = "0"
    '        'exe.Run()
    '        Dim AssetIDNew As String = ""
    '        If CalD02F3001(0, AssetIDNew) Then
    '            LoadTdbcAsset1()
    '            tdbcAssetS1ID.SelectedValue = AssetIDNew ' "<<"
    '        End If

    '    End If
    'End Sub

    Private Sub tdbcAssetS1ID_Validated(sender As Object, e As EventArgs) Handles tdbcAssetS1ID.Validated
        clsFilterCombo.FilterCombo(tdbcAssetS1ID, e)
        If tdbcAssetS1ID.FindStringExact(tdbcAssetS1ID.Text) = -1 Then
            tdbcAssetS1ID.Text = ""
            Exit Sub
        End If
        If tdbcAssetS1ID.Text = "<<" Then
            Dim AssetIDNew As String = ""
            If CalD02F3001(0, AssetIDNew) Then
                LoadTdbcAsset1()
                tdbcAssetS1ID.SelectedValue = AssetIDNew ' "<<"
            End If

        End If
    End Sub

    Private Sub tdbcAssetS1ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetS1ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAssetS1ID.Text = ""
    End Sub

#End Region

    Private Function CalD02F3001(ByVal IndexTab As Integer, ByRef AssetIDNew As String) As Boolean '
        Dim bSaved As Boolean = False
        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "IndexTab", IndexTab)
        Dim frm As Form = CallFormShowDialog("D02D1240", "D02F3001", arrPro)
        bSaved = L3Bool(GetProperties(frm, "SavedOk"))
        AssetIDNew = L3String(GetProperties(frm, "AssetID"))
        Return bSaved
    End Function

#Region "Events tdbcAssetS2ID"

    'Private Sub tdbcAssetS2ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS2ID.Close
    '    If tdbcAssetS2ID.FindStringExact(tdbcAssetS2ID.Text) = -1 Then
    '        tdbcAssetS2ID.Text = ""
    '    End If
    '    If tdbcAssetS2ID.Text = "<<" Then
    '        'Dim exe As New D02E1240(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString, gsDivisionID, giTranMonth, giTranYear)
    '        'exe.FormActive = D02E1240Form.D02F3001
    '        'exe.Key01ID = "1"
    '        'exe.Run()
    '        Dim AssetIDNew As String = ""
    '        If CalD02F3001(1, AssetIDNew) Then
    '            LoadTdbcAsset2()
    '            tdbcAssetS2ID.SelectedValue = AssetIDNew
    '        End If

    '    End If
    'End Sub

    Private Sub tdbcAssetS2ID_Validated(sender As Object, e As EventArgs) Handles tdbcAssetS2ID.Validated
        clsFilterCombo.FilterCombo(tdbcAssetS2ID, e)
        If tdbcAssetS2ID.FindStringExact(tdbcAssetS2ID.Text) = -1 Then
            tdbcAssetS2ID.Text = ""
        End If
        If tdbcAssetS2ID.Text = "<<" Then
            Dim AssetIDNew As String = ""
            If CalD02F3001(1, AssetIDNew) Then
                LoadTdbcAsset2()
                tdbcAssetS2ID.SelectedValue = AssetIDNew
            End If

        End If
    End Sub


    Private Sub tdbcAssetS2ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetS2ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAssetS2ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAssetS3ID"

    'Private Sub tdbcAssetS3ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS3ID.Close
    '    If tdbcAssetS3ID.FindStringExact(tdbcAssetS3ID.Text) = -1 Then tdbcAssetS3ID.Text = ""
    '    If tdbcAssetS3ID.Text = "<<" Then
    '        'Dim exe As New D02E1240(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString, gsDivisionID, giTranMonth, giTranYear)
    '        'exe.FormActive = D02E1240Form.D02F3001
    '        'exe.Key01ID = "2"
    '        'exe.Run()
    '        Dim AssetIDNew As String = ""
    '        If CalD02F3001(2, AssetIDNew) Then
    '            LoadTdbcAsset3()
    '            tdbcAssetS3ID.SelectedValue = AssetIDNew
    '        End If

    '    End If
    'End Sub

    Private Sub tdbcAssetS3ID_Validated(sender As Object, e As EventArgs) Handles tdbcAssetS3ID.Validated
        clsFilterCombo.FilterCombo(tdbcAssetS3ID, e)
        If tdbcAssetS3ID.FindStringExact(tdbcAssetS3ID.Text) = -1 Then tdbcAssetS3ID.Text = ""
        If tdbcAssetS3ID.Text = "<<" Then
            Dim AssetIDNew As String = ""
            If CalD02F3001(2, AssetIDNew) Then
                LoadTdbcAsset3()
                tdbcAssetS3ID.SelectedValue = AssetIDNew
            End If

        End If
    End Sub

    Private Sub tdbcAssetS3ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetS3ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAssetS3ID.Text = ""
    End Sub

#End Region



    Private Function HotKeyF2(ByVal InList As String, Optional ByVal sWhere As String = "") As String
        'Dim f As New D91F6010
        'f.FormPermision = "D91F6010"
        'f.InListID = InList '"2"
        'f.InWhere = sWhere '" ObjectTypeID = " & SQLString(tdbcObjectTypeID.Text)
        'f.WhereValue = ""
        'f.ShowDialog()
        'Dim sKeyID As String = f.OutPut01
        'f.Dispose()
        'Return sKeyID

        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "InListID", InList)
        SetProperties(arrPro, "InWhere", sWhere)
        Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6010", arrPro)
        Dim sKey As String = GetProperties(frm, "Output01").ToString
        Return sKey
    End Function



#Region "Events tdbcEmployeeID with txtEmployeeName"

    Private Sub tdbcEmployeeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcEmployeeID.Close
        If tdbcEmployeeID.FindStringExact(tdbcEmployeeID.Text) = -1 Then
            tdbcEmployeeID.Text = ""
            txtEmployeeName.Text = ""
        End If
    End Sub

    Private Sub tdbcEmployeeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcEmployeeID.SelectedValueChanged
        txtEmployeeName.Text = tdbcEmployeeID.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcEmployeeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcEmployeeID.KeyDown
        If clsFilterCombo.IsNewFilter Then
            Exit Sub ' TH filter dạng mới thì F2 gọi D99F5555 đã có sẵn
        End If
        'Dim sKeyID As String
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcEmployeeID.Text = ""
            txtEmployeeName.Text = ""
        End If
        If e.KeyCode = Keys.F2 Then
            Dim arrPro() As StructureProperties = Nothing
            SetProperties(arrPro, "InListID", "2")
            SetProperties(arrPro, "InWhere", " ObjectTypeID ='NV' ")
            Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6010", arrPro)
            Dim sKey As String = GetProperties(frm, "Output01").ToString
            If sKey <> "" Then
                tdbcEmployeeID.SelectedValue = sKey
                tdbcEmployeeID.Focus()
            End If
        End If
    End Sub

#End Region

#Region "Events tdbcMethodID with txtMethodName"

    Private Sub tdbcMethodID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcMethodID.Close
        If tdbcMethodID.FindStringExact(tdbcMethodID.Text) = -1 Then
            tdbcMethodID.Text = ""
            txtMethodName.Text = ""
        End If
    End Sub

    Private Sub tdbcMethodID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcMethodID.SelectedValueChanged
        txtMethodName.Text = tdbcMethodID.Columns(1).Value.ToString
        If tdbcMethodID.Text.Trim = "0" OrElse tdbcMethodID.Text.Trim = "2" Then 'ID 126368 23/12/2019 Lê Thị Thu Thảo Thêm trường hợp <> 2
            tdbcDeprTableID.SelectedValue = "-1"
            txtDeprTableName.Text = ""
            tdbcDeprTableID.Enabled = False
        Else
            tdbcDeprTableID.Enabled = True
            tdbcDeprTableID.SelectedIndex = 0
        End If
    End Sub

    Private Sub tdbcMethodID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcMethodID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcMethodID.Text = ""
            txtMethodName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcMethodEndID with txtMethodEndName"

    Private Sub tdbcMethodEndID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcMethodEndID.Close
        If tdbcMethodEndID.FindStringExact(tdbcMethodEndID.Text) = -1 Then
            tdbcMethodEndID.Text = ""
            txtMethodEndName.Text = ""
        End If
    End Sub

    Private Sub tdbcMethodEndID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcMethodEndID.SelectedValueChanged
        txtMethodEndName.Text = tdbcMethodEndID.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcMethodEndID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcMethodEndID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcMethodEndID.Text = ""
            txtMethodEndName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcAcode01ID"

    Private Sub tdbcAcode01ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode01ID.Close
        'If tdbcAcode01ID.FindStringExact(tdbcAcode01ID.Text) = -1 Then tdbcAcode01ID.Text = ""
    End Sub

    Private Sub tdbcAcode01ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode01ID.KeyDown
        If e.Alt = True Then
            tdbcAcode01ID.AutoDropDown = False
        Else
            tdbcAcode01ID.AutoDropDown = True
        End If
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode01ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode02ID"

    Private Sub tdbcAcode02ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode02ID.Close
        'If tdbcAcode02ID.FindStringExact(tdbcAcode02ID.Text) = -1 Then tdbcAcode02ID.Text = ""
    End Sub

    Private Sub tdbcAcode02ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode02ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode02ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode03ID"

    Private Sub tdbcAcode03ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode03ID.Close
        'If tdbcAcode03ID.FindStringExact(tdbcAcode03ID.Text) = -1 Then tdbcAcode03ID.Text = ""
    End Sub

    Private Sub tdbcAcode03ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode03ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode03ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode04ID"

    Private Sub tdbcAcode04ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode04ID.Close
        'If tdbcAcode04ID.FindStringExact(tdbcAcode04ID.Text) = -1 Then tdbcAcode04ID.Text = ""
    End Sub

    Private Sub tdbcAcode04ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode04ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode04ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode05ID"

    Private Sub tdbcAcode05ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode05ID.Close
        'If tdbcAcode05ID.FindStringExact(tdbcAcode05ID.Text) = -1 Then tdbcAcode05ID.Text = ""
    End Sub

    Private Sub tdbcAcode05ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode05ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode05ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode06ID"

    Private Sub tdbcAcode06ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode06ID.Close
        'If tdbcAcode06ID.FindStringExact(tdbcAcode06ID.Text) = -1 Then tdbcAcode06ID.Text = ""
    End Sub

    Private Sub tdbcAcode06ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode06ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode06ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode07ID"

    Private Sub tdbcAcode07ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode07ID.Close
        'If tdbcAcode07ID.FindStringExact(tdbcAcode07ID.Text) = -1 Then tdbcAcode07ID.Text = ""
    End Sub

    Private Sub tdbcAcode07ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode07ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode07ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode08ID"

    Private Sub tdbcAcode08ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode08ID.Close
        'If tdbcAcode08ID.FindStringExact(tdbcAcode08ID.Text) = -1 Then tdbcAcode08ID.Text = ""
    End Sub

    Private Sub tdbcAcode08ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode08ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode08ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode09ID"

    Private Sub tdbcAcode09ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode09ID.Close
        'If tdbcAcode09ID.FindStringExact(tdbcAcode09ID.Text) = -1 Then tdbcAcode09ID.Text = ""
    End Sub

    Private Sub tdbcAcode09ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode09ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode09ID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAcode10ID"

    Private Sub tdbcAcode10ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAcode10ID.Close
        'If tdbcAcode10ID.FindStringExact(tdbcAcode10ID.Text) = -1 Then tdbcAcode10ID.Text = ""
    End Sub

    Private Sub tdbcAcode10ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAcode10ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAcode10ID.Text = ""
    End Sub

#End Region


    Private Function SQLGetCaptionFrameIndex() As String
        Dim sSQL As String = ""
        sSQL = "Select FieldName,"
        sSQL &= IIf(gsLanguage = "01", "EngCaption", "VieCaption").ToString() & UnicodeJoin(gbUnicode) '9/8/2019, id 122558-Lỗi font chữ tại D02
        sSQL &= " as Caption, Disabled "
        sSQL &= "From D02T0039 WITH(NOLOCK)"
        Return sSQL
    End Function

    Private Sub AssignCaptionFrameIndex()
        Dim sSQL As String = SQLGetCaptionFrameIndex()
        Dim dt As DataTable = ReturnDataTable(sSQL)

        '9/8/2019, id 122558-Lỗi font chữ tại D02
        lblIndex1.Font = FontUnicode(gbUnicode, lblIndex1.Font.Style)
        lblIndex2.Font = FontUnicode(gbUnicode, lblIndex2.Font.Style)
        lblIndex3.Font = FontUnicode(gbUnicode, lblIndex3.Font.Style)
        lblIndex4.Font = FontUnicode(gbUnicode, lblIndex4.Font.Style)
        lblIndex5.Font = FontUnicode(gbUnicode, lblIndex5.Font.Style)
        lblIndex6.Font = FontUnicode(gbUnicode, lblIndex6.Font.Style)

        lblIndex1.Text = dt.Rows(0)("Caption").ToString()
        txtIndex1.Enabled = Not CType(dt.Rows(0)("Disabled").ToString(), Boolean)
        lblIndex2.Text = dt.Rows(1)("Caption").ToString()
        txtIndex2.Enabled = Not CType(dt.Rows(1)("Disabled").ToString(), Boolean)
        lblIndex3.Text = dt.Rows(2)("Caption").ToString()
        txtIndex3.Enabled = Not CType(dt.Rows(2)("Disabled").ToString(), Boolean)
        lblIndex4.Text = dt.Rows(3)("Caption").ToString()
        txtIndex4.Enabled = Not CType(dt.Rows(3)("Disabled").ToString(), Boolean)
        lblIndex5.Text = dt.Rows(4)("Caption").ToString()
        txtIndex5.Enabled = Not CType(dt.Rows(4)("Disabled").ToString(), Boolean)
        lblIndex6.Text = dt.Rows(5)("Caption").ToString()
        txtIndex6.Enabled = Not CType(dt.Rows(5)("Disabled").ToString(), Boolean)
    End Sub

    Private Function SQLGetCaptionInTab4() As String
        Dim sSQL As String = ""
        '9/8/2019, id 122558-Lỗi font chữ tại D02
        sSQL = "Select TypeCodeID, Disabled, " & IIf(geLanguage = EnumLanguage.Vietnamese, "VieTypeCodeName", "EngTypeCodeName").ToString & UnicodeJoin(gbUnicode) & " as AcodeCaption" & vbCrLf
        sSQL &= "From D02T0040 WITH(NOLOCK) "
        sSQL &= "Where Type='A'"
        sSQL &= "Order by TypeCodeID"
        Return sSQL
    End Function

    Private Sub AssignCaptionInTab4()
        Dim sSQL As String = SQLGetCaptionInTab4()
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            '9/8/2019, id 122558-Lỗi font chữ tại D02
            lblAcode01ID.Font = FontUnicode(gbUnicode, lblAcode01ID.Font.Style)
            lblAcode02ID.Font = FontUnicode(gbUnicode, lblAcode02ID.Font.Style)
            lblAcode03ID.Font = FontUnicode(gbUnicode, lblAcode03ID.Font.Style)
            lblAcode04ID.Font = FontUnicode(gbUnicode, lblAcode04ID.Font.Style)
            lblAcode05ID.Font = FontUnicode(gbUnicode, lblAcode05ID.Font.Style)
            lblAcode06ID.Font = FontUnicode(gbUnicode, lblAcode06ID.Font.Style)
            lblAcode07ID.Font = FontUnicode(gbUnicode, lblAcode07ID.Font.Style)
            lblAcode08ID.Font = FontUnicode(gbUnicode, lblAcode08ID.Font.Style)
            lblAcode09ID.Font = FontUnicode(gbUnicode, lblAcode09ID.Font.Style)
            lblAcode10ID.Font = FontUnicode(gbUnicode, lblAcode10ID.Font.Style)

            lblAcode01ID.Text = dt.Rows(0)("AcodeCaption").ToString()
            lblAcode02ID.Text = dt.Rows(1)("AcodeCaption").ToString()
            lblAcode03ID.Text = dt.Rows(2)("AcodeCaption").ToString()
            lblAcode04ID.Text = dt.Rows(3)("AcodeCaption").ToString()
            lblAcode05ID.Text = dt.Rows(4)("AcodeCaption").ToString()
            lblAcode06ID.Text = dt.Rows(5)("AcodeCaption").ToString()
            lblAcode07ID.Text = dt.Rows(6)("AcodeCaption").ToString()
            lblAcode08ID.Text = dt.Rows(7)("AcodeCaption").ToString()
            lblAcode09ID.Text = dt.Rows(8)("AcodeCaption").ToString()
            lblAcode10ID.Text = dt.Rows(9)("AcodeCaption").ToString()
        End If
    End Sub

    Private Sub Reload()
        MakeSameLocation()
        tdbgDetail_LockedColumns()
        'LoadTDBCombo()
        LoadTDBGrid()
        AssignCaptionInTab4()
        AssignCaptionFrameIndex()
    End Sub

    Private Sub MakeSameLocation()
        grpDetail.Left = lblTool.Left
        grpDetail.Top = txtTool.Top - 5
    End Sub

    Private Sub btnDetail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDetail.Click
        grpDetail.Visible = True
        grpIndex.Visible = False
    End Sub

    Private Sub btnCloseDetail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCloseDetail.Click
        grpDetail.Visible = False
        grpIndex.Visible = True
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FormatNumber()
        InputNumber(cneOQuantity, SqlDbType.Float, DxxFormat.D07_QuantityDecimals, , 28, 8)
        If _FormState = EnumFormState.FormAdd Then Exit Sub
        If txtConvertedAmount.Text <> "" Then txtConvertedAmount.Text = SQLNumber(txtConvertedAmount.Text, DxxFormat.D90_ConvertedDecimals)
        If txtRemainAmount.Text <> "" Then txtRemainAmount.Text = SQLNumber(txtRemainAmount.Text, DxxFormat.D90_ConvertedDecimals)
        If txtDepreciationAmount.Text <> "" Then txtDepreciationAmount.Text = SQLNumber(txtDepreciationAmount.Text, DxxFormat.D90_ConvertedDecimals)
        If txtAmountDepreciation.Text <> "" Then txtAmountDepreciation.Text = SQLNumber(txtAmountDepreciation.Text, DxxFormat.D90_ConvertedDecimals)
        If txtPercentage.Text <> "" Then txtPercentage.Text = SQLNumber(txtPercentage.Text, DxxFormat.DefaultNumber2)
        If txtIndex1.Text <> "" Then txtIndex1.Text = SQLNumber(txtIndex1.Text, "N6")
        If txtIndex2.Text <> "" Then txtIndex2.Text = SQLNumber(txtIndex2.Text, "N6")
        If txtIndex3.Text <> "" Then txtIndex3.Text = SQLNumber(txtIndex3.Text, "N6")
        If txtIndex4.Text <> "" Then txtIndex4.Text = SQLNumber(txtIndex4.Text, "N6")
        If txtIndex5.Text <> "" Then txtIndex5.Text = SQLNumber(txtIndex5.Text, "N6")
        If txtIndex6.Text <> "" Then txtIndex6.Text = SQLNumber(txtIndex6.Text, "N6")

    End Sub

    Private Sub btnConvertedAmount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConvertedAmount.Click
        Dim f As New D02F0602()
        f.VisibleControlsConvertedAmount(True)
        f.VisibleControlsDepreciation(False)
        f.VisibleControlsHistory(False)
        f.Mode = 1
        f.AssetID = txtAssetID.Text
        f.AssetName = txtAssetName.Text
        f.ShowDialog()
        f.Dispose()
    End Sub

    Private Sub btnDepreciate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDepreciate.Click
        Dim f As New D02F0602()
        f.VisibleControlsDepreciation(True)
        f.VisibleControlsConvertedAmount(False)
        f.VisibleControlsHistory(False)
        f.Mode = 2
        f.AssetID = txtAssetID.Text
        f.AssetName = txtAssetName.Text
        f.ShowDialog()
        f.Dispose()
    End Sub

    Private Sub btnHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnHistory.Click
        Dim f As New D02F0602()
        f.VisibleControlsHistory(True)
        f.VisibleControlsConvertedAmount(False)
        f.VisibleControlsDepreciation(False)
        f.Mode = 3
        f.AssetID = txtAssetID.Text
        f.AssetName = txtAssetName.Text
        f.ShowDialog()
        f.Dispose()
    End Sub

    'Private Sub tdbcAssetS1ID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetS1ID.LostFocus
    '    If tdbcAssetS1ID.Text <> "" Or tdbcAssetS2ID.Text <> "" Or tdbcAssetS3ID.Text <> "" Then
    '        If bAuto Then
    '            gnNewLastKey = 0
    '            _S1 = IIf(IsDBNull(tdbcAssetS1ID.Text) Or tdbcAssetS1ID.Text = "<<", "", tdbcAssetS1ID.Text).ToString
    '            D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, txtAssetID, False, _TableName)
    '            If Not gbCheckLastKey Then Exit Sub

    '        End If
    '    End If
    'End Sub

    Private Sub tdbcAssetS1ID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetS1ID.SelectedValueChanged
        If _FormState <> EnumFormState.FormEdit Then
            If bAuto = 1 Then
                gnNewLastKey = 0
                _S1 = IIf(IsDBNull(tdbcAssetS1ID.Text) Or tdbcAssetS1ID.Text = "<<", "", tdbcAssetS1ID.Text).ToString
                D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, txtAssetID, False, _TableName)
            End If
        End If

    End Sub

    'Private Sub tdbcAssetS2ID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetS2ID.LostFocus
    '    If tdbcAssetS1ID.Text <> "" Or tdbcAssetS2ID.Text <> "" Or tdbcAssetS3ID.Text <> "" Then
    '        If bAuto Then
    '            gnNewLastKey = 0
    '            _S2 = IIf(IsDBNull(tdbcAssetS2ID.Text) Or tdbcAssetS2ID.Text = "<<", "", tdbcAssetS2ID.Text).ToString
    '            D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, txtAssetID, False, _TableName)
    '            If Not gbCheckLastKey Then Exit Sub
    '        End If
    '    End If
    'End Sub

    Private Sub tdbcAssetS2ID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetS2ID.SelectedValueChanged
        If _FormState <> EnumFormState.FormEdit Then
            If bAuto = 1 Then
                gnNewLastKey = 0
                _S2 = IIf(tdbcAssetS2ID.Text = "<<", "", tdbcAssetS2ID.Text).ToString
                D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, txtAssetID, False, _TableName)
            End If
        End If

    End Sub

    'Private Sub tdbcAssetS3ID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetS3ID.LostFocus
    '    If tdbcAssetS1ID.Text <> "" Or tdbcAssetS2ID.Text <> "" Or tdbcAssetS3ID.Text <> "" Then
    '        If bAuto Then
    '            gnNewLastKey = 0
    '            _S3 = IIf(IsDBNull(tdbcAssetS3ID.Text) Or tdbcAssetS3ID.Text = "<<", "", tdbcAssetS3ID.Text).ToString
    '            D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, txtAssetID, False, _TableName)
    '            If Not gbCheckLastKey Then Exit Sub
    '        End If
    '    End If
    'End Sub

    Private Sub tdbcAssetS3ID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetS3ID.SelectedValueChanged
        If _FormState <> EnumFormState.FormEdit Then
            If bAuto = 1 Then
                gnNewLastKey = 0
                _S3 = IIf(tdbcAssetS3ID.Text = "<<", "", tdbcAssetS3ID.Text).ToString
                D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, txtAssetID, False, _TableName)
            End If
        End If

    End Sub

    Private Sub tdbgDetail_LockedColumns()
        tdbgDetail.Splits(SPLIT0).DisplayColumns(COL_OrderNum).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    Private Sub tdbgDetail_AfterDelete(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbgDetail.AfterDelete
        'UpdateOrderNum(tdbgDetail, COL_OrderNum)
        UpdateTDBGOrderNum(tdbgDetail, COL_OrderNum)
    End Sub

    Private Sub tdbgDetail_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbgDetail.KeyPress
        Select Case tdbgDetail.Col
            Case COL_EquipmentID
                e.KeyChar = UCase(e.KeyChar) 'Nhập các ký tự hoa
            Case COL_EquipmentQuantity
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_EquipmentValue
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub tdbgDetail_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgDetail.AfterColUpdate
        'Select Case e.ColIndex
        '    Case COL_OrderNum
        '    Case COL_EquipmentID
        '        tdbgDetail.Columns(COL_OrderNum).Text = (Val(tdbgDetail.Bookmark) + 1).ToString
        '    Case COL_EquipmentName
        '        tdbgDetail.Columns(COL_OrderNum).Text = (Val(tdbgDetail.Bookmark) + 1).ToString
        '    Case COL_EquipmentQuantity
        '        tdbgDetail.Columns(COL_OrderNum).Text = (Val(tdbgDetail.Bookmark) + 1).ToString
        '    Case COL_EquipmentValue
        '        tdbgDetail.Columns(COL_OrderNum).Text = (Val(tdbgDetail.Bookmark) + 1).ToString
        '    Case COL_ObjectTypeID
        '        tdbgDetail.Columns(COL_OrderNum).Text = (Val(tdbgDetail.Bookmark) + 1).ToString
        '        LoadtdbdObjectID(tdbgDetail.Columns(COL_ObjectTypeID).Text)
        '    Case COL_ObjectID
        '        tdbgDetail.Columns(COL_OrderNum).Text = (Val(tdbgDetail.Bookmark) + 1).ToString
        '    Case COL_Notes
        '        tdbgDetail.Columns(COL_OrderNum).Text = (Val(tdbgDetail.Bookmark) + 1).ToString
        'End Select

        UpdateTDBGOrderNum(tdbgDetail, COL_OrderNum, e.ColIndex)
    End Sub
    Dim sLastKey As String
    Dim bNewIGE As Boolean = False

    Private Sub btnSetNewKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSetNewKey.Click
        If iAssetAuto = 1 Then
            Dim sAssetIDOld As String = ""
            sAssetIDOld = txtAssetID.Text
            '_OutputLength = txtAssetID.Text.Length
            D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, txtAssetID, True, _TableName)
            If txtAssetID.Text <> sAssetIDOld Then
                bAuto = 0
            End If
        ElseIf iAssetAuto = 2 Then
            Dim f As New D02F5705
            f.NewKeyString = sKeyString
            Dim _dtData As DataTable = CType(tdbcIGEMethodID.DataSource, DataTable)
            If _dtData IsNot Nothing AndAlso _dtData.Rows.Count > 0 Then
                f.IGEMethodID = ReturnValueC1Combo(tdbcIGEMethodID)
            End If
            f.ShowDialog()
            If f.Result = True Then
                If f.ID <> "" Then
                    txtAssetID.Text = f.ID
                    sLastKey = f.LastKey.ToString
                End If
            End If
            f.Dispose()

            Exit Sub
        End If

    End Sub

    Private Function SQLGetAssetID() As String
        Dim sSQL As String = ""
        sSQL = "Select 1 From D02T0001 WITH(NOLOCK) "
        sSQL &= "Where AssetID=" & SQLString(txtAssetID.Text) '& " and DivisionID=" & SQLString(gsDivisionID)
        Return sSQL
    End Function

    Private Function AllowSave() As Boolean
        If Not chkIsTools.Checked Then
            Dim iTypeLength As Integer = 0
            If iAssetAuto <> 2 Then
                If tdbcAssetS1ID.Enabled And tdbcAssetS1ID.Text = "" Then
                    D99C0008.MsgNotYetChoose(rL3("Ma_tai_san") & " 1")
                    tdbcAssetS1ID.Focus()
                    Return False
                End If
                If tdbcAssetS2ID.Enabled And tdbcAssetS2ID.Text = "" Then
                    D99C0008.MsgNotYetChoose(rL3("Ma_tai_san") & " 2")
                    tdbcAssetS2ID.Focus()
                    Return False
                End If
                If tdbcAssetS3ID.Enabled And tdbcAssetS3ID.Text = "" Then
                    D99C0008.MsgNotYetChoose(rL3("Ma_tai_san") & " 3")
                    tdbcAssetS3ID.Focus()
                    Return False
                End If
            End If

            If _FormState <> EnumFormState.FormEdit Then
                If txtAssetID.Text = "" Then
                    D99C0008.MsgNotYetEnter(rL3("Ma_tai_san"))
                    txtAssetID.Focus()
                    Return False
                End If
                If bAuto = 0 Then
                    iTypeLength = TypeLength(4)
                    If iTypeLength <> 0 Then
                        If iTypeLength <> txtAssetID.Text.Trim.Length Then
                            D99C0008.MsgL3(rL3("Chieu_dai_cua_ma_tai_san_phai_bang") & Space(1) & iTypeLength & Space(1) & rL3("_ky_tu"))
                            txtAssetID.Focus()
                            Exit Function
                        End If
                    End If
                End If

            End If

            ' update 14/5/2013 id 56522 - bổ sung kiểm tra trong Module kế thừa
            If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
                Dim sSQL As String = SQLGetAssetID()
                Dim dt As DataTable = ReturnDataTable(sSQL)
                If dt.Rows.Count > 0 Then
                    D99C0008.MsgDuplicatePKey()
                    txtAssetID.Focus()
                    Return False
                End If
            End If

            If txtAssetName.Text = "" Then
                D99C0008.MsgNotYetEnter(rL3("Ten_tai_san"))
                txtAssetName.Focus()
                Return False
            End If

            If tdbcSupplierOTID.Text <> "" And tdbcSupplierID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Nha_cung_cap"))
                tab.SelectedTab = tab.TabPages(0)
                tdbcSupplierID.Focus()
                Return False
            End If
            'If _FormState <> EnumFormState.FormEdit Then
            If tdbcAssetAccountID.Enabled = True And tdbcAssetAccountID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Tai_khoan_tai_san"))
                tab.SelectedTab = tab.TabPages(1)
                tdbcAssetAccountID.Focus()
                Return False
            End If
            If tdbcDepAccountID.Enabled = True And tdbcDepAccountID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Tai_khoan_khau_hao"))
                tab.SelectedTab = tab.TabPages(1)
                tdbcDepAccountID.Focus()
                Return False
            End If

            If tdbcMethodID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Phuong_phap_khau_hao"))
                tab.SelectedTab = tab.TabPages(1)
                tdbcMethodID.Focus()
                Return False
            End If
            If tdbcMethodEndID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Xu_ly_khau_hao_ky_cuoi"))
                tab.SelectedTab = tab.TabPages(1)
                tdbcMethodEndID.Focus()
                Return False
            End If
            If tdbcDeprTableID.Enabled = True And tdbcDeprTableID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Bang_khau_hao"))
                tab.SelectedTab = tab.TabPages(1)
                tdbcDeprTableID.Focus()
                Return False
            End If
            If tdbcAssignmentTypeID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Kieu_phan_bo"))
                tab.SelectedTab = tab.TabPages(1)
                tdbcAssignmentTypeID.Focus()
                Return False
            End If

            If _FormState <> EnumFormState.FormEdit Then
                If c1datePeriod.Text <> "" Then
                    If CDbl(c1datePeriod.Text.Substring(3, 4)) < 1900 OrElse CDbl(c1datePeriod.Text.Substring(3, 4)) > 2100 Then
                        D99C0008.MsgL3(rL3("Ky_su_dung_phai_nam_trong_khoang_tu_1900_den_2100"))
                        tab.SelectedTab = tab.TabPages(1)
                        c1datePeriod.Focus()
                        Return False
                    End If
                End If
                If c1dateDepPeriod.Text <> "" Then
                    If CDbl(c1dateDepPeriod.Text.Substring(3, 4)) < 1900 OrElse CDbl(c1dateDepPeriod.Text.Substring(3, 4)) > 2100 Then
                        D99C0008.MsgL3(rL3("Ky_bat_dau_phai_nam_trong_khoang_tu_1900_den_2100"))
                        tab.SelectedTab = tab.TabPages(1)
                        c1dateDepPeriod.Focus()
                        Return False
                    End If
                End If
                If c1dateTranDate.Text <> "" Then
                    If CDbl(c1dateTranDate.Text.Substring(3, 4)) < 1900 OrElse CDbl(c1dateTranDate.Text.Substring(3, 4)) > 2100 Then
                        D99C0008.MsgL3(rL3("Ky_hinh_thanh_phai_nam_trong_khoang_tu_1900_den_2100"))
                        tab.SelectedTab = tab.TabPages(1)
                        c1dateTranDate.Focus()
                        Return False
                    End If
                End If
                'Thiên Huỳnh Edit 21/06/2010
                If c1dateDepPeriod.Text <> "" And c1datePeriod.Text <> "" Then
                    If CDbl(c1dateDepPeriod.Text.Substring(3, 4)) < CDbl(c1datePeriod.Text.Substring(3, 4)) Then
                        If D99C0008.MsgAsk(rL3("Ky_bat_dau_tinh_khau_hao_phai_lon_hon_hoac_bang_ky_su_dung") & vbCrLf & rL3("MSG000021"), MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            tab.SelectedTab = tab.TabPages(1)
                            c1dateDepPeriod.Focus()
                            Return False
                        End If
                    ElseIf CDbl(c1dateDepPeriod.Text.Substring(3, 4)) = CDbl(c1datePeriod.Text.Substring(3, 4)) Then
                        If CDbl(c1dateDepPeriod.Text.Substring(0, 2)) < CDbl(c1datePeriod.Text.Substring(0, 2)) Then
                            If D99C0008.MsgAsk(rL3("Ky_bat_dau_tinh_khau_hao_phai_lon_hon_hoac_bang_ky_su_dung") & vbCrLf & rL3("MSG000021"), MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                tab.SelectedTab = tab.TabPages(1)
                                c1dateDepPeriod.Focus()
                                Return False
                            End If
                        End If
                    End If
                End If
                If c1dateDepPeriod.Text <> "" And c1dateTranDate.Text <> "" Then
                    If CDbl(c1dateDepPeriod.Text.Substring(3, 4)) < CDbl(c1dateTranDate.Text.Substring(3, 4)) Then
                        If D99C0008.MsgAsk(rL3("Ky_bat_dau_tinh_khau_hao_phai_lon_hon_hoac_bang_ky_hinh_thanh") & vbCrLf & rL3("MSG000021"), MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            tab.SelectedTab = tab.TabPages(1)
                            c1dateDepPeriod.Focus()
                            Return False
                        End If
                    ElseIf CDbl(c1dateDepPeriod.Text.Substring(3, 4)) = CDbl(c1dateTranDate.Text.Substring(3, 4)) Then
                        If CDbl(c1dateDepPeriod.Text.Substring(0, 2)) < CDbl(c1dateTranDate.Text.Substring(0, 2)) Then
                            If D99C0008.MsgAsk(rL3("Ky_bat_dau_tinh_khau_hao_phai_lon_hon_hoac_bang_ky_hinh_thanh") & vbCrLf & rL3("MSG000021"), MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                tab.SelectedTab = tab.TabPages(1)
                                c1dateDepPeriod.Focus()
                                Return False
                            End If
                        End If
                    End If
                End If
                If c1datePeriod.Text <> "" And c1dateTranDate.Text <> "" Then
                    If CDbl(c1datePeriod.Text.Substring(3, 4)) < CDbl(c1dateTranDate.Text.Substring(3, 4)) Then
                        If D99C0008.MsgAsk(rL3("Ky_su_dung_phai_lon_hon_hoac_bang_ky_hinh_thanh") & vbCrLf & rL3("MSG000021"), MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            tab.SelectedTab = tab.TabPages(1)
                            c1datePeriod.Focus()
                            Return False
                        End If
                    ElseIf CDbl(c1datePeriod.Text.Substring(3, 4)) = CDbl(c1dateTranDate.Text.Substring(3, 4)) Then
                        If CDbl(c1datePeriod.Text.Substring(0, 2)) < CDbl(c1dateTranDate.Text.Substring(0, 2)) Then
                            If D99C0008.MsgAsk(rL3("Ky_su_dung_phai_lon_hon_hoac_bang_ky_hinh_thanh") & vbCrLf & rL3("MSG000021"), MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                tab.SelectedTab = tab.TabPages(1)
                                c1datePeriod.Focus()
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
            If D02Systems.IsCalDepByDate = True Then
                If c1dateDepDate.Text = "" Then
                    D99C0008.MsgNotYetEnter(lblDepDate.Text)
                    tab.SelectedTab = tab.TabPages(1)
                    c1dateDepDate.Focus()
                    Return False
                End If
            End If
            If c1dateDepDate.Text <> "" And c1dateDepPeriod.Text <> "" Then
                If CDate(c1dateDepDate.Value).Year <> CDate(c1dateDepPeriod.Value).Year OrElse CDate(c1dateDepDate.Value).Month <> CDate(c1dateDepPeriod.Value).Month Then
                    D99C0008.MsgL3(rL3("Ngay_khau_hao_phai_nam_trong_ky_khau_hao"), L3MessageBoxIcon.Exclamation)
                    tab.SelectedTab = tab.TabPages(1)
                    c1dateDepDate.Focus()
                    Return False
                End If
            End If

        End If


        'Incident 78836 
        'Nếu có nhập một cột trong 3 cột thì bắt buộc nhập còn lại
        For i As Integer = 0 To tdbgDetail.RowCount - 1
            If tdbgDetail(i, COL_EquipmentID).ToString() = "" AndAlso tdbgDetail(i, COL_EquipmentName).ToString() = "" Then
                'Khong kiem tra gi het
            Else
                If tdbgDetail(i, COL_EquipmentID).ToString() = "" Then
                    D99C0008.MsgNotYetEnter(rL3("Ma_thiet_bi"))
                    tab.SelectedTab = tab.TabPages(2)
                    tdbgDetail.Focus()
                    tdbgDetail.SplitIndex = 0
                    tdbgDetail.Col = COL_EquipmentID
                    tdbgDetail.Row = i
                    Return False
                End If
                If tdbgDetail(i, COL_EquipmentName).ToString() = "" Then
                    D99C0008.MsgNotYetEnter(rL3("Ten_thiet_bi"))
                    tab.SelectedTab = tab.TabPages(2)
                    tdbgDetail.Focus()
                    tdbgDetail.SplitIndex = 0
                    tdbgDetail.Col = COL_EquipmentName
                    tdbgDetail.Row = i
                    Return False
                End If

                '7/4/2017, 	Phạm Thị Thu: id 96093-[CDS] Thẻ TSCĐ - Danh mục TSCĐ theo chủng loại
                If tdbgDetail(i, COL_EquipmentID).ToString() <> "" Then
                    If Number(tdbgDetail(i, COL_UnitPrice)) = 0 Then
                        D99C0008.MsgNotYetEnter(rL3("Don_gia"))
                        tab.SelectedTab = tab.TabPages(2)
                        tdbgDetail.Focus()
                        tdbgDetail.SplitIndex = 0
                        tdbgDetail.Col = COL_UnitPrice
                        tdbgDetail.Row = i
                        Return False
                    End If
                End If

                If tdbgDetail(i, COL_ObjectTypeID).ToString() <> "" And tdbgDetail(i, COL_ObjectID).ToString() = "" Then
                    D99C0008.MsgNotYetChoose(rL3("Phong_ban"))
                    tab.SelectedTab = tab.TabPages(2)
                    tdbgDetail.Focus()
                    tdbgDetail.SplitIndex = 0
                    tdbgDetail.Col = COL_ObjectID
                    tdbgDetail.Row = i
                    Return False
                End If
            End If
        Next
        If chkIsTools.Checked Then
            If txtAssetName.Text = "" Then
                D99C0008.MsgNotYetEnter(rL3("Ten_tai_san"))
                txtAssetName.Focus()
                Return False
            End If
            If tdbcUnitID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rL3("Don_vi_tinh"))
                tab.SelectedTab = tab06
                tdbcUnitID.Focus()
                Return False
            End If
            If tdbcAccountID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rL3("TK_ton_kho"))
                tab.SelectedTab = tab06
                tdbcAccountID.Focus()
                Return False
            End If
            If tdbcMethodIDCCDC.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rL3("Phuong_phap_tinh_gia"))
                tab.SelectedTab = tab06
                tdbcMethodIDCCDC.Focus()
                Return False
            End If
        End If
        Return True
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0010
    '# Created User: HUỲNH KHANH
    '# Created Date: 07/09/2015 03:16:56
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0010() As String
        Dim sSQL As String = ""
        sSQL &= ("-- kiem tra truoc khi luu" & vbCrlf)
        sSQL &= "Exec D02P0010 "
        sSQL &= SQLString(txtAssetID.Text) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString(Me.Name) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostName, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, int, NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) 'TranYear, int, NOT NULL
        Return sSQL
    End Function



    Private Sub SaveLastKey1(ByVal strS1 As String, ByVal strS2 As String, ByVal strS3 As String, ByVal sOutputOrder As String, ByVal iOutputLength As Integer, ByVal sSeperator As String, ByVal bFlagSave As Boolean, ByVal sTableName As String)
        Dim iOutputOrder As Integer = -1
        Select Case sOutputOrder
            Case "NSSS"
                iOutputOrder = D99D0041.OutOrderEnum.lmNSSS
            Case "SNSS"
                iOutputOrder = D99D0041.OutOrderEnum.lmSNSS
            Case "SSNS"
                iOutputOrder = D99D0041.OutOrderEnum.lmSSNS
            Case "SSSN"
                iOutputOrder = D99D0041.OutOrderEnum.lmSSSN
        End Select
        CreateIGEVoucherNo(strS1, strS2, strS3, CType(iOutputOrder, D99D0041.OutOrderEnum), iOutputLength, sSeperator, True, sTableName)
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        btnSave.Focus()
        If btnSave.Focused = False Then Exit Sub

        If D99C0008.MsgAskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowSave() Then Exit Sub
        If _FormState <> EnumFormState.FormEdit Then
            If Not CheckStore(SQLStoreD02P0010) Then Exit Sub
        End If

        _savedOK = False
        Dim sSQL As String = ""
        txtTool.Text = GetTool() '
        Select Case _FormState
            Case EnumFormState.FormAdd, EnumFormState.FormCopy
                If bAuto = 1 Then
                    _S1 = IIf(tdbcAssetS1ID.Text = "<<", "", tdbcAssetS1ID.Text).ToString
                    _S2 = IIf(tdbcAssetS2ID.Text = "<<", "", tdbcAssetS2ID.Text).ToString
                    _S3 = IIf(tdbcAssetS3ID.Text = "<<", "", tdbcAssetS3ID.Text).ToString
                    SaveLastKey1(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, True, _TableName)
                End If

                If chkIsTools.Checked Then
                    sSQL &= SQLInsertD02T1001().ToString & vbCrLf
                    sSQL &= SQLStoreD02P1037(0).ToString & vbCrLf
                Else
                    sSQL = SQLInsertD02T0001() & vbCrLf
                End If
                sSQL &= SQLStoreD02P1031() & vbCrLf ' update 18/7/2013 id 57380
                sSQL &= SQLInsertD02T4001s() & vbCrLf

            Case EnumFormState.FormEdit
                If chkIsTools.Checked Then
                    sSQL = SQLUpdateD02T1001().ToString & vbCrLf
                    sSQL &= SQLStoreD02P1037(1).ToString & vbCrLf
                Else
                    sSQL = SQLUpdateD02T0001().ToString & vbCrLf
                End If
                sSQL &= SQLDeleteD02T4001() & vbCrLf
                sSQL &= SQLInsertD02T4001s() & vbCrLf
        End Select
        If iAssetAuto = 2 Then
            If D02Systems.IsShowFormAutoCreate = True And _bFormD02F0087 Then '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
                sSQL &= _sSQLD91T1001_SaveLastKey
            Else
                If tdbcIGEMethodID.Text <> "" Then ' nếu có chọn Mã PP tự động mới thực hiện đoạn lệnh update Lastkey, ngược lại thì không.
                    If sLastKey = "0" Then
                    ElseIf sLastKey = "1" Then
                        sSQL &= SQLInsertD91T1001().ToString
                    Else
                        sSQL &= SQLUpdateD91T1001().ToString
                    End If
                End If
            End If

        End If

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL)
        If bRunSQL Then
            SaveOK()
            _savedOK = True
            SQLInsertD02T0004()
            If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormCopy Then
                'ExecuteAuditLog("Assets", "01", txtAssetID.Text, txtAssetName.Text, "", "", "")
                Lemon3.D91.RunAuditLog("02", "Assets", "01", txtAssetID.Text, txtAssetName.Text)
                btnSave.Enabled = False
                btnNext.Enabled = True
                btnNext.Focus()
            ElseIf _FormState = EnumFormState.FormEdit Then
                'ExecuteAuditLog("Assets", "02", txtAssetID.Text, txtAssetName.Text, "", "", "")
                Lemon3.D91.RunAuditLog("02", "Assets", "02", txtAssetID.Text, txtAssetName.Text)
                btnSave.Enabled = True
                btnClose.Enabled = True
                btnClose.Focus()
            End If
            'LoadImage()
        Else
            SaveNotOK()
            _savedOK = False
            btnSave.Enabled = True
            btnClose.Enabled = True
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T1001
    '# Created User: HUỲNH KHANH
    '# Created Date: 07/10/2014 10:50:58
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T1001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Update KeyString & Lastkey" & vbCrlf)
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
    '# Created User: HUỲNH KHANH
    '# Created Date: 07/10/2014 10:50:58
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD91T1001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Update KeyString & Lastkey" & vbCrlf)
        sSQL.Append("Update D91T1001 Set ")
        sSQL.Append("LastKey = " & SQLNumber(sLastKey)) 'varchar[250], NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("KeyString = " & SQLString(sKeyString)) 'int, NOT NULL
        sSQL.Append(" And ModuleID = " & SQLString("02")) 'varchar[20], NOT NULL
        sSQL.Append(" And FormID = " & SQLString("D02F0070")) 'varchar[20], NOT NULL

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0001
    '# Created User: Thiên Huỳnh
    '# Created Date: 24/06/2010 10:15:48
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0001 Set ")
        sSQL.Append("AssetNameU = " & SQLStringUnicode(txtAssetName.Text, gbUnicode, True) & COMMA) 'varchar[100], NULL
        sSQL.Append("ShortNameU = " & SQLStringUnicode(txtShortName.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("DivisionID = " & SQLString(gsDivisionID) & COMMA) 'varchar[20], NUL
        sSQL.Append("CountryID = " & SQLString(txtCountryID.Text) & COMMA) 'varchar[20], NULL
        If txtMadeYear.Text = "" Then
            sSQL.Append("MadeYear = NULL" & COMMA) 'int, NULL
        Else
            sSQL.Append("MadeYear = " & SQLNumber(txtMadeYear.Text) & COMMA) 'int, NULL
        End If
        sSQL.Append("SeriNo = " & SQLString(txtSeriNo.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("Version = " & SQLString(txtVersion.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("AssetNo = " & SQLString(txtAssetNo.Text) & COMMA) 'varchar[20], NULL
        If IsDBNull(txtAssetTag.Text) Or txtAssetTag.Text = "" Then
            'sSQL.Append("AssetTag = NULL" & COMMA) 'varchar[100], NULL
            'sSQL.Append("AssetTagU = NULL" & COMMA) 'varchar[100], NULL
        Else
            'sSQL.Append("AssetTag = " & SQLStringUnicode(txtAssetTag.Text, gbUnicode, False) & COMMA) 'varchar[100], NULL
            sSQL.Append("AssetTagU = " & SQLStringUnicode(txtAssetTag.Text, gbUnicode, True) & COMMA) 'varchar[100], NULL
        End If

        If c1datePeriod.Text.Length > 4 Then
            sSQL.Append("UseMonth = " & SQLNumber(c1datePeriod.Text.Substring(0, 2)) & COMMA) 'tinyint, NULL
            sSQL.Append("UseYear = " & SQLNumber(c1datePeriod.Text.Substring(3, 4)) & COMMA) 'smallint, NULL
        Else
            sSQL.Append("UseMonth = " & SQLNumber(0) & COMMA) 'tinyint, NULL
            sSQL.Append("UseYear = " & SQLNumber(0) & COMMA) 'smallint, NULL
        End If

        If c1dateDepPeriod.Text.Length > 4 Then
            sSQL.Append("DepMonth = " & SQLNumber(c1dateDepPeriod.Text.Substring(0, 2)) & COMMA) 'tinyint, NULL
            sSQL.Append("DepYear = " & SQLNumber(c1dateDepPeriod.Text.Substring(3, 4)) & COMMA) 'smallint, NULL
        Else
            sSQL.Append("DepMonth = " & SQLNumber(0) & COMMA) 'tinyint, NULL
            sSQL.Append("DepYear = " & SQLNumber(0) & COMMA) 'smallint, NULL
        End If
        If _Completed = False Then
            If tdbcObjectTypeID.Enabled Then
                sSQL.Append("ObjectTypeID = " & SQLString(tdbcObjectTypeID.Text) & COMMA) 'varchar[20], NULL
            End If
            If tdbcObjectID.Enabled Then
                sSQL.Append("ObjectID = " & SQLString(tdbcObjectID.Text) & COMMA) 'varchar[20], NULL
            End If
            If tdbcEmployeeID.Enabled Then
                sSQL.Append("EmployeeID = " & SQLString(tdbcEmployeeID.Text) & COMMA) 'varchar[20], NULL
            End If
            sSQL.Append("AssetAccountID = " & SQLString(tdbcAssetAccountID.Text) & COMMA) 'varchar[20], NULL
            sSQL.Append("DepAccountID = " & SQLString(tdbcDepAccountID.Text) & COMMA) 'varchar[20], NULL

            sSQL.Append("ConvertedAmount = " & SQLMoney(txtConvertedAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
            sSQL.Append("RemainAmount = " & SQLMoney(txtRemainAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
            sSQL.Append("DepreciatedPeriod = " & SQLNumber(txtDepreciatedPeriod.Text) & COMMA) 'int, NULL
            sSQL.Append("DepreciatedAmount = " & SQLMoney(txtDepreciationAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
            sSQL.Append("AmountDepreciation = " & SQLMoney(txtAmountDepreciation.Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
            sSQL.Append("ServiceLife = " & SQLNumber(txtServiceLife.Text) & COMMA) 'int, NULL
            sSQL.Append("Percentage = " & SQLMoney(txtPercentage.Text, DxxFormat.DefaultNumber2) & COMMA) 'money, NULL
        End If

        If Not bIsDepreciated Then
            sSQL.Append("MethodID = " & SQLNumber(tdbcMethodID.Text) & COMMA) 'tinyint, NOT NULL
            sSQL.Append("MethodEndID = " & SQLNumber(tdbcMethodEndID.Text) & COMMA) 'tinyint, NOT NULL
            sSQL.Append("DeprTableID = " & SQLString(tdbcDeprTableID.Text) & COMMA) 'varchar[20], NULL
            sSQL.Append("AssignmentTypeID = " & SQLString(tdbcAssignmentTypeID.Text) & COMMA) 'varchar[20], NOT NULL
        End If

        sSQL.Append("FullNameU = " & SQLStringUnicode(txtEmployeeName.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("UnitNameU = " & SQLStringUnicode(txtUnitName.Text, gbUnicode, True) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("NotesU = " & SQLStringUnicode(txtNotes.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("SpecificationU = " & SQLStringUnicode(txtSpecification.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("Index1 = " & SQLMoney(txtIndex1.Text, "N6") & COMMA) 'money, NULL
        sSQL.Append("Index2 = " & SQLMoney(txtIndex2.Text, "N6") & COMMA) 'money, NULL
        sSQL.Append("Index3 = " & SQLMoney(txtIndex3.Text, "N6") & COMMA) 'money, NULL
        sSQL.Append("Index4 = " & SQLMoney(txtIndex4.Text, "N6") & COMMA) 'money, NULL
        sSQL.Append("Index5 = " & SQLMoney(txtIndex5.Text, "N6") & COMMA) 'money, NULL
        sSQL.Append("Index6 = " & SQLMoney(txtIndex6.Text, "N6") & COMMA) 'money, NULL
        sSQL.Append("ACode01ID = " & SQLString(tdbcAcode01ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode02ID = " & SQLString(tdbcAcode02ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode03ID = " & SQLString(tdbcAcode03ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode04ID = " & SQLString(tdbcAcode04ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode05ID = " & SQLString(tdbcAcode05ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode06ID = " & SQLString(tdbcAcode06ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode07ID = " & SQLString(tdbcAcode07ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode08ID = " & SQLString(tdbcAcode08ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode09ID = " & SQLString(tdbcAcode09ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ACode10ID = " & SQLString(tdbcAcode10ID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NOT NULL
        sSQL.Append("ToolU = " & SQLStringUnicode(txtTool.Text, gbUnicode, True) & COMMA) 'varchar[100], NULL
        sSQL.Append("Maintainable = " & SQLNumber(chkMaintainable.Checked) & COMMA) 'tinyint, NOT NULL
        sSQL.Append("SupplierOTID = " & SQLString(tdbcSupplierOTID.Text) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("SupplierID = " & SQLString(tdbcSupplierID.Text) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("PurchaseDate = " & SQLDateSave(c1datePurchaseDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("DepDate = " & SQLDateSave(c1dateDepDate.Value) & COMMA)
        sSQL.Append("FANum01 = " & SQLMoney(txtFANum01.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum02 = " & SQLMoney(txtFANum02.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum03 = " & SQLMoney(txtFANum03.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum04 = " & SQLMoney(txtFANum04.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum05 = " & SQLMoney(txtFANum05.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum06 = " & SQLMoney(txtFANum06.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum07 = " & SQLMoney(txtFANum07.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum08 = " & SQLMoney(txtFANum08.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum09 = " & SQLMoney(txtFANum09.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FANum10 = " & SQLMoney(txtFANum10.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("FADate01 = " & SQLDateSave(c1dateFADate01.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate02 = " & SQLDateSave(c1dateFADate02.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate03 = " & SQLDateSave(c1dateFADate03.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate04 = " & SQLDateSave(c1dateFADate04.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate05 = " & SQLDateSave(c1dateFADate05.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate06 = " & SQLDateSave(c1dateFADate06.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate07 = " & SQLDateSave(c1dateFADate07.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate08 = " & SQLDateSave(c1dateFADate08.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate09 = " & SQLDateSave(c1dateFADate09.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FADate10 = " & SQLDateSave(c1dateFADate10.Text) & COMMA) 'datetime, NULL
        sSQL.Append("FAString01U = " & SQLStringUnicode(txtFAString01.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString02U = " & SQLStringUnicode(txtFAString02.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString03U = " & SQLStringUnicode(txtFAString03.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString04U = " & SQLStringUnicode(txtFAString04.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString05U = " & SQLStringUnicode(txtFAString05.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString06U = " & SQLStringUnicode(txtFAString06.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString07U = " & SQLStringUnicode(txtFAString07.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString08U = " & SQLStringUnicode(txtFAString08.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString09U = " & SQLStringUnicode(txtFAString09.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("FAString10U = " & SQLStringUnicode(txtFAString10.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("LocationID = " & SQLString(ReturnValueC1Combo(tdbcLocationID)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("MaintainDate = " & SQLDateSave(c1dateMaintainDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("ManagementObjID  = " & SQLString(ReturnValueC1Combo(tdbcObjectID2)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("ManagementObjTypeID = " & SQLString(ReturnValueC1Combo(tdbcObjectTypeID2)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("AssetConditionID = " & SQLString(ReturnValueC1Combo(tdbcAssetConditionName))) 'varchar[50], NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("AssetID = " & SQLString(txtAssetID.Text))
        Return sSQL
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T4001
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 21/09/2009 11:42:39
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T4001() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T4001"
        sSQL &= " Where "
        sSQL &= "DivisionID = " & SQLString(gsDivisionID) & " And "
        sSQL &= "AssetID = " & SQLString(txtAssetID.Text)
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0001
    '# Create User: Hoàng Đức Thịnh
    '# Create Date: 03/08/2006 10:40:06
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0001() As String
        Dim sSQL As String = ""
        sSQL = "Insert Into D02T0001("
        sSQL &= "TranMonth, TranYear, AssetS1ID, AssetS2ID, AssetS3ID, AssetID, AssetName, AssetNameU, "
        sSQL &= "ShortNameU, AssetAccountID, DepAccountID, MethodID, MethodEndID, DeprTableID,CountryID, "
        sSQL &= "MadeYear, UseMonth, UseYear, DepMonth, DepYear, Version, SeriNo, AssetNo, AssetTagU, "
        sSQL &= "NotesU, SpecificationU, ObjectTypeID, ObjectID, EmployeeID, FullNameU, "
        sSQL &= "ConvertedAmount, DepreciatedAmount, ServiceLife, DepreciatedPeriod, Percentage,"
        sSQL &= "AmountDepreciation, RemainAmount,UnitNameU, Index1, Index2,Index3,Index4, Index5, Index6, "
        sSQL &= "ACode01ID, ACode02ID, ACode03ID, ACode04ID, ACode05ID, ACode06ID, ACode07ID,ACode08ID, ACode09ID, ACode10ID, "
        sSQL &= "ToolU, DivisionID, AssignmentTypeID, Maintainable, SupplierOTID, SupplierID, PurchaseDate,"
        sSQL &= "CreateDate, LastModifyDate, CreateUserID, LastModifyUserID, DepDate,"
        sSQL &= "FANum01, FANum02, "
        sSQL &= "FANum03, FANum04, FANum05, FANum06, FANum07, "
        sSQL &= "FANum08, FANum09, FANum10, FADate01, FADate02, "
        sSQL &= "FADate03, FADate04, FADate05, FADate06, FADate07, "
        sSQL &= "FADate08, FADate09, FADate10, FAString01U, FAString02U, "
        sSQL &= "FAString03U, FAString04U, FAString05U, FAString06U, FAString07U, "
        sSQL &= "FAString08U, FAString09U, FAString10U,LocationID, MaintainDate, ManagementObjID, ManagementObjTypeID, AssetConditionID" ' ManagementObjID"
        sSQL &= ") Values("
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, TinyInt, NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, SmallInt, NULL
        If iAssetAuto = 2 Then
            sSQL &= SQLString("") & COMMA 'AssetS1ID, VarChar[20], NULL
            sSQL &= SQLString("") & COMMA 'AssetS2ID, VarChar[20], NULL
            sSQL &= SQLString("") & COMMA 'AssetS3ID, VarChar[20], NULL
        Else
            sSQL &= SQLString(tdbcAssetS1ID.Text) & COMMA 'AssetS1ID, VarChar[20], NULL
            sSQL &= SQLString(tdbcAssetS2ID.Text) & COMMA 'AssetS2ID, VarChar[20], NULL
            sSQL &= SQLString(tdbcAssetS3ID.Text) & COMMA 'AssetS3ID, VarChar[20], NULL
        End If

        sSQL &= SQLString(txtAssetID.Text) & COMMA 'AssetID [KEY], VarChar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'AssetName, VarChar[100], NULL
        sSQL &= SQLStringUnicode(txtAssetName.Text, gbUnicode, True) & COMMA 'AssetNameU, VarChar[100], NULL
        sSQL &= SQLStringUnicode(txtShortName.Text, gbUnicode, True) & COMMA 'ShortNameU, VarChar[20], NULL
        sSQL &= SQLString(tdbcAssetAccountID.Text) & COMMA 'AssetAccountID, VarChar[20], NULL
        sSQL &= SQLString(tdbcDepAccountID.Text) & COMMA 'DepAccountID, VarChar[20], NULL
        sSQL &= SQLNumber(tdbcMethodID.Text) & COMMA 'MethodID, TinyInt, NOT NULL
        sSQL &= SQLNumber(tdbcMethodEndID.Text) & COMMA 'MethodEndID, TinyInt, NOT NULL
        sSQL &= SQLString(tdbcDeprTableID.Text) & COMMA 'DeprTableID, VarChar[20], NULL
        sSQL &= SQLString(txtCountryID.Text) & COMMA 'CountryID, VarChar[20], NULL
        sSQL &= SQLNumber(txtMadeYear.Text) & COMMA 'MadeYear, Int, NULL
        If c1datePeriod.Text.Length > 4 Then
            sSQL &= SQLNumber(c1datePeriod.Text.Substring(0, 2)) & COMMA 'UseMonth, TinyInt, NULL
            sSQL &= SQLNumber(c1datePeriod.Text.Substring(3, 4)) & COMMA 'UseYear, SmallInt, NULL
        Else
            sSQL &= SQLNumber(0) & COMMA
            sSQL &= SQLNumber(0) & COMMA
        End If

        If c1dateDepPeriod.Text.Length > 4 Then
            sSQL &= SQLNumber(c1dateDepPeriod.Text.Substring(0, 2)) & COMMA 'DepMonth, TinyInt, NULL
            sSQL &= SQLNumber(c1dateDepPeriod.Text.Substring(3, 4)) & COMMA 'DepYear, SmallInt, NULL
        Else
            sSQL &= SQLNumber(0) & COMMA
            sSQL &= SQLNumber(0) & COMMA
        End If
        sSQL &= SQLString(txtVersion.Text) & COMMA 'Version, VarChar[20], NULL
        sSQL &= SQLString(txtSeriNo.Text) & COMMA 'SeriNo, VarChar[20], NULL
        sSQL &= SQLString(txtAssetNo.Text) & COMMA 'AssetNo, VarChar[20], NULL
        sSQL &= SQLStringUnicode(txtAssetTag.Text, gbUnicode, True) & COMMA 'AssetTag, VarChar[100], NULL
        sSQL &= SQLStringUnicode(txtNotes.Text, gbUnicode, True) & COMMA 'Notes, VarChar[250], NULL
        sSQL &= SQLStringUnicode(txtSpecification.Text, gbUnicode, True) & COMMA 'Specification, VarChar[250], NULL
        sSQL &= SQLString(tdbcObjectTypeID.Text) & COMMA 'ObjectTypeID, VarChar[20], NULL
        sSQL &= SQLString(tdbcObjectID.Text) & COMMA 'ObjectID, VarChar[20], NULL
        sSQL &= SQLString(tdbcEmployeeID.Text) & COMMA 'EmployeeID, VarChar[20], NULL
        sSQL &= SQLStringUnicode(txtEmployeeName.Text, gbUnicode, True) & COMMA 'FullName, VarChar[250], NULL
        sSQL &= SQLMoney(txtConvertedAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA 'ConvertedAmount, Money, NULL
        sSQL &= SQLMoney(txtDepreciationAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA 'DepreciatedAmount, Money, NULL
        sSQL &= SQLNumber(txtServiceLife.Text) & COMMA 'ServiceLife, Int, NULL
        sSQL &= SQLNumber(txtDepreciatedPeriod.Text) & COMMA 'DepreciatedPeriod, Int, NULL
        sSQL &= SQLMoney(txtPercentage.Text, DxxFormat.DefaultNumber2) & COMMA 'Percentage, Money, NULL
        sSQL &= SQLMoney(txtAmountDepreciation.Text, DxxFormat.D90_ConvertedDecimals) & COMMA 'AmountDepreciation, Money, NULL
        sSQL &= SQLMoney(txtRemainAmount.Text, DxxFormat.D90_ConvertedDecimals) & COMMA 'RemainAmount, Money, NULL
        sSQL &= SQLStringUnicode(txtUnitName.Text, gbUnicode, True) & COMMA 'UnitName, VarChar[20], NOT NULL
        sSQL &= SQLMoney(txtIndex1.Text, "N6") & COMMA 'Index1, Money, NULL
        sSQL &= SQLMoney(txtIndex2.Text, "N6") & COMMA 'Index2, Money, NULL
        sSQL &= SQLMoney(txtIndex3.Text, "N6") & COMMA 'Index3, Money, NULL
        sSQL &= SQLMoney(txtIndex4.Text, "N6") & COMMA 'Index4, Money, NULL
        sSQL &= SQLMoney(txtIndex5.Text, "N6") & COMMA 'Index5, Money, NULL
        sSQL &= SQLMoney(txtIndex6.Text, "N6") & COMMA 'Index6, Money, NULL
        sSQL &= SQLString(tdbcAcode01ID.Text) & COMMA 'ACode01ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode02ID.Text) & COMMA 'ACode02ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode03ID.Text) & COMMA 'ACode03ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode04ID.Text) & COMMA 'ACode04ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode05ID.Text) & COMMA 'ACode05ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode06ID.Text) & COMMA 'ACode06ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode07ID.Text) & COMMA 'ACode07ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode08ID.Text) & COMMA 'ACode08ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode09ID.Text) & COMMA 'ACode09ID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAcode10ID.Text) & COMMA 'ACode10ID, VarChar[20], NULL
        sSQL &= SQLStringUnicode(txtTool.Text, gbUnicode, True) & COMMA 'Tool, VarChar[100], NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, VarChar[20], NULL
        sSQL &= SQLString(tdbcAssignmentTypeID.Text) & COMMA 'AssignmentTypeID, VarChar[20], NOT NULL
        sSQL &= SQLNumber(chkMaintainable.Checked) & COMMA 'Maintainable, TinyInt, NOT NULL
        sSQL &= SQLString(tdbcSupplierOTID.Text) & COMMA 'SupplierOTID, VarChar[20], NOT NULL
        sSQL &= SQLString(tdbcSupplierID.Text) & COMMA 'SupplierID, VarChar[20], NOT NULL
        sSQL &= SQLDateSave(c1datePurchaseDate.Value) & COMMA 'PurchaseDate, DateTime, NULL

        If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
            sSQL &= "GetDate()" & COMMA 'CreateDate, DateTime, NOT NULL
        Else
            sSQL &= SQLDateTimeSave(sCreateDate) & COMMA 'CreateDate, datetime, NULL
        End If
        sSQL &= "GetDate()" & COMMA 'LastModifyDate, DateTime, NOT NULL
        If _FormState = EnumFormState.FormAdd OrElse _FormState = EnumFormState.FormCopy Then
            sSQL &= SQLString(gsUserID) & COMMA 'CreateUserID, VarChar[20], NOT NULL
        Else
            sSQL &= SQLString(sCreateUserID) & COMMA 'CreateUserID, varchar[20], NOT NULL
        End If
        sSQL &= SQLString(gsUserID) & COMMA 'LastModifyUserID, VarChar[20], NOT NULL
        sSQL &= SQLDateSave(c1dateDepDate.Value) & COMMA
        sSQL &= SQLMoney(txtFANum01.Text) & COMMA 'FANum01, money, NOT NULL
        sSQL &= SQLMoney(txtFANum02.Text) & COMMA 'FANum02, money, NOT NULL
        sSQL &= SQLMoney(txtFANum03.Text) & COMMA 'FANum03, money, NOT NULL
        sSQL &= SQLMoney(txtFANum04.Text) & COMMA 'FANum04, money, NOT NULL
        sSQL &= SQLMoney(txtFANum05.Text) & COMMA 'FANum05, money, NOT NULL
        sSQL &= SQLMoney(txtFANum06.Text) & COMMA 'FANum06, money, NOT NULL
        sSQL &= SQLMoney(txtFANum07.Text) & COMMA 'FANum07, money, NOT NULL
        sSQL &= SQLMoney(txtFANum08.Text) & COMMA 'FANum08, money, NOT NULL
        sSQL &= SQLMoney(txtFANum09.Text) & COMMA 'FANum09, money, NOT NULL
        sSQL &= SQLMoney(txtFANum10.Text) & COMMA 'FANum10, money, NOT NULL
        sSQL &= SQLDateSave(c1dateFADate01.Value) & COMMA 'FADate01, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate02.Value) & COMMA 'FADate02, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate03.Value) & COMMA 'FADate03, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate04.Value) & COMMA 'FADate04, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate05.Value) & COMMA 'FADate05, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate06.Value) & COMMA 'FADate06, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate07.Value) & COMMA 'FADate07, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate08.Value) & COMMA 'FADate08, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate09.Value) & COMMA 'FADate09, datetime, NULL
        sSQL &= SQLDateSave(c1dateFADate10.Value) & COMMA 'FADate10, datetime, NULL
        sSQL &= SQLStringUnicode(txtFAString01.Text, gbUnicode, True) & COMMA 'FAString01, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString02.Text, gbUnicode, True) & COMMA 'FAString02U, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString03.Text, gbUnicode, True) & COMMA 'FAString03U, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString04.Text, gbUnicode, True) & COMMA 'FAString04U, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString05.Text, gbUnicode, True) & COMMA 'FAString05U, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString06.Text, gbUnicode, True) & COMMA 'FAString06U, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString07.Text, gbUnicode, True) & COMMA 'FAString07U, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString08.Text, gbUnicode, True) & COMMA 'FAString08U, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString09.Text, gbUnicode, True) & COMMA 'FAString09U, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtFAString10.Text, gbUnicode, True) & COMMA 'FAStrng10U, nvarchar, NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcLocationID)) & COMMA 'LocationID, varchar[50], NOT NULL
        sSQL &= SQLDateSave(c1dateMaintainDate.Value) & COMMA 'MaintainDate, datetime, NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcObjectID2)) & COMMA 'ManagementObjID  , varchar[50], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcObjectTypeID2)) & COMMA 'ManagementObjTypeID, varchar[50], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAssetConditionName)) 'AssetConditionID, varchar[50], NOT NULL
        sSQL &= ")"
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T1001
    '# Created User: HUỲNH KHANH
    '# Created Date: 10/10/2014 04:45:25
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T1001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Luu du lieu vao bang D02T1001" & vbCrlf)
        sSQL.Append("Insert Into D02T1001(")
        sSQL.Append("AssetID, AssetNameU, NotesU, ")
        sSQL.Append("DivisionID, CreateUserID, CreateDate, LastModifyUserID, LastModifyDate, ")
        sSQL.Append("ACode01ID, ACode02ID, ACode03ID, ACode04ID, ACode05ID, ")
        sSQL.Append("ACode06ID, ACode07ID, ACode08ID, ACode09ID, ACode10ID, ")
        sSQL.Append(" ToolU, ObjectTypeID, ObjectID, " & vbCrLf)
        sSQL.Append("ManagementObjTypeID, ManagementObjID, MaintainDate, AssetConditionID, " & vbCrLf)
        sSQL.Append("SpecificationU, CountryID, MadeYear, SeriNo, " & vbCrLf)
        sSQL.Append("Version, Index1, Index2, Index3, Index4, Index5, Index6" & vbCrLf)
        sSQL.Append(", SetupDate, SetupVoucherID, OQuantity, CQuantity,ReceiverID, LocationID, SupplierOTID, SupplierID, ChargeObjType,")

        sSQL.Append("FAString01U, FAString02U, " & vbCrLf)
        sSQL.Append("FAString03U, FAString04U, FAString05U, " & vbCrLf)
        sSQL.Append("FAString06U, FAString07U, " & vbCrLf)
        sSQL.Append("FAString08U, FAString09U, FAString10U, " & vbCrLf)
        sSQL.Append("FANum01, FANum02, FANum03, FANum04, FANum05, " & vbCrLf)
        sSQL.Append("FANum06, FANum07, FANum08, FANum09, FANum10, " & vbCrLf)
        sSQL.Append("FADate01, FADate02, FADate03, FADate04, FADate05, " & vbCrLf)
        sSQL.Append("FADate06, FADate07, FADate08, FADate09, FADate10")

        sSQL.Append(") Values(" & vbCrLf)
        sSQL.Append(SQLString(txtAssetID.Text) & COMMA) 'AssetID [KEY], varchar[50], NOT NULL
        sSQL.Append(SQLStringUnicode(txtAssetName, True) & COMMA) 'AssetNameU, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtNotes, True) & COMMA) 'NotesU, nvarchar[1000], NOT NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[50], NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[50], NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode01ID)) & COMMA) 'ACode01ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode02ID)) & COMMA) 'ACode02ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode03ID)) & COMMA) 'ACode03ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode04ID)) & COMMA) 'ACode04ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode05ID)) & COMMA) 'ACode05ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode06ID)) & COMMA) 'ACode06ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode07ID)) & COMMA) 'ACode07ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode08ID)) & COMMA) 'ACode08ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode09ID)) & COMMA) 'ACode09ID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAcode10ID)) & COMMA) 'ACode10ID, varchar[20], NOT NULL
        sSQL.Append(SQLStringUnicode(txtTool, True) & COMMA) 'ToolU, nvarchar[500], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcObjectTypeID6)) & COMMA) 'ObjectTypeID, varchar[50], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcObjectID6)) & COMMA & vbCrLf) 'ObjectID, varchar[50], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcManagementObTypeID6)) & COMMA) 'ManagementObTypeID, varchar[50], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcManagementObID6)) & COMMA) 'ManagementObID, varchar[50], NOT NULL
        sSQL.Append(SQLDateSave(c1dateMaintainDate.Value) & COMMA) 'MaintainDate, datetime, NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAssetConditionName)) & COMMA & vbCrLf) 'AssetConditionID, varchar[50], NOT NULL
        sSQL.Append(SQLStringUnicode(txtSpecification, True) & COMMA) 'SpecificationU, nvarchar[2000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtCountryID, False) & COMMA) 'CountryID, varchar[500], NOT NULL
        sSQL.Append(SQLNumber(txtMadeYear.Text) & COMMA) 'MadeYear, int, NOT NULL
        sSQL.Append(SQLStringUnicode(txtSeriNo, False) & COMMA) 'SeriNo, varchar[100], NOT NULL
        sSQL.Append(SQLString(txtVersion.Text) & COMMA) 'Version, varchar[20], NOT NULL
        sSQL.Append(SQLMoney(txtIndex1.Text) & COMMA & vbCrLf) 'Index1, money, NOT NULL
        sSQL.Append(SQLMoney(txtIndex2.Text) & COMMA) 'Index2, money, NOT NULL
        sSQL.Append(SQLMoney(txtIndex3.Text) & COMMA) 'Index3, money, NOT NULL
        sSQL.Append(SQLMoney(txtIndex4.Text) & COMMA) 'Index4, money, NOT NULL
        sSQL.Append(SQLMoney(txtIndex5.Text) & COMMA & vbCrLf) 'Index5, money, NOT NULL
        sSQL.Append(SQLMoney(txtIndex6.Text) & COMMA) 'Index6, money, NOT NULL
        sSQL.Append(SQLDateSave(c1dateSetupDate.Value) & COMMA)
        sSQL.Append(SQLString(txtSetupVoucherID.Text) & COMMA)
        sSQL.Append(SQLMoney(cneOQuantity.Text, DxxFormat.D07_QuantityDecimals) & COMMA)
        sSQL.Append(SQLMoney(txtCQuantity.Text, DxxFormat.D07_QuantityDecimals) & COMMA)
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcReceiverID)) & COMMA) 'ReceiverID, varchar[20], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcLocationIDID6)) & COMMA) 'LocationID, varchar[50], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcSupplierOTIDID6)) & COMMA) 'SupplierOTID, varchar[50], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcSupplierIDID6)) & COMMA) 'SupplierID, varchar[50], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcChargeObjType)) & COMMA) 'ChargeObjType, varchar[50], NOT NULL  --ID : 252774
        sSQL.Append(SQLStringUnicode(txtFAString01, True) & COMMA) 'FAString01U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString02, True) & COMMA) 'FAString02U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString03, True) & COMMA) 'FAString03U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString04, True) & COMMA & vbCrLf) 'FAString04U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString05, True) & COMMA) 'FAString05U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString06, True) & COMMA & vbCrLf) 'FAString06U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString07, True) & COMMA) 'FAString07U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString08, True) & COMMA & vbCrLf) 'FAString08U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString09, True) & COMMA) 'FAString09U, nvarchar[500], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString10, True) & COMMA) 'FAString10U, nvarchar[500], NOT NULL
        sSQL.Append(SQLMoney(txtFANum01.Text) & COMMA & vbCrLf) 'FANum01, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum02.Text) & COMMA) 'FANum02, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum03.Text) & COMMA) 'FANum03, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum04.Text) & COMMA) 'FANum04, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum05.Text) & COMMA & vbCrLf) 'FANum05, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum06.Text) & COMMA) 'FANum06, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum07.Text) & COMMA) 'FANum07, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum08.Text) & COMMA) 'FANum08, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum09.Text) & COMMA & vbCrLf) 'FANum09, decimal, NOT NULL
        sSQL.Append(SQLMoney(txtFANum10.Text) & COMMA) 'FANum10, decimal, NOT NULL
        sSQL.Append(SQLDateSave(c1dateFADate01.Value) & COMMA) 'FADate01, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate02.Value) & COMMA) 'FADate02, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate03.Value) & COMMA & vbCrLf) 'FADate03, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate04.Value) & COMMA) 'FADate04, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate05.Value) & COMMA) 'FADate05, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate06.Value) & COMMA) 'FADate06, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate07.Value) & COMMA & vbCrLf) 'FADate07, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate08.Value) & COMMA) 'FADate08, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate09.Value) & COMMA) 'FADate09, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate10.Value)) 'FADate10, datetime, NULL

        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1037
    '# Created User: HUỲNH KHANH
    '# Created Date: 10/10/2014 04:49:40
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1037(ByVal iMode As Integer) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Luu danh muc ccdc" & vbCrLf)
        sSQL &= "Exec D02P1037 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[250], NOT NULL
        sSQL &= SQLString(txtAssetID.Text) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcUnitID)) & COMMA 'UnitID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcAccountID)) & COMMA 'AccountID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcMethodIDCCDC)) & COMMA 'MethodID, varchar[20], NOT NULL
        sSQL &= SQLNumber(iMode) 'Mode, tinyint, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T1001
    '# Created User: HUỲNH KHANH
    '# Created Date: 10/10/2014 04:53:27
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T1001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Cap nhat du lieu bang D02T1001" & vbCrlf)
        sSQL.Append("Update D02T1001 Set ")
        'sSQL.Append("AssetID = " & SQLString(txtAssetID.Text) & COMMA) '[KEY], varchar[50], NOT NULL
        sSQL.Append("AssetNameU = " & SQLStringUnicode(txtAssetName, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("NotesU = " & SQLStringUnicode(txtNotes, True) & COMMA) 'nvarchar[1000], NOT NULL
        sSQL.Append("DivisionID = " & SQLString(gsDivisionID) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NOT NULL
        sSQL.Append("ACode01ID = " & SQLString(ReturnValueC1Combo(tdbcAcode01ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode02ID = " & SQLString(ReturnValueC1Combo(tdbcAcode02ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode03ID = " & SQLString(ReturnValueC1Combo(tdbcAcode03ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode04ID = " & SQLString(ReturnValueC1Combo(tdbcAcode04ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode05ID = " & SQLString(ReturnValueC1Combo(tdbcAcode05ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode06ID = " & SQLString(ReturnValueC1Combo(tdbcAcode06ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode07ID = " & SQLString(ReturnValueC1Combo(tdbcAcode07ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode08ID = " & SQLString(ReturnValueC1Combo(tdbcAcode08ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode09ID = " & SQLString(ReturnValueC1Combo(tdbcAcode09ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ACode10ID = " & SQLString(ReturnValueC1Combo(tdbcAcode10ID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ToolU = " & SQLStringUnicode(txtTool, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("ObjectTypeID = " & SQLString(ReturnValueC1Combo(tdbcObjectTypeID6)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("ObjectID = " & SQLString(ReturnValueC1Combo(tdbcObjectID6)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("ManagementObjTypeID = " & SQLString(ReturnValueC1Combo(tdbcManagementObTypeID6)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("ManagementObjID = " & SQLString(ReturnValueC1Combo(tdbcManagementObID6)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("MaintainDate = " & SQLDateSave(c1dateMaintainDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("AssetConditionID = " & SQLString(ReturnValueC1Combo(tdbcAssetConditionName)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("SpecificationU = " & SQLStringUnicode(txtSpecification, True) & COMMA) 'nvarchar[2000], NOT NULL
        sSQL.Append("CountryID = " & SQLStringUnicode(txtCountryID, False) & COMMA) 'varchar[500], NOT NULL
        sSQL.Append("MadeYear = " & SQLNumber(txtMadeYear.Text) & COMMA) 'int, NOT NULL
        sSQL.Append("SeriNo = " & SQLStringUnicode(txtSeriNo, False) & COMMA) 'varchar[100], NOT NULL
        sSQL.Append("Version = " & SQLString(txtVersion.Text) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("Index1 = " & SQLMoney(txtIndex1.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("Index2 = " & SQLMoney(txtIndex2.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("Index3 = " & SQLMoney(txtIndex3.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("Index4 = " & SQLMoney(txtIndex4.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("Index5 = " & SQLMoney(txtIndex5.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("Index6 = " & SQLMoney(txtIndex6.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("OQuantity = " & SQLMoney(cneOQuantity.Text, DxxFormat.D07_QuantityDecimals) & COMMA) 'datetime, NULL
        sSQL.Append("CQuantity = " & SQLMoney(txtCQuantity.Text, DxxFormat.D07_QuantityDecimals) & COMMA) 'datetime, NULL
        sSQL.Append("SetupDate = " & SQLDateSave(c1dateSetupDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("SetupVoucherID = " & SQLString(txtSetupVoucherID.Text) & COMMA) 'varchar[500], NOT NULL
        sSQL.Append("ReceiverID = " & SQLString(ReturnValueC1Combo(tdbcReceiverID)) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("LocationID = " & SQLString(ReturnValueC1Combo(tdbcLocationIDID6)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("SupplierOTID = " & SQLString(ReturnValueC1Combo(tdbcSupplierOTIDID6)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("SupplierID = " & SQLString(ReturnValueC1Combo(tdbcSupplierIDID6)) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("ChargeObjType = " & SQLString(ReturnValueC1Combo(tdbcChargeObjType)) & COMMA) 'varchar[50], NOT NULL --ID  : 252774
        sSQL.Append("FAString01U = " & SQLStringUnicode(txtFAString01, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString02U = " & SQLStringUnicode(txtFAString02, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString03U = " & SQLStringUnicode(txtFAString03, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString04U = " & SQLStringUnicode(txtFAString04, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString05U = " & SQLStringUnicode(txtFAString05, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString06U = " & SQLStringUnicode(txtFAString06, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString07U = " & SQLStringUnicode(txtFAString07, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString08U = " & SQLStringUnicode(txtFAString08, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString09U = " & SQLStringUnicode(txtFAString09, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FAString10U = " & SQLStringUnicode(txtFAString10, True) & COMMA) 'nvarchar[500], NOT NULL
        sSQL.Append("FANum01 = " & SQLMoney(txtFANum01.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum02 = " & SQLMoney(txtFANum02.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum03 = " & SQLMoney(txtFANum03.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum04 = " & SQLMoney(txtFANum04.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum05 = " & SQLMoney(txtFANum05.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum06 = " & SQLMoney(txtFANum06.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum07 = " & SQLMoney(txtFANum07.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum08 = " & SQLMoney(txtFANum08.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum09 = " & SQLMoney(txtFANum09.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FANum10 = " & SQLMoney(txtFANum10.Text) & COMMA) 'decimal, NOT NULL
        sSQL.Append("FADate01 = " & SQLDateSave(c1dateFADate01.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate02 = " & SQLDateSave(c1dateFADate02.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate03 = " & SQLDateSave(c1dateFADate03.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate04 = " & SQLDateSave(c1dateFADate04.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate05 = " & SQLDateSave(c1dateFADate05.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate06 = " & SQLDateSave(c1dateFADate06.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate07 = " & SQLDateSave(c1dateFADate07.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate08 = " & SQLDateSave(c1dateFADate08.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate09 = " & SQLDateSave(c1dateFADate09.Value) & COMMA) 'datetime, NULL
        sSQL.Append("FADate10 = " & SQLDateSave(c1dateFADate10.Value)) 'datetime, NULL

        sSQL.Append(" Where ")
        sSQL.Append("AssetID = " & SQLString(txtAssetID.Text))

        Return sSQL
    End Function



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T4001s
    '# Create User: Hoàng Đức Thịnh
    '# Create Date: 03/08/2006 11:54:07
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T4001s() As String
        Dim sRet As String = ""
        Dim sSQL As String = ""
        For i As Integer = 0 To tdbgDetail.RowCount - 1
            If tdbgDetail(i, COL_EquipmentID).ToString <> "" Then
                sSQL = ""
                sSQL &= "Insert Into D02T4001("
                sSQL &= "DivisionID, AssetID, EquipmentID, OrderNum, EquipmentNameU, " & vbCrLf
                sSQL &= "NotesU, EquipmentQuantity, ObjectTypeID, ObjectID, EquipmentValue, IsTool, " & vbCrLf
                sSQL &= "UnitPrice, TaxAmount, AcceptanceTime, PurchaseDate"
                sSQL &= ") Values ("
                sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID [KEY], VarChar[20], NOT NULL
                sSQL &= SQLString(txtAssetID.Text) & COMMA 'AssetID [KEY], VarChar[20], NOT NULL
                sSQL &= SQLString(tdbgDetail(i, COL_EquipmentID)) & COMMA 'EquipmentID [KEY], VarChar[20], NOT NULL
                sSQL &= SQLNumber(tdbgDetail(i, COL_OrderNum)) & COMMA 'OrderNum, BigInt, NULL
                sSQL &= SQLStringUnicode(tdbgDetail(i, COL_EquipmentName), gbUnicode, True) & COMMA & vbCrLf 'EquipmentName, VarChar[50], NULL
                sSQL &= SQLStringUnicode(tdbgDetail(i, COL_Notes), gbUnicode, True) & COMMA 'Notes, VarChar[250], NULL
                sSQL &= SQLMoney(tdbgDetail(i, COL_EquipmentQuantity)) & COMMA 'EquipmentQuantity, Decimal, NULL
                sSQL &= SQLString(tdbgDetail(i, COL_ObjectTypeID)) & COMMA 'ObjectTypeID, VarChar[20], NULL
                sSQL &= SQLString(tdbgDetail(i, COL_ObjectID)) & COMMA 'ObjectID, VarChar[20], NULL
                sSQL &= SQLMoney(tdbgDetail(i, COL_EquipmentValue)) & COMMA 'EquipmentValue, Money, NOT NULL
                sSQL &= SQLNumber(chkIsTools.Checked) & COMMA & vbCrLf
                '7/4/2017, 	Phạm Thị Thu: id 96093-[CDS] Thẻ TSCĐ - Danh mục TSCĐ theo chủng loại
                sSQL &= SQLMoney(tdbgDetail(i, COL_UnitPrice)) & COMMA 'UnitPrice, Decimal, NULL
                sSQL &= SQLMoney(tdbgDetail(i, COL_TaxAmount)) & COMMA 'TaxAmount, Decimal, NULL
                sSQL &= SQLDateSave(tdbgDetail(i, COL_AcceptanceTime)) & COMMA 'AcceptanceTime, DateTime, NULL
                sSQL &= SQLDateSave(tdbgDetail(i, COL_PurchaseDate))  'PurchaseDate, DateTime, NULL

                sSQL &= ")"
                sRet &= sSQL & vbCrLf
            End If

        Next
        Return sRet
    End Function

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        ClearText(Me)
        _FormState = EnumFormState.FormAdd
        _bFormD02F0087 = False
        tdbcAssetS1ID.SelectedValue = "-1"
        tdbcAssetS2ID.SelectedValue = "-1"
        tdbcAssetS3ID.SelectedValue = "-1"
        'txtAssetID.Text = ""
        txtAssetName.Text = ""
        txtNotes.Text = ""
        tdbcObjectTypeID.SelectedValue = "-1"
        tdbcObjectTypeID2.SelectedValue = "-1"
        tdbcEmployeeID.SelectedValue = "-1"
        txtEmployeeName.Text = ""
        txtShortName.Text = ""
        txtAssetTag.Text = ""
        txtAssetNo.Text = ""
        tdbcSupplierOTID.SelectedValue = "-1"
        chkMaintainable.Checked = False
        tdbcAssetAccountID.SelectedValue = "-1"
        tdbcDepAccountID.SelectedValue = "-1"
        GetAutoAssetInfo()
        'LoadAddNew()
        tdbcMethodID.SelectedValue = "-1"
        txtMethodName.Text = ""
        tdbcMethodEndID.SelectedValue = "-1"
        txtMethodEndName.Text = ""
        tdbcDeprTableID.SelectedValue = "-1"
        txtDeprTableName.Text = ""
        tdbcAssignmentTypeID.SelectedValue = "-1"
        txtAssignmentTypeName.Text = ""
        txtSpecification.Text = ""
        txtCountryID.Text = ""
        txtMadeYear.Text = ""
        txtSeriNo.Text = ""
        txtVersion.Text = ""
        txtUnitName.Text = ""
        txtTool.Text = ""
        txtIndex1.Text = ""
        txtIndex2.Text = ""
        txtIndex3.Text = ""
        txtIndex4.Text = ""
        txtIndex5.Text = ""
        txtIndex6.Text = ""
        mAssetID = ""
        LoadTDBGrid()
        tdbcAcode01ID.SelectedValue = "-1"
        tdbcAcode02ID.SelectedValue = "-1"
        tdbcAcode03ID.SelectedValue = "-1"
        tdbcAcode04ID.SelectedValue = "-1"
        tdbcAcode05ID.SelectedValue = "-1"
        tdbcAcode06ID.SelectedValue = "-1"
        tdbcAcode07ID.SelectedValue = "-1"
        tdbcAcode08ID.SelectedValue = "-1"
        tdbcAcode09ID.SelectedValue = "-1"
        tdbcAcode10ID.SelectedValue = "-1"
        picImage.Image = Nothing
        tab.SelectedTab = tab.TabPages(0)
        btnNext.Enabled = False
        btnSave.Enabled = True
        LoadAddNew()
        tdbcAssetS1ID.Focus()
        tdbcAssetS1ID_SelectedValueChanged(Nothing, Nothing)
        tdbcAssetS2ID_SelectedValueChanged(Nothing, Nothing)
        tdbcAssetS3ID_SelectedValueChanged(Nothing, Nothing)

        'tdbcUnitID.SelectedValue = "-1"
        'tdbcAccountID.SelectedValue = "-1"
        'tdbcMethodIDCCDC.SelectedValue = "-1"
        chkIsTools.Checked = False
        If iAssetAuto = 2 Then
            tdbcIGEMethodID.SelectedValue = sDefaultIGEMethodID
        End If

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0613
    '# Create User: Hoàng Đức Thịnh
    '# Create Date: 04/08/2006 07:54:22
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0613() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0613 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, VarChar[20], NOT NULL
        sSQL &= SQLString(mAssetID) 'AssetID, VarChar[20], NOT NULL
        Return sSQL
    End Function

    Private Function GetOpenFileDialog() As OpenFileDialog
        Dim open As OpenFileDialog = New OpenFileDialog()
        open.CheckPathExists = True
        open.CheckPathExists = True
        open.AddExtension = False
        open.Multiselect = False
        open.Title = "Add file"
        open.InitialDirectory = "C:\"
        open.Filter = "BMP file (*.bmp)|*.bmp|JPG file (*.jpg)|*.jpg|JPEG file (*.jpeg)|*.jpeg|GIF file (*.gif)|*.gif"
        open.FilterIndex = 1
        open.ValidateNames = True
        open.RestoreDirectory = True
        Return open
    End Function

    Private Sub SQLInsertD02T0004()
        If sPathImage = "" Then Exit Sub
        Dim image As Byte()
        Dim f As FileInfo = New FileInfo(sPathImage)

        Dim fileLength As Long = f.Length
        Dim fs As FileStream = New FileStream(sPathImage, FileMode.Open, FileAccess.Read, FileShare.Read)
        image = New Byte(Convert.ToInt32(fileLength)) {}
        Dim iBytesRead As Integer = fs.Read(image, 0, Convert.ToInt32(fileLength))
        fs.Close()
        Dim sSQL As String = ""
        sSQL = "Delete From D02T0004 Where AssetID=" & SQLString(txtAssetID.Text) & " and DivisionID=" & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "Insert Into D02T0004(AssetID, DivisionID, AssetImage, Extension) " & vbCrLf
        sSQL &= "Values(@AssetID, @DivisionID, @AssetImage, @Extension)"
        Dim conn As SqlConnection = New SqlConnection(gsConnectionString)
        Dim cmd As SqlCommand = New SqlCommand(sSQL, conn)

        cmd.Parameters.Add("@AssetID", SqlDbType.VarChar)
        cmd.Parameters.Add("@DivisionID", SqlDbType.VarChar)
        cmd.Parameters.Add("@AssetImage", SqlDbType.Image)
        cmd.Parameters.Add("@Extension", SqlDbType.VarChar)
        cmd.Parameters("@AssetID").Value = txtAssetID.Text
        cmd.Parameters("@DivisionID").Value = gsDivisionID
        cmd.Parameters("@AssetImage").Value = image
        cmd.Parameters("@Extension").Value = f.Extension
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()
        cmd.Dispose()
        conn.Dispose()
    End Sub

    Private Sub LoadImage()
        Dim dtImage As DataTable

        Dim sSQL As String = "Select AssetImage From D02T0004 WITH(NOLOCK) Where AssetID = " & SQLString(txtAssetID.Text) & " and DivisionID=" & SQLString(gsDivisionID)
        dtImage = ReturnDataTable(sSQL)
        If dtImage.Rows.Count = 0 Then Exit Sub

        picImage.Image = ReturnImage(dtImage.Rows(0).Item(0))
        'Dim conn As SqlConnection = New SqlConnection(gsConnectionString)
        'Dim cmd As SqlCommand = New SqlCommand(sSQL, conn)
        'conn.Open()
        'Dim image As Byte() = DirectCast(cmd.ExecuteScalar(), Byte())
        'conn.Close()
        'conn.Dispose()
        'sSQL = "Select Extension From D02T0004 Where AssetID = " & SQLString(txtAssetID.Text) & " and DivisionID=" & SQLString(gsDivisionID)
        'Dim FileName As String = "Image" & Trim(ReturnScalar(sSQL))
        'Dim FInfo As FileInfo = New FileInfo(FileName)
        'If FInfo.Exists Then FInfo.Delete()
        'Dim fs As FileStream = New FileStream(FileName, FileMode.CreateNew, FileAccess.Write)
        'fs.Write(image, 0, image.Length)
        'fs.Flush()
        'fs.Close()
        'fs.Dispose()
        'picImage.Load(FileName)
        'FInfo.Delete()
    End Sub

    Public Function ReturnImage(ByVal objExpression As Object) As Image
        Try
            If IsDBNull(objExpression) = False Then
                Dim ms As New System.IO.MemoryStream(CType(objExpression, Byte()))
                Dim img As Image = Image.FromStream(ms)
                Return img
            Else
                Return Nothing
            End If
        Catch ex As Exception

        End Try
        Return Nothing
    End Function

    'Public Function LoadImage1() As System.Drawing.Image
    '    Try
    '        Dim sSQL As String = "Select AssetImage From D02T0004 Where AssetID = " & SQLString(txtAssetID.Text) & " and DivisionID=" & SQLString(gsDivisionID)
    '        Dim cmdSelect As SqlCommand = New SqlCommand(sSQL, gConn)
    '        Dim byarrImg As Byte() = DirectCast(cmdSelect.ExecuteScalar(), Byte())
    '        Dim sfn As String = Convert.ToString(DateTime.Now.ToFileTime())
    '        Dim fs As FileStream = New FileStream(sfn, FileMode.CreateNew, FileAccess.Write)

    '        fs.Write(byarrImg, 0, byarrImg.Length)
    '        fs.Flush()
    '        fs.Close()

    '        cmdSelect.Dispose()

    '        Return Image.FromFile(sfn)
    '    Catch ex As Exception
    '        Return Nothing
    '    Finally
    '    End Try
    'End Function

    Private Sub btnImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImage.Click
        Dim open As OpenFileDialog = GetOpenFileDialog()
        open.ShowDialog()
        sPathImage = open.FileName
        open.Dispose()
        If sPathImage <> "" Then picImage.Load(sPathImage)
    End Sub

    Private Sub btnAttact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAttact.Click
        'Dim frm As New D91F4010
        'With frm
        '    .FormName = "D91F4010"
        '    .FormState = _FormState
        '    .TableName = "D02T0001" 'Truyền giá trị khác nhau từng module
        '    .Key01ID = txtAssetID.Text 'Giá trị khóa chính
        '    .Key02ID = "" 'Theo TL phân tích quy định
        '    .Key03ID = "" 'Theo TL phân tích quy định
        '    .Key04ID = "" 'Theo TL phân tích quy định
        '    .Key05ID = "" 'Theo TL phân tích quy định

        '    .ShowDialog()
        'End With
        'btnAttact.Text = rL3("Dinh_ke_m") & Space(1) & "(" & ReturnAttachmentNumber("D02T0001", txtAssetID.Text) & ")"  'Đính kèm

        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "TableName", "D02T0001")
        SetProperties(arrPro, "Key1ID", txtAssetID.Text)
        SetProperties(arrPro, "Status", L3Byte(IIf(_FormState = EnumFormState.FormView, 0, 1)))
        SetProperties(arrPro, "bNewDatabase", False) 'Lưu database mới ATT, không phải database hiện tại
        CallFormShowDialog("D91D0340", "D91F4010", arrPro)
        btnAttact.Text = rL3("Dinh_ke_m") & Space(1) & " (" & ReturnAttachmentNumber("D02T0001", txtAssetID.Text) & ")" 'Đính kèm
    End Sub

    Private Sub btnNote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNote.Click
        'Dim frm As New D91F4010
        'With frm
        '    .FormName = "D91F2010"
        '    .FormState = _FormState
        '    .TableName = "D02T0001" 'Truyền giá trị khác nhau từng module
        '    .Key01ID = txtAssetID.Text 'Giá trị khóa chính 
        '    .Key02ID = "" 'Theo TL phân tích quy định
        '    .Key03ID = "" 'Theo TL phân tích quy định
        '    .Key04ID = "" 'Theo TL phân tích quy định
        '    .Key05ID = "" 'Theo TL phân tích quy định

        '    .ShowDialog()
        'End With
        'btnNote.Text = rL3("Ghi__chu") & Space(1) & "(" & ReturnNotesNumber("D02T0001", txtAssetID.Text) & ")"  'Ghi chú
        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "TableName", "D02T0001")
        SetProperties(arrPro, "Key1ID", txtAssetID.Text)
        SetProperties(arrPro, "Status", L3Byte(IIf(_FormState = EnumFormState.FormView, 0, 1)))
        SetProperties(arrPro, "bNewDatabase", False) 'Lưu database mới ATT, không phải database hiện tại
        CallFormShowDialog("D91D0340", "D91F2010", arrPro)
        btnNote.Text = rL3("Ghi__chu") & Space(1) & "(" & ReturnNotesNumber("D02T0001", txtAssetID.Text) & ")"  'Ghi chú
    End Sub

    Private Sub txtAssetID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAssetID.LostFocus
        txtAssetID.Text = txtAssetID.Text.ToUpper()
    End Sub

    Private Sub txtMadeYear_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMadeYear.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub txtIndex1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIndex1.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtIndex2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIndex2.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtIndex3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIndex3.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtIndex4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIndex4.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtIndex5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIndex5.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtIndex6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIndex6.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub


    Private Sub c1dateTranDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles c1dateTranDate.KeyDown
        e.Handled = False
    End Sub

    Private Sub c1dateTranDate_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles c1dateTranDate.KeyPress
        e.Handled = False
    End Sub

    Private Sub c1dateDepPeriod_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles c1dateDepPeriod.KeyDown
        e.Handled = False
    End Sub

    Private Sub c1dateDepPeriod_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles c1dateDepPeriod.KeyPress
        e.Handled = False
    End Sub

    Private Sub c1datePeriod_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles c1datePeriod.KeyDown
        e.Handled = False
    End Sub

    Private Sub c1datePeriod_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles c1datePeriod.KeyPress
        e.Handled = False
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Thiet_lap_danh_muc_tai_san_co_dinh_-__D02F1031") & UnicodeCaption(gbUnicode) 'ThiÕt lËp danh móc tªi s¶n cç ¢Ünh -  D02F1031
        '================================================================ 
        lblDepDate.Text = rl3("Ngay_bat_dau_khau_hao")
        lblAssetS1ID.Text = rl3("Ma_tai_san") 'Mã tài sản
        lblAssetName.Text = rl3("Ten_tai_san") 'Tên tài sản
        lblObjectTypeID.Text = rl3("Bo_phan_tiep_nhan") 'Bộ phận tiếp nhận
        lblEmployeeID.Text = rl3("Nguoi_tiep_nhan") 'Người tiếp nhận
        lblShortName.Text = rl3("Ten_tat") 'Tên tắt
        lblAssetTag.Text = rl3("The_tai_san") 'Thẻ tài sản
        lblAssetNo.Text = rl3("So_hieu") 'Số hiệu
        lblSupplierOTID.Text = rl3("Nha_cung_cap") 'Nhà cung cấp
        lbltePurchaseDate.Text = rl3("Ngay_mua") 'Ngày mua
        lblConvertedAmount.Text = rl3("Nguyen_gia") 'Nguyên giá
        lblRemainAmount.Text = rl3("Gia_tri_con_lai") 'Giá trị còn lại
        lblDepreciatedPeriod.Text = rl3("Thoi_gian_da_khau_hao_(ky)") 'Thời gian đã khấu hao (kỳ)
        lblDepreciationAmount.Text = rl3("Muc_khau_hao") 'Mức khấu hao
        lblAmountDepreciation.Text = rl3("Hao_mon_luy_ke") 'Hao mòn lũy kế
        lblServiceLife.Text = rl3("Thoi_gian_khau_hao_(ky)") 'Thời gian khấu hao (kỳ)
        lblPercentage.Text = rl3("Ty_le_khau_hao") 'Tỷ lệ khấu hao
        lblMethodID.Text = rl3("Phuong_phap_KH") 'Phương pháp KH
        lblMethodEndID.Text = rl3("Xu_ly_khau_hao_ky_cuoi") 'Xử lý khấu hao kỳ cuối
        lblDeprTableID.Text = rl3("Bang_khau_hao") 'Bảng khấu hao
        lblAssignmentTypeID.Text = rl3("Kieu_phan_bo") 'Kiểu phân bổ
        lblPeriod.Text = rl3("Ky_su_dung") 'Kỳ sử dụng
        lblPeriodDep.Text = rl3("Ky_bat_dau_tinh_khau_hao") 'Kỳ bắt đầu tính khấu hao
        lblPeriodTran.Text = rl3("Ky_hinh_thanh") 'Kỳ hình thành
        lblAssetAccountID.Text = rl3("Tai_khoan_tai_san") 'Tài khoản tài sản
        lblDepAccountID.Text = rl3("Tai_khoan_khau_hao") 'Tài khoản khấu hao
        'lblIndex1.Text = rl3("Số_giờ_binh_quân") 'Số giờ bình quân
        'lblIndex2.Text = rl3("Số_giờ_thực_tế") 'Số giờ thực tế
        lblIndex3.Text = rl3("Chi_so") & " 3" 'Chỉ số 3
        lblIndex4.Text = rl3("Chi_so") & " 4" 'Chỉ số 4
        lblIndex5.Text = rl3("Chi_so") & " 5" 'Chỉ số 5
        lblIndex6.Text = rl3("Chi_so") & " 6" 'Chỉ số 6
        lblSpecification.Text = rl3("Dac_diem") 'Đặc điểm
        lblCountryID.Text = rl3("Nuoc_san_xuat") 'Nước sản xuất
        lblMadeYear.Text = rl3("Nam_san_xuat") 'Năm sản xuất
        lblSeriNo.Text = rl3("So_Seri") 'Số Sêri
        lblVersion.Text = rl3("The_he") 'Thế hệ
        lblTool.Text = rl3("Thiet_bi_dinh_kem") 'Thiết bị đính kèm
        lblUnitName.Text = rl3("Don_vi_tinh") 'Đơn vị tính
        lblNotes.Text = rL3("Ghi_chu") 'Ghi chú
        lblLocationID.Text = rL3("Vi_tri") 'Vị trí
        lblUnitID.Text = rL3("Don_vi_tinh") 'Đơn vị tính
        lblMethodIDCCDC.Text = rL3("Phuong_phap_tinh_gia") 'Phương pháp tính giá
        lblAccountID.Text = rL3("TK_ton_kho") 'TK tồn kho
        lblReceiverID.Text = rL3("Nguoi_tiep_nhan") 'Người tiếp nhận
        lblLocationIDID6.Text = rL3("Vi_tri") 'Vị trí
        lblSupplierID.Text = rL3("Nha_cung_cap") 'Nhà cung cấp
        '================================================================ 

        btnConvertedAmount.Text = rl3("Nguyen_gia") 'Nguyên giá
        btnDepreciate.Text = rl3("Khau_hao") 'Khấu hao
        btnHistory.Text = rl3("Lich_su") 'Lịch sử
        btnNote.Text = rl3("_Ghi_chu") '&Ghi chú
        btnAttact.Text = rl3("Dinh__kem") 'Đính &kèm
        btnImage.Text = rl3("_Hinh_anh") '&Hình ảnh

        btnDetail.Text = rl3("Chi_tiet") 'Chi tiết
        btnSave.Text = rl3("_Luu") '&Lưu
        btnNext.Text = rl3("_Nhap_tiep") '&Nhập tiếp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        chkMaintainable.Text = rL3("Bao_duong") 'Bảo dưỡng
        chkIsTools.Text = rL3("Cong_cu_dung_cu") 'Công cụ dụng cụ
        '================================================================ 
        grpDetail.Text = rl3("Thiet_bi_dinh_kem") 'Thiết bị đính kèm
        grpIndex.Text = rl3("Cac_chi_so") 'Các chỉ số
        '============================================================== 
        tab01.Text = "1. " & rl3("Thong_tin_quan_ly") '1. Thông tin quản lý
        tab02.Text = "2. " & rl3("Thong_tin_tai_chinh") '2. Thông tin tài chính
        tab03.Text = "3. " & rl3("Thong_tin_ky_thuat")  '3. Thông tin kỹ thuật
        tab04.Text = "4. " & rl3("Thong_tin_phan_tich")  '4. Thông tin phân tích
        tab05.Text = "5. " & rL3("Thong_tin_phu")  '4. Thông tin phân tích
        tab06.Text = "6. " & rL3("Cong_cu_dung_cu") '6. Công cụ dụng cụ
        '================================================================ 
        tdbcLocationID.Columns("LocationID").Caption = rL3("Ma") 'Mã
        tdbcLocationID.Columns("LocationName").Caption = rL3("Ten") 'Tên
        tdbcAssetS1ID.Columns("AssetS1ID").Caption = rl3("Ma") 'Mã
        tdbcAssetS1ID.Columns("AssetS1Name").Caption = rl3("Ten") 'Tên
        tdbcAssetS2ID.Columns("AssetS2ID").Caption = rl3("Ma") 'Mã
        tdbcAssetS2ID.Columns("AssetS2Name").Caption = rl3("Ten") 'Tên
        tdbcAssetS3ID.Columns("AssetS3ID").Caption = rl3("Ma") 'Mã
        tdbcAssetS3ID.Columns("AssetS3Name").Caption = rl3("Ten") 'Tên
        tdbcSupplierID.Columns("ObjectID").Caption = rl3("Ma") 'Mã
        tdbcSupplierID.Columns("ObjectName").Caption = rl3("Ten") 'Tên
        tdbcSupplierOTID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã
        tdbcSupplierOTID.Columns("ObjectTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcEmployeeID.Columns("EmployeeID").Caption = rl3("Ma") 'Mã
        tdbcEmployeeID.Columns("EmployeeName").Caption = rl3("Ten") 'Tên
        tdbcObjectID.Columns("ObjectID").Caption = rl3("Ma") 'Mã
        tdbcObjectID.Columns("ObjectName").Caption = rl3("Ten") 'Tên
        tdbcObjectTypeID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã
        tdbcObjectTypeID.Columns("ObjectTypeName").Caption = rL3("Ten") 'Tên
        tdbcAssignmentTypeID.Columns("AssignmentTypeID").Caption = rl3("Ma") 'Mã
        tdbcAssignmentTypeID.Columns("AssignmentTypeName").Caption = rl3("Ten") 'Tên
        tdbcDeprTableID.Columns("DeprTableID").Caption = rl3("Ma") 'Mã
        tdbcDeprTableID.Columns("DeprTableName").Caption = rl3("Ten") 'Tên
        tdbcMethodEndID.Columns("MethodEndID").Caption = rl3("Ma") 'Mã
        tdbcMethodEndID.Columns("MethodEndName").Caption = rl3("Ten") 'Tên
        tdbcMethodID.Columns("MethodID").Caption = rl3("Ma") 'Mã
        tdbcMethodID.Columns("MethodName").Caption = rl3("Ten") 'Tên
        tdbcDepAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbcDepAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        tdbcAssetAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbcAssetAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        tdbcAcode10ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode10ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode09ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode09ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode08ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode08ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode07ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode07ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode06ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode06ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode05ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode05ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode04ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode04ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode03ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode03ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode02ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode02ID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAcode01ID.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcAcode01ID.Columns("Description").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcIGEMethodID.Columns("IGEMethodID").Caption = rL3("Ma") 'Mã
        tdbcIGEMethodID.Columns("IGEMethodName").Caption = rL3("Ten") 'Tên
        tdbcMethodIDCCDC.Columns("MethodID").Caption = rL3("Ma") 'Mã
        tdbcMethodIDCCDC.Columns("MethodName").Caption = rL3("Ten") 'Tên
        tdbcAccountID.Columns("AccountID").Caption = rL3("Ma") 'Mã
        tdbcAccountID.Columns("AccountName").Caption = rL3("Ten") 'Tên
        tdbcReceiverID.Columns("EmployeeID").Caption = rL3("Ma") 'Mã
        tdbcReceiverID.Columns("EmployeeName").Caption = rL3("Ten") 'Tên
        tdbcLocationIDID6.Columns("LocationID").Caption = rL3("Ma") 'Mã
        tdbcLocationIDID6.Columns("LocationName").Caption = rL3("Ten") 'Tên
        tdbcSupplierOTIDID6.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcSupplierOTIDID6.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcSupplierIDID6.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcSupplierIDID6.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        '================================================================ 
        lblManagementObTypeID6.Text = rL3("Bo_phan_quan_ly") 'Bộ phận quản lý
        lblObjectTypeID6.Text = rL3("Bo_phan_tiep_nhan") 'Bộ phận tiếp nhận
        lblSetupVoucherID.Text = rL3("So_phieu_hinh_thanh") 'Số phiếu hình thành
        lblSetupDate.Text = rL3("Ngay_hinh_thanh") 'Ngày hình thành
        '================================================================ 
        lblOQuantity.Text = rL3("So_luong_nguyen") 'Số lượng nguyên
        lblCQuantity.Text = rL3("So_luong_quy_doi") 'Số lượng quy đổi
        '================================================================ 
        tdbcManagementObID6.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcManagementObID6.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        tdbcManagementObTypeID6.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcManagementObTypeID6.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        tdbcObjectID6.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcObjectID6.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        tdbcObjectTypeID6.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcObjectTypeID6.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        '================================================================ 
        tdbdObjectID.Columns("ObjectID").Caption = rl3("Ma") 'Mã
        tdbdObjectID.Columns("ObjectName").Caption = rl3("Ten") 'Tên
        tdbdObjectTypeID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã 
        tdbdObjectTypeID.Columns("ObjectTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        tdbgDetail.Columns("OrderNum").Caption = rl3("STT") 'STT
        tdbgDetail.Columns("EquipmentID").Caption = rl3("Ma_thiet_bi_dinh_kem") 'Mã thiết bị đính kèm
        tdbgDetail.Columns("EquipmentName").Caption = rL3("Ten_thiet_bi_dinh_kem") 'Tên thiết bị đính kèm
        tdbgDetail.Columns("EquipmentQuantity").Caption = rL3("So_luong") 'Số lượng
        tdbgDetail.Columns("UnitPrice").Caption = rL3("Don_gia")       ' Đơn giá
        tdbgDetail.Columns("EquipmentValue").Caption = rL3("Gia_tri_") 'Giá trị
        tdbgDetail.Columns("TaxAmount").Caption = rL3("Tien_thue_GTGT")      ' Tiền thuế GTGT
        tdbgDetail.Columns("AcceptanceTime").Caption = rL3("Thoi_gian_nghiem_thu")  ' Thời gian nghiệm thu
        tdbgDetail.Columns("PurchaseDate").Caption = rL3("Ngay_mua")    ' Ngày mua
        tdbgDetail.Columns("ObjectTypeID").Caption = rL3("Loai_phong_ban") 'Loại phòng ban
        tdbgDetail.Columns("ObjectID").Caption = rL3("Ma_phong_ban") 'Mã phòng ban
        tdbgDetail.Columns("Notes").Caption = rL3("Ghi_chu") 'Ghi chú
        '================================================================ 
        lblObjectTypeID2.Text = rL3("Bo_phan_quan_ly") 'Bộ phận quản lý
        '================================================================ 
        tdbcObjectTypeID2.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcObjectTypeID2.Columns("ObjectTypeName").Caption = rL3("Ten") 'Tên
        tdbcObjectID2.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcObjectID2.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        '================================================================ 
        lblMaintainDate.Text = rL3("Han_bao_hanh") 'Hạn bảo hành
        lblAssetConditionID.Text = rL3("Tinh_trang") 'Tình trạng
        '================================================================ 
        tdbcAssetConditionName.Columns("AssetConditionID").Caption = rL3("Ma") 'Mã
        tdbcAssetConditionName.Columns("Description").Caption = rL3("Ten") 'Tên

        '================================================================ 
        lblChargeObjType.Text = rL3("Bo_phan_chiu_phi") 'Bộ phận chi phí
        '=====================================================  =========== 
        tdbcChargeObjType.Columns("ChargeObjType").Caption = rL3("Ma") 'Mã
        tdbcChargeObjType.Columns("ChargeObjTypeName").Caption = rL3("Ten") 'Tên

    End Sub

    Private Sub tdbgDetail_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbgDetail.ComboSelect
        Select Case e.ColIndex
            Case COL_ObjectTypeID
                tdbgDetail.Columns(COL_ObjectTypeID).Text = tdbdObjectTypeID.Columns("ObjectTypeID").Value.ToString
                tdbgDetail.Columns(COL_ObjectID).Text = ""

            Case COL_ObjectID
                tdbgDetail.Columns(COL_ObjectID).Text = tdbdObjectID.Columns("ObjectID").Value.ToString
        End Select
    End Sub

    Private Sub tdbgDetail_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbgDetail.BeforeColUpdate
        Select Case e.ColIndex
            Case COL_EquipmentID
                e.Cancel = L3IsID(tdbgDetail, e.ColIndex)
            Case COL_EquipmentQuantity
                If Not IsNumeric(tdbgDetail.Columns(COL_EquipmentQuantity).Text) Then e.Cancel = True
            Case COL_EquipmentValue
                If Not IsNumeric(tdbgDetail.Columns(COL_EquipmentValue).Text) Then e.Cancel = True
            Case COL_ObjectTypeID
                If tdbgDetail.Columns(COL_ObjectTypeID).Text <> tdbdObjectTypeID.Columns("ObjectTypeID").Text Then
                    tdbgDetail.Columns(COL_ObjectTypeID).Text = ""
                End If
            Case COL_ObjectID
                If tdbgDetail.Columns(COL_ObjectID).Text <> tdbdObjectID.Columns("ObjectID").Text Then
                    tdbgDetail.Columns(COL_ObjectID).Text = ""
                End If
            Case COL_Notes
        End Select
    End Sub

#Region "Events tdbcAssetAccountID"

    Private Sub tdbcAssetAccountID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetAccountID.Close
        'If tdbcAssetAccountID.FindStringExact(tdbcAssetAccountID.Text) = -1 Then tdbcAssetAccountID.Text = ""
    End Sub

    Private Sub tdbcAssetAccountID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetAccountID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcAssetAccountID.Text = ""
    End Sub

#End Region

#Region "Events tdbcDepAccountID"

    Private Sub tdbcDepAccountID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDepAccountID.Close
        'If tdbcDepAccountID.FindStringExact(tdbcDepAccountID.Text) = -1 Then tdbcDepAccountID.Text = ""
    End Sub

    Private Sub tdbcDepAccountID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcDepAccountID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcDepAccountID.Text = ""
    End Sub

#End Region

#Region "Events tdbcAssignmentTypeID with txtAssignmentTypeName"

    Private Sub tdbcAssignmentTypeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssignmentTypeID.Close
        If tdbcAssignmentTypeID.FindStringExact(tdbcAssignmentTypeID.Text) = -1 Then
            tdbcAssignmentTypeID.Text = ""
            txtAssignmentTypeName.Text = ""
        End If
    End Sub

    Private Sub tdbcAssignmentTypeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssignmentTypeID.SelectedValueChanged
        txtAssignmentTypeName.Text = tdbcAssignmentTypeID.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcAssignmentTypeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssignmentTypeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcAssignmentTypeID.Text = ""
            txtAssignmentTypeName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcDeprTableID with txtDeprTableName"

    Private Sub tdbcDeprTableID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDeprTableID.Close
        If tdbcDeprTableID.FindStringExact(tdbcDeprTableID.Text) = -1 Then
            tdbcDeprTableID.Text = ""
            txtDeprTableName.Text = ""
        End If
    End Sub

    Private Sub tdbcDeprTableID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDeprTableID.SelectedValueChanged
        txtDeprTableName.Text = tdbcDeprTableID.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcDeprTableID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcDeprTableID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcDeprTableID.Text = ""
            txtDeprTableName.Text = ""
        End If
    End Sub

#End Region

    Private Sub tdbgDetail_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbgDetail.KeyDown
        If e.Shift And e.KeyCode = Keys.Insert Then
            HotKeyShiftInsert(tdbgDetail, 0, COL_OrderNum, tdbgDetail.Columns.Count)
        End If
        If e.KeyCode = Keys.F7 Then
            HotKeyF7(tdbgDetail)
        End If
        If e.KeyCode = Keys.F8 Then
            HotKeyF8(tdbgDetail)
        End If
    End Sub

    Private Sub tdbgDetail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbgDetail.KeyUp
        'If e.KeyCode = Keys.Enter Then
        '    HotKeyEnterGrid(tdbgDetail, COL_OrderNum, e)
        'End If
    End Sub

    Private Sub tdbgDetail_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbgDetail.RowColChange
        If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        If tdbgDetail.AddNewMode = C1.Win.C1TrueDBGrid.AddNewModeEnum.AddNewCurrent Then
            tdbgDetail.Columns(COL_EquipmentName).Text = "" ' Gán 1 cột bất kỳ ="" cho lưới
        End If

        '--- Đổ nguồn cho các Dropdown phụ thuộc
        Select Case tdbgDetail.Col
            Case COL_ObjectID
                LoadtdbdObjectID(tdbgDetail.Columns(COL_ObjectTypeID).Text)
        End Select
    End Sub

    Private Sub pnlAsset_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        mAssetID_D02F1031 = txtAssetID.Text
    End Sub

    Private Sub D02F1031_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '12/10/2020, id 144622-Tài sản cố định_Lỗi chưa cảnh báo khi lưu
        If _FormState = EnumFormState.FormEdit Then
            If Not _savedOK Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            If (txtAssetID.Text <> "" Or txtAssetName.Text <> "" Or txtNotes.Text <> "") Then
                If Not _savedOK Then
                    If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub D02F1031_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt Then
            Select Case e.KeyCode

                Case Keys.D1, Keys.NumPad1
                    tab.SelectedTab = tab.TabPages("tab01")
                    Exit Sub
                Case Keys.D2, Keys.NumPad2
                    tab.SelectedTab = tab.TabPages("tab02")
                    Exit Sub
                Case Keys.D3, Keys.NumPad3
                    tab.SelectedTab = tab.TabPages("tab03")
                    Exit Sub
                Case Keys.D4, Keys.NumPad4
                    tab.SelectedTab = tab.TabPages("tab04")
                    Exit Sub
                Case Keys.D5, Keys.NumPad5
                    tab.SelectedTab = tab.TabPages("tab05")
                    Exit Sub
            End Select
        End If
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub



    Private Sub D02F1031_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If bLoadFormState = False Then FormState = _formState
        Loadlanguage()
        SetBackColorObligatory()

        'If _FormState = EnumFormState.FormAdd Then
        '    GetAutoAssetInfo()
        'End If
        FormatNumber()
        InputbyUnicode(Me, gbUnicode)
        CheckIdTextBox(txtAssetID)
        CheckIdTextBox(txtCountryID)
        CheckIdTextBox(txtSeriNo, 20)
        CheckIdTextBox(txtVersion)
        CheckIdTextBox(txtAssetNo)
        'CheckIdTDBGrid(tdbgDetail, COL_EquipmentID, False)
        SetResolutionForm(Me)
        If _FormState = EnumFormState.FormAdd Then
            chkIsTools_CheckedChanged(Nothing, Nothing)
        End If

        

        InputDateInTrueDBGrid(tdbgDetail, COL_AcceptanceTime, COL_PurchaseDate)

        'UseFilterComboObjectID(tdbcObjectID, tdbcObjectTypeID, True) 'Tab 1: Bộ phận tiếp nhận
        'UseFilterComboObjectID(tdbcObjectID2, tdbcObjectTypeID2, True) 'Tab 1: Bộ phận quản lý
        'UseFilterComboObjectID(tdbcSupplierID, tdbcSupplierOTID, True) 'Tab 1: Nhà cung cấp
        'UseFilterComboObjectID(tdbcObjectID6, tdbcObjectTypeID6, True) 'Tab 6: Bộ phận tiếp nhận 
        'UseFilterComboObjectID(tdbcManagementObID6, tdbcManagementObTypeID6, True) 'Tab 6: Bộ phận quản lý
        'UseFilterComboObjectID(tdbcSupplierIDID6, tdbcSupplierOTIDID6, True) 'Tab 6: Nhà cung cấp
        'UseFilterCombo(True, tdbcEmployeeID, tdbcLocationID, tdbcAssetAccountID, tdbcDepAccountID, tdbcAssetConditionName, tdbcAcode01ID, tdbcAcode02ID, tdbcAcode03ID, tdbcAcode04ID, tdbcAcode05ID, tdbcAcode06ID, tdbcAcode07ID, tdbcAcode08ID, tdbcAcode09ID, tdbcAcode10ID, tdbcReceiverID, tdbcLocationIDID6, tdbcUnitID, tdbcAccountID, tdbcMethodIDCCDC)

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub SetBackColorObligatory()
        tdbcAssetS1ID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAssetS2ID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAssetS3ID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtAssetID.BackColor = COLOR_BACKCOLOROBLIGATORY
        txtAssetName.BackColor = COLOR_BACKCOLOROBLIGATORY
        ' tdbcObjectTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        ' tdbcObjectID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        ' tdbcEmployeeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAssetAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDepAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        c1datePeriod.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateDepPeriod.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateTranDate.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcMethodID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcMethodEndID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDeprTableID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAssignmentTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcUnitID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcMethodIDCCDC.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY

        If D02Systems.IsCalDepByDate = True Then c1dateDepDate.BackColor = COLOR_BACKCOLOROBLIGATORY '17/8/2020, Đặng Ngọc Tài:id 142642-SVI_Bổ sung tính năng tạo mức khấu hao theo ngày trong Kỳ đầu tiên module Tài sản cố định
    End Sub



    Private Sub CheckEdit()
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0613()
        If CheckStore(sSQL) = False Then
            grp02.Enabled = False
            tdbcDepAccountID.Enabled = False
        End If
    End Sub

    Private Sub txtIndex1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndex1.Validated
        txtIndex1.Text = SQLNumber(txtIndex1.Text, "N6")
    End Sub

    Private Sub txtIndex2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndex2.Validated
        txtIndex2.Text = SQLNumber(txtIndex2.Text, "N6")
    End Sub

    Private Sub txtIndex3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndex3.Validated
        txtIndex3.Text = SQLNumber(txtIndex3.Text, "N6")
    End Sub

    Private Sub txtIndex4_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndex4.Validated
        txtIndex4.Text = SQLNumber(txtIndex4.Text, "N6")
    End Sub

    Private Sub txtIndex5_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndex5.Validated
        txtIndex5.Text = SQLNumber(txtIndex5.Text, "N6")
    End Sub

    Private Sub txtIndex6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndex6.Validated
        txtIndex6.Text = SQLNumber(txtIndex6.Text, "N6")
    End Sub

    Private Sub D02F1031_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        'If _FormState = EnumFormState.FormAdd Then
        '    'GetAutoAssetInfo()
        '    If bAuto Then
        '        gnNewLastKey = 0
        '        _S1 = IIf(IsDBNull(tdbcAssetS1ID.Text) Or tdbcAssetS1ID.Text = "<<", "", tdbcAssetS1ID.Text).ToString
        '        _S2 = IIf(IsDBNull(tdbcAssetS2ID.Text) Or tdbcAssetS2ID.Text = "<<", "", tdbcAssetS2ID.Text).ToString
        '        _S3 = IIf(IsDBNull(tdbcAssetS3ID.Text) Or tdbcAssetS3ID.Text = "<<", "", tdbcAssetS3ID.Text).ToString
        '        D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _OutputOrder, _OutputLength, _Seperator, txtAssetID, False, _TableName)
        '        ' tdbcAssetS1ID.Focus()
        '        If Not gbCheckLastKey Then Exit Sub
        '        txtAssetName.Focus()

        '    Else
        '        txtAssetID.Text = ""
        '        txtAssetID.Focus()
        '    End If
        '    txtAssetName.Focus()
        'End If
        If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormCopy Then
            If iAssetAuto = 2 Then
                tdbcIGEMethodID.Focus()
            Else
                tdbcAssetS1ID.Focus()
            End If
        ElseIf _FormState = EnumFormState.FormEdit Or _FormState = EnumFormState.FormView Then
            txtAssetName.Focus()
        End If

        'If chkIsTools.Checked Then
        '    EnabledTabPage(New TabPage() {tab06}, True)
        '    tab.SelectedTab = tab06
        'Else
        '    EnabledTabPage(New TabPage() {tab06}, False)
        '    tab.SelectedTab = tab01
        'End If
    End Sub


    Private Sub LoadCaption()
        Dim bUseSpec As Boolean = False
        Dim sSQL As String = SQLStoreD02P0015()
        Dim dt As New DataTable
        dt = ReturnDataTable(sSQL)
        If (dt.Rows.Count > 0) Then
            Str01.Font = FontUnicode(gbUnicode)
            Str02.Font = FontUnicode(gbUnicode)
            Str03.Font = FontUnicode(gbUnicode)
            Str04.Font = FontUnicode(gbUnicode)
            Str05.Font = FontUnicode(gbUnicode)
            Str06.Font = FontUnicode(gbUnicode)
            Str07.Font = FontUnicode(gbUnicode)
            Str08.Font = FontUnicode(gbUnicode)
            Str09.Font = FontUnicode(gbUnicode)
            Str10.Font = FontUnicode(gbUnicode)

            Num01.Font = FontUnicode(gbUnicode)
            Num02.Font = FontUnicode(gbUnicode)
            Num03.Font = FontUnicode(gbUnicode)
            Num04.Font = FontUnicode(gbUnicode)
            Num05.Font = FontUnicode(gbUnicode)
            Num06.Font = FontUnicode(gbUnicode)
            Num07.Font = FontUnicode(gbUnicode)
            Num08.Font = FontUnicode(gbUnicode)
            Num09.Font = FontUnicode(gbUnicode)
            Num10.Font = FontUnicode(gbUnicode)
            Date01.Font = FontUnicode(gbUnicode)
            Date02.Font = FontUnicode(gbUnicode)
            Date03.Font = FontUnicode(gbUnicode)
            Date04.Font = FontUnicode(gbUnicode)
            Date05.Font = FontUnicode(gbUnicode)
            Date06.Font = FontUnicode(gbUnicode)
            Date07.Font = FontUnicode(gbUnicode)
            Date08.Font = FontUnicode(gbUnicode)
            Date09.Font = FontUnicode(gbUnicode)
            Date10.Font = FontUnicode(gbUnicode)

            If geLanguage = EnumLanguage.Vietnamese Then
                Str01.Text = dt.Rows(10).Item("Data84").ToString
            Else
                Str01.Text = dt.Rows(10).Item("Data01").ToString
            End If
            Str01.Tag = dt.Rows(10).Item("DecimalNum")
            txtFAString01.Enabled = L3Bool(dt.Rows(10).Item("Disabled"))
            txtFAString01.MaxLength = L3Int(dt.Rows(10).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str02.Text = dt.Rows(11).Item("Data84").ToString
            Else
                Str02.Text = dt.Rows(11).Item("Data01").ToString
            End If
            txtFAString02.Enabled = L3Bool(dt.Rows(11).Item("Disabled"))
            txtFAString02.MaxLength = L3Int(dt.Rows(11).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str03.Text = dt.Rows(12).Item("Data84").ToString
            Else
                Str03.Text = dt.Rows(12).Item("Data01").ToString
            End If
            txtFAString03.Enabled = L3Bool(dt.Rows(12).Item("Disabled"))
            txtFAString03.MaxLength = L3Int(dt.Rows(12).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str04.Text = dt.Rows(13).Item("Data84").ToString
            Else
                Str04.Text = dt.Rows(13).Item("Data01").ToString
            End If
            txtFAString04.Enabled = L3Bool(dt.Rows(13).Item("Disabled"))
            txtFAString04.MaxLength = L3Int(dt.Rows(13).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str05.Text = dt.Rows(14).Item("Data84").ToString
            Else
                Str05.Text = dt.Rows(14).Item("Data01").ToString
            End If
            txtFAString05.Enabled = L3Bool(dt.Rows(14).Item("Disabled"))
            txtFAString05.MaxLength = L3Int(dt.Rows(14).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str06.Text = dt.Rows(15).Item("Data84").ToString
            Else
                Str06.Text = dt.Rows(15).Item("Data01").ToString
            End If
            txtFAString06.Enabled = L3Bool(dt.Rows(15).Item("Disabled"))
            txtFAString06.MaxLength = L3Int(dt.Rows(15).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str07.Text = dt.Rows(16).Item("Data84").ToString
            Else
                Str07.Text = dt.Rows(16).Item("Data01").ToString
            End If
            txtFAString07.Enabled = L3Bool(dt.Rows(16).Item("Disabled"))
            txtFAString07.MaxLength = L3Int(dt.Rows(16).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str08.Text = dt.Rows(17).Item("Data84").ToString
            Else
                Str08.Text = dt.Rows(17).Item("Data01").ToString
            End If
            txtFAString08.Enabled = L3Bool(dt.Rows(17).Item("Disabled"))
            txtFAString08.MaxLength = L3Int(dt.Rows(17).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str09.Text = dt.Rows(18).Item("Data84").ToString
            Else
                Str09.Text = dt.Rows(18).Item("Data01").ToString
            End If
            txtFAString09.Enabled = L3Bool(dt.Rows(18).Item("Disabled"))
            txtFAString09.MaxLength = L3Int(dt.Rows(18).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str10.Text = dt.Rows(19).Item("Data84").ToString
            Else
                Str10.Text = dt.Rows(19).Item("Data01").ToString
            End If
            txtFAString10.Enabled = L3Bool(dt.Rows(19).Item("Disabled"))
            txtFAString10.MaxLength = L3Int(dt.Rows(19).Item("DecimalNum"))

            If geLanguage = EnumLanguage.Vietnamese Then
                Num01.Text = dt.Rows(0).Item("Data84").ToString
            Else
                Num01.Text = dt.Rows(0).Item("Data01").ToString
            End If
            txtFANum01.Enabled = L3Bool(dt.Rows(0).Item("Disabled"))
            txtFANum01.Tag = dt.Rows(0).Item("DecimalNum")
            If geLanguage = EnumLanguage.Vietnamese Then
                Num02.Text = dt.Rows(1).Item("Data84").ToString
            Else
                Num02.Text = dt.Rows(1).Item("Data01").ToString
            End If
            txtFANum02.Enabled = L3Bool(dt.Rows(1).Item("Disabled"))
            txtFANum02.Tag = dt.Rows(1).Item("DecimalNum")
            If geLanguage = EnumLanguage.Vietnamese Then
                Num03.Text = dt.Rows(2).Item("Data84").ToString
            Else
                Num03.Text = dt.Rows(2).Item("Data01").ToString
            End If
            txtFANum03.Tag = dt.Rows(2).Item("DecimalNum")
            txtFANum03.Enabled = L3Bool(dt.Rows(2).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num04.Text = dt.Rows(3).Item("Data84").ToString
            Else
                Num04.Text = dt.Rows(3).Item("Data01").ToString
            End If
            txtFANum04.Tag = dt.Rows(3).Item("DecimalNum")
            txtFANum04.Enabled = L3Bool(dt.Rows(3).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num05.Text = dt.Rows(4).Item("Data84").ToString
            Else
                Num05.Text = dt.Rows(4).Item("Data01").ToString
            End If
            txtFANum05.Tag = dt.Rows(4).Item("DecimalNum")
            txtFANum05.Enabled = L3Bool(dt.Rows(4).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num06.Text = dt.Rows(5).Item("Data84").ToString
            Else
                Num06.Text = dt.Rows(5).Item("Data01").ToString
            End If
            txtFANum06.Tag = dt.Rows(5).Item("DecimalNum")
            txtFANum06.Enabled = L3Bool(dt.Rows(5).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num07.Text = dt.Rows(6).Item("Data84").ToString
            Else
                Num07.Text = dt.Rows(6).Item("Data01").ToString
            End If
            txtFANum07.Tag = dt.Rows(6).Item("DecimalNum")
            txtFANum07.Enabled = L3Bool(dt.Rows(6).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num08.Text = dt.Rows(7).Item("Data84").ToString
            Else
                Num08.Text = dt.Rows(7).Item("Data01").ToString
            End If
            txtFANum08.Tag = dt.Rows(7).Item("DecimalNum")
            txtFANum08.Enabled = L3Bool(dt.Rows(7).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num09.Text = dt.Rows(8).Item("Data84").ToString
            Else
                Num09.Text = dt.Rows(8).Item("Data01").ToString
            End If
            txtFANum09.Tag = dt.Rows(8).Item("DecimalNum")
            txtFANum09.Enabled = L3Bool(dt.Rows(8).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num10.Text = dt.Rows(9).Item("Data84").ToString
            Else
                Num10.Text = dt.Rows(9).Item("Data01").ToString
            End If
            txtFANum10.Tag = dt.Rows(9).Item("DecimalNum")
            txtFANum10.Enabled = L3Bool(dt.Rows(9).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date01.Text = dt.Rows(20).Item("Data84").ToString
            Else
                Date01.Text = dt.Rows(20).Item("Data01").ToString
            End If
            c1dateFADate01.Enabled = L3Bool(dt.Rows(20).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date02.Text = dt.Rows(21).Item("Data84").ToString
            Else
                Date02.Text = dt.Rows(21).Item("Data01").ToString
            End If
            c1dateFADate02.Enabled = L3Bool(dt.Rows(21).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date03.Text = dt.Rows(22).Item("Data84").ToString
            Else
                Date03.Text = dt.Rows(22).Item("Data01").ToString
            End If
            c1dateFADate03.Enabled = L3Bool(dt.Rows(22).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date04.Text = dt.Rows(23).Item("Data84").ToString
            Else
                Date04.Text = dt.Rows(23).Item("Data01").ToString
            End If
            c1dateFADate04.Enabled = L3Bool(dt.Rows(23).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date05.Text = dt.Rows(24).Item("Data84").ToString
            Else
                Date05.Text = dt.Rows(24).Item("Data01").ToString
            End If
            c1dateFADate05.Enabled = L3Bool(dt.Rows(24).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date06.Text = dt.Rows(25).Item("Data84").ToString
            Else
                Date06.Text = dt.Rows(25).Item("Data01").ToString
            End If
            c1dateFADate06.Enabled = L3Bool(dt.Rows(25).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date07.Text = dt.Rows(26).Item("Data84").ToString
            Else
                Date07.Text = dt.Rows(26).Item("Data01").ToString
            End If
            c1dateFADate07.Enabled = L3Bool(dt.Rows(26).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date08.Text = dt.Rows(27).Item("Data84").ToString
            Else
                Date08.Text = dt.Rows(27).Item("Data01").ToString
            End If
            c1dateFADate08.Enabled = L3Bool(dt.Rows(27).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date09.Text = dt.Rows(28).Item("Data84").ToString
            Else
                Date09.Text = dt.Rows(28).Item("Data01").ToString
            End If
            c1dateFADate09.Enabled = L3Bool(dt.Rows(28).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date10.Text = dt.Rows(29).Item("Data84").ToString
            Else
                Date10.Text = dt.Rows(29).Item("Data01").ToString
            End If
            c1dateFADate10.Enabled = L3Bool(dt.Rows(29).Item("Disabled"))
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0015
    '# Created User: Lê Đình Thái
    '# Created Date: 08/11/2011 03:37:03
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0015() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0015 "
        sSQL &= SQLString("D02T0001") & COMMA 'TableName, nvarchar, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, nvarchar, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function



    Private Sub txtFANum01_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum01.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum02_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum02.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum03_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum03.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum04_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum04.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum05_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum05.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum06_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum06.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum07_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum07.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum08_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum08.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum09_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum09.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum10_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum10.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Function InsertFormat(ByVal ONumber As Object) As String
        Dim iNumber As Int16 = Convert.ToInt16(ONumber)
        Dim sRet As String = "#,##0"
        If iNumber = 0 Then
        Else
            sRet &= "." & Strings.StrDup(iNumber, "0")
        End If
        Return sRet
    End Function
    Private Sub txtFANum01_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum01.LostFocus
        txtFANum01.Text = SQLNumber(IIf(txtFANum01.Text = "", 0, txtFANum01.Text), InsertFormat(txtFANum01.Tag))
    End Sub
    Private Sub txtFANum02_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum02.LostFocus
        txtFANum02.Text = SQLNumber(IIf(txtFANum02.Text = "", 0, txtFANum02.Text), InsertFormat(txtFANum02.Tag))
    End Sub
    Private Sub txtFANum03_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum03.LostFocus
        txtFANum03.Text = SQLNumber(IIf(txtFANum03.Text = "", 0, txtFANum03.Text), InsertFormat(txtFANum03.Tag))
    End Sub
    Private Sub txtFANum04_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum04.LostFocus
        txtFANum04.Text = SQLNumber(IIf(txtFANum04.Text = "", 0, txtFANum04.Text), InsertFormat(txtFANum04.Tag))
    End Sub
    Private Sub txtFANum05_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum05.LostFocus
        txtFANum05.Text = SQLNumber(IIf(txtFANum05.Text = "", 0, txtFANum05.Text), InsertFormat(txtFANum05.Tag))
    End Sub
    Private Sub txtFANum06_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum06.LostFocus
        txtFANum06.Text = SQLNumber(IIf(txtFANum06.Text = "", 0, txtFANum06.Text), InsertFormat(txtFANum06.Tag))
    End Sub
    Private Sub txtFANum07_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum07.LostFocus
        txtFANum07.Text = SQLNumber(IIf(txtFANum07.Text = "", 0, txtFANum07.Text), InsertFormat(txtFANum07.Tag))
    End Sub
    Private Sub txtFANum08_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum08.LostFocus
        txtFANum08.Text = SQLNumber(IIf(txtFANum08.Text = "", 0, txtFANum08.Text), InsertFormat(txtFANum08.Tag))
    End Sub
    Private Sub txtFANum09_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum09.LostFocus
        txtFANum09.Text = SQLNumber(IIf(txtFANum09.Text = "", 0, txtFANum09.Text), InsertFormat(txtFANum09.Tag))
    End Sub
    Private Sub txtFANum10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum10.LostFocus
        txtFANum10.Text = SQLNumber(IIf(txtFANum10.Text = "", 0, txtFANum10.Text), InsertFormat(txtFANum10.Tag))
    End Sub

    Private Sub txtFAString10_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString10.KeyPress
        If (txtFAString10.MaxLength = 0) Then
            e.Handled = True

        End If
    End Sub

    Private Sub txtFAString9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString09.KeyPress
        If (txtFAString09.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString01.KeyPress
        If (txtFAString01.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString02.KeyPress
        If (txtFAString02.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString03.KeyPress
        If (txtFAString03.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString04.KeyPress
        If (txtFAString04.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString05.KeyPress
        If (txtFAString05.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString06.KeyPress
        If (txtFAString06.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString07.KeyPress
        If (txtFAString07.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString08.KeyPress
        If (txtFAString08.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1031
    '# Created User: Hoàng Nhân
    '# Created Date: 18/07/2013 03:41:38
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1031() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Ke thua ma TS vao JCode" & vbCrlf)
        sSQL &= "Exec D02P1031 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(txtAssetID.Text) 'AssetID, varchar[20], NOT NULL
        Return sSQL
    End Function

#Region "Events tdbcLocationID with txtLocationName"

    Private Sub tdbcLocationID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLocationID.SelectedValueChanged
        If tdbcLocationID.SelectedValue Is Nothing Then
            txtLocationName.Text = ""
        Else
            txtLocationName.Text = tdbcLocationID.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcLocationID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLocationID.LostFocus
        If tdbcLocationID.FindStringExact(tdbcLocationID.Text) = -1 Then
            tdbcLocationID.Text = ""
        End If
    End Sub

#End Region

    Private Sub tdbcName_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLocationID.Close
        tdbcName_Validated(sender, Nothing)
    End Sub

    Private Sub tdbcName_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLocationID.Validated
        Dim tdbc As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)

        FilterCombo(tdbc, e, True)
        tdbc.Text = tdbc.WillChangeToText
    End Sub

#Region "Events tdbcIGEMethodID with txtAssetID"

    Dim sKeyString As String = ""
    Private Sub tdbcIGEMethodID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcIGEMethodID.SelectedValueChanged
        If _bFormD02F0087 = True Then Exit Sub '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu

        Dim sSQL As String = ""
        sSQL &= SQLDeleteD91T1000() & vbCrLf
        sSQL &= SQLInsertD91T1000().ToString & vbCrLf
        sSQL &= SQLStoreD91P1000() & vbCrLf
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            sLastKey = dt.Rows(0).Item("LastKey").ToString
            sKeyString = dt.Rows(0).Item("KeyString").ToString
            _S1 = sKeyString
            If dt.Rows(0).Item("Status").ToString = "1" Then
                D99C0008.MsgL3(ConvertVietwareFToUnicode(dt.Rows(0).Item("Message").ToString))
            Else
                txtAssetID.Text = dt.Rows(0).Item("ID").ToString
            End If
        End If
        sSQL = SQLDeleteD91T1000() & vbCrLf
        ExecuteSQL(sSQL)
    End Sub

    Private Sub tdbcIGEMethodID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcIGEMethodID.LostFocus
        If tdbcIGEMethodID.FindStringExact(tdbcIGEMethodID.Text) = -1 Then
            tdbcIGEMethodID.Text = ""
        End If
    End Sub

#End Region

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P1000
    '# Created User: HUỲNH KHANH
    '# Created Date: 02/10/2014 01:17:35
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD91P1000() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Sinh Ma cap nhat" & vbCrlf)
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
        sSQL &= SQLNumber("20") & COMMA 'Length, tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'IsD07F0011, tinyint, NOT NULL
        sSQL &= SQLNumber(0) 'NewLastKey, int, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD91T1000
    '# Created User: HUỲNH KHANH
    '# Created Date: 02/10/2014 01:18:22
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD91T1000() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Xoa du lieu" & vbCrlf)
        sSQL &= "Delete From D91T1000"
        sSQL &= " Where "
        sSQL &= "UserID = " & SQLString(gsUserID) & " And "
        sSQL &= "HostID = " & SQLString(My.Computer.Name) & " And "
        sSQL &= "FormID = " & SQLString("D02F0070") & " And "
        sSQL &= "ModuleID = " & SQLString("02")
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T1000
    '# Created User: HUỲNH KHANH
    '# Created Date: 02/10/2014 01:18:57
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T1000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Insert du lieu " & vbCrlf)
        sSQL.Append("Insert Into D91T1000(")
        sSQL.Append("UserID, HostID, ModuleID, FormID")
        sSQL.Append(") Values(" & vbCrlf)
        sSQL.Append(SQLString(gsUserID) & COMMA) 'UserID, varchar[50], NOT NULL
        sSQL.Append(SQLString(My.Computer.Name) & COMMA) 'HostID, varchar[50], NOT NULL
        sSQL.Append(SQLString("02") & COMMA) 'ModuleID, varchar[50], NOT NULL
        sSQL.Append(SQLString("D02F0070")) 'FormID, varchar[50], NOT NULL   
        sSQL.Append(")")

        Return sSQL
    End Function

    Private Sub chkIsTools_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIsTools.CheckedChanged
        If D02Systems.IsAssetIDForD02D43 Then
            If chkIsTools.Checked Then
                EnabledTabPage(New TabPage() {tab01, tab02}, False)
                EnabledTabPage(New TabPage() {tab06, tab05}, True)
                tab.SelectedTab = tab06
                btnConvertedAmount.Enabled = False
                btnDepreciate.Enabled = False
                txtUnitName.Visible = False
                lblUnitName.Visible = False
            Else
                EnabledTabPage(New TabPage() {tab01, tab02, tab05}, True)
                EnabledTabPage(New TabPage() {tab06}, False)
                tab.SelectedTab = tab01
                btnConvertedAmount.Enabled = True
                btnDepreciate.Enabled = True
                txtUnitName.Visible = True
                lblUnitName.Visible = True
            End If
            LoadTDBGrid()
        Else
            EnabledTabPage(New TabPage() {tab06}, False)
        End If

    End Sub


#Region "Events tdbcSupplierOTID load tdbcSupplierID with txtSupplierName"
    Private Sub tdbcSupplierOTID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTID.GotFocus
        'Dùng phím Enter
        tdbcSupplierOTID.Tag = tdbcSupplierOTID.Text
    End Sub

    Private Sub tdbcSupplierOTID_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcSupplierOTID.MouseDown
        'Di chuyển chuột
        tdbcSupplierOTID.Tag = tdbcSupplierOTID.Text
    End Sub

    Private Sub tdbcSupplierOTID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTID.SelectedValueChanged
        clsFilterCombo.LoadtdbcObjectID(tdbcSupplierID, dtObjectID, ReturnValueC1Combo(tdbcSupplierOTID))
        tdbcSupplierID.Text = ""
        VisibleColumnsObjectID(tdbcSupplierID, ReturnValueC1Combo(tdbcSupplierOTID))
    End Sub

    Private Sub tdbcSupplierOTID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTID.LostFocus
        'If tdbcSupplierOTID.FindStringExact(tdbcSupplierOTID.Text) = -1 OrElse tdbcSupplierOTID.SelectedValue Is Nothing Then
        '    tdbcSupplierOTID.Text = ""
        '    LoadtdbcObjectID(tdbcSupplierID, "-1")
        '    tdbcSupplierID.Text = ""
        '    Exit Sub
        'End If
        'LoadtdbcObjectID(tdbcSupplierID, tdbcSupplierOTID.SelectedValue.ToString())
        'tdbcSupplierID.Text = ""

        ' ''''''''''''''''''''

        If tdbcSupplierOTID.Tag Is Nothing Then Exit Sub
        If tdbcSupplierOTID.Tag.ToString = tdbcSupplierOTID.Text Then
            If clsFilterCombo.IsNewFilter And tdbcSupplierOTID.FindStringExact(tdbcSupplierOTID.Text) = -1 Then
                clsFilterCombo.LoadtdbcObjectID(tdbcSupplierID, dtObjectID, "-1")
                VisibleColumnsObjectID(tdbcSupplierID, "-1")
            End If
            Exit Sub
        End If
        If tdbcSupplierOTID.FindStringExact(tdbcSupplierOTID.Text) = -1 Then
            tdbcSupplierOTID.Text = ""
            clsFilterCombo.LoadtdbcObjectID(tdbcSupplierID, dtObjectID, "-1")
            VisibleColumnsObjectID(tdbcSupplierID, "-1")
            tdbcSupplierID.Text = ""
            Exit Sub
        End If
        tdbcSupplierID.Text = ""
    End Sub

    Private Sub tdbcSupplierID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierID.SelectedValueChanged
        If tdbcSupplierID.SelectedValue Is Nothing Then
            txtSupplierName.Text = ""
        Else
            txtSupplierName.Text = tdbcSupplierID.Columns("ObjectName").Value.ToString
        End If
    End Sub

    Private Sub tdbcSupplierID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierID.Validated
        clsFilterCombo.FilterCombo(tdbcSupplierID, e)
        If tdbcSupplierID.FindStringExact(tdbcSupplierID.Text) = -1 Then
            tdbcSupplierID.Text = ""
            txtSupplierName.Text = ""
        End If
    End Sub

    Private Sub tdbcSupplierID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSupplierID.KeyDown
        'If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
        '    tdbcSupplierID.Text = ""
        '    txtSupplierName.Text = ""
        'End If
        If clsFilterCombo.IsNewFilter Then
            Exit Sub ' TH filter dạng mới thì F2 gọi D99F5555 đã có sẵn
        End If
        If e.KeyCode = Keys.F2 Then
            Dim sKeyID As String = HotKeyF2("2", " ObjectTypeID = " & SQLString(tdbcSupplierOTID.Text))
            If sKeyID <> "" Then
                tdbcSupplierID.SelectedValue = sKeyID
                tdbcSupplierID.Focus()
            End If
        End If
    End Sub

#End Region

#Region "Events tdbcObjectTypeID load tdbcObjectID with txtObjectName"
    Private Sub tdbcObjectTypeID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.GotFocus
        'Dùng phím Enter
        tdbcObjectTypeID.Tag = tdbcObjectTypeID.Text
    End Sub

    Private Sub tdbcObjectTypeID_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcObjectTypeID.MouseDown
        'Di chuyển chuột
        tdbcObjectTypeID.Tag = tdbcObjectTypeID.Text
    End Sub

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.SelectedValueChanged
        clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
        tdbcObjectID.Text = ""
        VisibleColumnsObjectID(tdbcObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
    End Sub

    Private Sub tdbcObjectTypeID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.LostFocus
        'If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 OrElse tdbcObjectTypeID.SelectedValue Is Nothing Then
        '    tdbcObjectTypeID.Text = ""
        '    LoadtdbcObjectID(tdbcObjectID, "-1")
        '    tdbcObjectID.Text = ""
        '    Exit Sub
        'End If
        ''ID 78424 13/08/2015
        'clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID))
        'tdbcObjectID.SelectedValue = "%"

        ' ''''''''''''''''''''
        '5/9/2018, id 112182-Lỗi mất thông tin phòng ban và đơn vị khi truy vấn mã CCDC tại D02
        If tdbcObjectTypeID.Tag Is Nothing Then Exit Sub
        If tdbcObjectTypeID.Tag.ToString = tdbcObjectTypeID.Text Then
            If clsFilterCombo.IsNewFilter And tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 Then
                clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, "-1")
                VisibleColumnsObjectID(tdbcObjectID, "-1")
            End If
            Exit Sub
        End If
        If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 Then
            tdbcObjectTypeID.Text = ""
            clsFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObjectID, "-1")
            VisibleColumnsObjectID(tdbcObjectID, "-1")
            tdbcObjectID.Text = ""
            Exit Sub
        End If
        tdbcObjectID.Text = ""
    End Sub

    Private Sub VisibleColumnsObjectID(tdbc As C1.Win.C1List.C1Combo, sObjectTypeID As String)
        tdbc.Splits(0).DisplayColumns("ObjectTypeID").Visible = (sObjectTypeID = "" Or sObjectTypeID = "-1") And (clsFilterCombo.IsNewFilter)
    End Sub

    Private Sub tdbcObjectID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID.SelectedValueChanged
        If tdbcObjectID.SelectedValue Is Nothing Then
            txtObjectName.Text = ""
        Else
            txtObjectName.Text = tdbcObjectID.Columns("ObjectName").Value.ToString
        End If
    End Sub

    'ID 78424 12/08/2015
    Private Sub tdbcObjectID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID.Validated
        clsFilterCombo.FilterCombo(tdbcObjectID, e)
        If tdbcObjectID.FindStringExact(tdbcObjectID.Text) = -1 Then
            tdbcObjectID.Text = ""
            txtObjectName.Text = ""
        End If
    End Sub

    Private Sub tdbcObjectID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcObjectID.KeyDown
        If clsFilterCombo.IsNewFilter Then Exit Sub ' TH filter dạng mới thì F2 gọi D99F5555 đã có sẵn

        If e.KeyCode = Keys.F2 Then
            Dim sKeyID As String = HotKeyF2("2", " ObjectTypeID = " & SQLString(tdbcObjectTypeID.Text))
            If sKeyID <> "" Then
                tdbcObjectID.SelectedValue = sKeyID
                tdbcObjectID.Focus()
            End If
        End If
    End Sub

#End Region

#Region "Events tdbcObjectTypeID2 load tdbcObjectID2 with txtObjectName2"

    Private Sub tdbcObjectTypeID2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID2.GotFocus
        'Dùng phím Enter
        tdbcObjectTypeID2.Tag = tdbcObjectTypeID2.Text
    End Sub

    Private Sub tdbcObjectTypeID2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcObjectTypeID2.MouseDown
        'Di chuyển chuột
        tdbcObjectTypeID2.Tag = tdbcObjectTypeID2.Text
    End Sub

    Private Sub tdbcObjectTypeID2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID2.SelectedValueChanged
        clsFilterCombo.LoadtdbcObjectID(tdbcObjectID2, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID2))
        tdbcObjectID2.Text = ""
        VisibleColumnsObjectID(tdbcObjectID2, ReturnValueC1Combo(tdbcObjectTypeID2))
    End Sub

    Private Sub tdbcObjectTypeID2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID2.LostFocus
        'If tdbcObjectTypeID2.FindStringExact(tdbcObjectTypeID2.Text) = -1 OrElse tdbcObjectTypeID2.SelectedValue Is Nothing Then
        '    tdbcObjectTypeID2.Text = ""
        '    LoadtdbcObjectID(tdbcObjectID2, "-1")
        '    tdbcObjectID2.Text = ""
        '    Exit Sub
        'End If
        ''ID 78424 13/08/2015
        'clsFilterCombo.LoadtdbcObjectID(tdbcObjectID2, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID2))
        'tdbcObjectID2.SelectedValue = "%"

        ' ''''''''''''''''''''
        '5/9/2018, id 112182-Lỗi mất thông tin phòng ban và đơn vị khi truy vấn mã CCDC tại D02
        If tdbcObjectTypeID2.Tag Is Nothing Then Exit Sub
        If tdbcObjectTypeID2.Tag.ToString = tdbcObjectTypeID2.Text Then
            If clsFilterCombo.IsNewFilter And tdbcObjectTypeID2.FindStringExact(tdbcObjectTypeID2.Text) = -1 Then
                clsFilterCombo.LoadtdbcObjectID(tdbcObjectID2, dtObjectID, "-1")
                VisibleColumnsObjectID(tdbcObjectID2, "-1")
            End If
            Exit Sub
        End If
        If tdbcObjectTypeID2.FindStringExact(tdbcObjectTypeID2.Text) = -1 Then
            tdbcObjectTypeID2.Text = ""
            clsFilterCombo.LoadtdbcObjectID(tdbcObjectID2, dtObjectID, "-1")
            VisibleColumnsObjectID(tdbcObjectID2, "-1")
            tdbcObjectID2.Text = ""
            Exit Sub
        End If
        tdbcObjectID2.Text = ""
    End Sub

    Private Sub tdbcObjectID2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID2.SelectedValueChanged
        If tdbcObjectID2.SelectedValue Is Nothing Then
            txtObjectName2.Text = ""
        Else
            txtObjectName2.Text = tdbcObjectID2.Columns("ObjectName").Value.ToString
        End If
    End Sub

    Private Sub tdbcObjectID2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID2.Validated
        clsFilterCombo.FilterCombo(tdbcObjectID2, e)
        If tdbcObjectID2.FindStringExact(tdbcObjectID2.Text) = -1 Then
            tdbcObjectID2.Text = ""
            txtObjectName2.Text = ""
        End If
    End Sub

    Private Sub tdbcObjectID2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcObjectID2.KeyDown
        'If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
        '    tdbcObjectID.Text = ""
        '    txtObjectName.Text = ""
        'End If
        If clsFilterCombo.IsNewFilter Then
            Exit Sub ' TH filter dạng mới thì F2 gọi D99F5555 đã có sẵn
        End If

        If e.KeyCode = Keys.F2 Then
            Dim sKeyID As String = HotKeyF2("2", " ObjectTypeID = " & SQLString(tdbcObjectTypeID2.Text))
            If sKeyID <> "" Then
                tdbcObjectID2.SelectedValue = sKeyID
                tdbcObjectID2.Focus()
            End If
        End If
    End Sub

#End Region

#Region "Events tdbcAssetConditionID"

    Private Sub tdbcAssetConditionID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetConditionName.LostFocus
        'If tdbcAssetConditionName.FindStringExact(tdbcAssetConditionName.Text) = -1 Then tdbcAssetConditionName.Text = ""
    End Sub

#End Region


    'Private Sub LoadtdbcObjectID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sObjectTypeID As String)
    '    LoadDataSource(tdbc, ReturnTableFilter(dtObjectID, "ObjectTypeID = " & SQLString(sObjectTypeID), True), gbUnicode)
    'End Sub

    '    Private Sub LoadtdbcObjectID2(ByVal ID As String)
    '        LoadDataSource(tdbcObjectID2, ReturnTableFilter(dtObjectID, "ObjectTypeID = " & SQLString(ID), True), gbUnicode)
    '    End Sub
    '
    '    Private Sub LoadtdbcSupplierID(ByVal ID As String)
    '        LoadDataSource(tdbcSupplierID, ReturnTableFilter(dtObjectID, "ObjectTypeID = " & SQLString(ID), True), gbUnicode)
    '    End Sub

    Private Sub LoadtdbdObjectID(ByVal ID As String)
        LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObjectID, "ObjectTypeID=" & SQLString(ID), True), gbUnicode)
    End Sub

#Region "Events tdbcObjectTypeID6 load tdbcObjectID6 with txtObjectName6"
    Private Sub tdbcObjectTypeID6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID6.GotFocus
        'Dùng phím Enter
        tdbcObjectTypeID6.Tag = tdbcObjectTypeID6.Text
    End Sub

    Private Sub tdbcObjectTypeID6_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcObjectTypeID6.MouseDown
        'Di chuyển chuột
        tdbcObjectTypeID6.Tag = tdbcObjectTypeID6.Text
    End Sub

    Private Sub tdbcObjectTypeID6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID6.SelectedValueChanged
        'LoadtdbcObjectID(tdbcObjectID6, ReturnValueC1Combo(tdbcObjectTypeID6))
        clsFilterCombo.LoadtdbcObjectID(tdbcObjectID6, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID6))
        tdbcObjectID6.Text = ""
        VisibleColumnsObjectID(tdbcObjectID6, ReturnValueC1Combo(tdbcObjectTypeID6))
    End Sub

    Private Sub tdbcObjectTypeID6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID6.LostFocus
        'If tdbcObjectTypeID6.FindStringExact(tdbcObjectTypeID6.Text) = -1 OrElse tdbcObjectTypeID6.SelectedValue Is Nothing Then
        '    tdbcObjectTypeID6.Text = ""
        '    LoadtdbcObjectID(tdbcObjectID6, "-1")
        '    tdbcObjectID6.Text = ""
        '    Exit Sub
        'End If
        ''ID 78424 13/08/2015
        'clsFilterCombo.LoadtdbcObjectID(tdbcObjectID6, dtObjectID, ReturnValueC1Combo(tdbcObjectTypeID6))
        'tdbcObjectID.SelectedValue = "%"

        ' ''''''''''''''''''''
        '5/9/2018, id 112182-Lỗi mất thông tin phòng ban và đơn vị khi truy vấn mã CCDC tại D02
        If tdbcObjectTypeID6.Tag Is Nothing Then Exit Sub
        If tdbcObjectTypeID6.Tag.ToString = tdbcObjectTypeID6.Text Then
            If clsFilterCombo.IsNewFilter And tdbcObjectTypeID6.FindStringExact(tdbcObjectTypeID6.Text) = -1 Then
                clsFilterCombo.LoadtdbcObjectID(tdbcObjectID6, dtObjectID, "-1")
                VisibleColumnsObjectID(tdbcObjectID6, "-1")
            End If
            Exit Sub
        End If
        If tdbcObjectTypeID6.FindStringExact(tdbcObjectTypeID6.Text) = -1 Then
            tdbcObjectTypeID6.Text = ""
            clsFilterCombo.LoadtdbcObjectID(tdbcObjectID6, dtObjectID, "-1")
            VisibleColumnsObjectID(tdbcObjectID6, "-1")
            tdbcObjectID6.Text = ""
            Exit Sub
        End If
        tdbcObjectID6.Text = ""
    End Sub

    Private Sub tdbcObjectID6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID6.SelectedValueChanged
        If tdbcObjectID6.SelectedValue Is Nothing Then
            txtObjectName6.Text = ""
        Else
            txtObjectName6.Text = tdbcObjectID6.Columns("ObjectName").Value.ToString
        End If
    End Sub

    Private Sub tdbcObjectID6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID6.Validated
        clsFilterCombo.FilterCombo(tdbcObjectID6, e)
        If tdbcObjectID6.FindStringExact(tdbcObjectID6.Text) = -1 Then
            tdbcObjectID6.Text = ""
            txtObjectName6.Text = ""
        End If
    End Sub
#End Region


#Region "Events tdbcManagementObTypeID6 load tdbcManagementObID6 with txtManagementObName"
    Private Sub tdbcManagementObTypeID6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObTypeID6.GotFocus
        'Dùng phím Enter
        tdbcManagementObTypeID6.Tag = tdbcManagementObTypeID6.Text
    End Sub

    Private Sub tdbcManagementObTypeID6_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcManagementObTypeID6.MouseDown
        'Di chuyển chuột
        tdbcManagementObTypeID6.Tag = tdbcManagementObTypeID6.Text
    End Sub

    Private Sub tdbcManagementObTypeID6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObTypeID6.SelectedValueChanged
        clsFilterCombo.LoadtdbcObjectID(tdbcManagementObID6, dtObjectID, ReturnValueC1Combo(tdbcManagementObTypeID6))
        tdbcManagementObID6.Text = ""
        VisibleColumnsObjectID(tdbcManagementObID6, ReturnValueC1Combo(tdbcManagementObTypeID6))
    End Sub

    Private Sub tdbcManagementObTypeID6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObTypeID6.LostFocus
        'If tdbcManagementObTypeID6.FindStringExact(tdbcManagementObTypeID6.Text) = -1 OrElse tdbcManagementObTypeID6.SelectedValue Is Nothing Then
        '    tdbcManagementObTypeID6.Text = ""
        '    LoadtdbcObjectID(tdbcManagementObID6, "-1")
        '    tdbcManagementObID6.Text = ""
        '    Exit Sub
        'End If
        ''ID 78424 13/08/2015
        'clsFilterCombo.LoadtdbcObjectID(tdbcManagementObID6, dtObjectID, ReturnValueC1Combo(tdbcManagementObTypeID6))
        'tdbcObjectID.SelectedValue = "%"

        ' ''''''''''''''''''''
        '5/9/2018, id 112182-Lỗi mất thông tin phòng ban và đơn vị khi truy vấn mã CCDC tại D02
        If tdbcManagementObTypeID6.Tag Is Nothing Then Exit Sub
        If tdbcManagementObTypeID6.Tag.ToString = tdbcManagementObTypeID6.Text Then
            If clsFilterCombo.IsNewFilter And tdbcManagementObTypeID6.FindStringExact(tdbcManagementObTypeID6.Text) = -1 Then
                clsFilterCombo.LoadtdbcObjectID(tdbcManagementObID6, dtObjectID, "-1")
                VisibleColumnsObjectID(tdbcManagementObID6, "-1")
            End If
            Exit Sub
        End If
        If tdbcManagementObTypeID6.FindStringExact(tdbcManagementObTypeID6.Text) = -1 Then
            tdbcManagementObTypeID6.Text = ""
            clsFilterCombo.LoadtdbcObjectID(tdbcManagementObID6, dtObjectID, "-1")
            VisibleColumnsObjectID(tdbcManagementObID6, "-1")
            tdbcManagementObID6.Text = ""
            Exit Sub
        End If
        tdbcManagementObID6.Text = ""
    End Sub

    Private Sub tdbcManagementObID6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObID6.SelectedValueChanged
        If tdbcManagementObID6.SelectedValue Is Nothing Then
            txtManagementObName.Text = ""
        Else
            txtManagementObName.Text = tdbcManagementObID6.Columns("ObjectName").Value.ToString
        End If
    End Sub

    Private Sub tdbcManagementObID6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObID6.Validated
        clsFilterCombo.FilterCombo(tdbcManagementObID6, e)
        If tdbcManagementObID6.FindStringExact(tdbcManagementObID6.Text) = -1 Then
            tdbcManagementObID6.Text = ""
            txtManagementObName.Text = ""
        End If
    End Sub

#End Region

    Private Function GetTool() As String
        Dim sTool As String = ""
        For i As Integer = 0 To tdbgDetail.RowCount - 1
            Dim sValue As String = L3String(tdbgDetail(i, COL_EquipmentID))
            If sValue = "" Then Continue For
            If sTool <> "" Then sTool &= ","
            sTool &= sValue
        Next
        Return sTool

    End Function

#Region "Events tdbcReceiverID with txtReceiverID"

    Private Sub tdbcReceiverID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcReceiverID.Close
        'If tdbcReceiverID.FindStringExact(tdbcReceiverID.Text) = -1 Then
        '    tdbcReceiverID.Text = ""
        '    txtReceiverName.Text = ""
        'End If
    End Sub

    Private Sub tdbcReceiverID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcReceiverID.SelectedValueChanged

        If tdbcReceiverID.SelectedValue Is Nothing Then
            txtReceiverName.Text = ""
        Else
            txtReceiverName.Text = tdbcReceiverID.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcReceiverID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcReceiverID.KeyDown
        ''Dim sKeyID As String
        'If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
        '    tdbcReceiverID.Text = ""
        '    txtReceiverName.Text = ""
        'End If
        'If e.KeyCode = Keys.F2 Then
        '    Dim arrPro() As StructureProperties = Nothing
        '    SetProperties(arrPro, "InListID", "2")
        '    SetProperties(arrPro, "InWhere", " ObjectTypeID ='NV' ")
        '    Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6010", arrPro)
        '    Dim sKey As String = GetProperties(frm, "Output01").ToString
        '    If sKey <> "" Then
        '        tdbcReceiverID.SelectedValue = sKey
        '        tdbcReceiverID.Focus()
        '    End If
        'End If
    End Sub

#End Region
#Region "Events tdbcLocationIDID6 with tdbcLocationNameID6"

    Private Sub tdbcLocationIDID6_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLocationIDID6.SelectedValueChanged
        If tdbcLocationIDID6.SelectedValue Is Nothing Then
            txtLocationNameID6.Text = ""
        Else
            txtLocationNameID6.Text = tdbcLocationIDID6.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcLocationIDID6_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLocationIDID6.LostFocus
        'If tdbcLocationIDID6.FindStringExact(tdbcLocationIDID6.Text) = -1 Then
        '    tdbcLocationIDID6.Text = ""
        '    txtLocationNameID6.Text = ""
        'End If
    End Sub

#End Region

#Region "Events tdbcSupplierOTIDID6 load tdbcSupplierIDID6 with txtSupplierNameID6"
    Private Sub tdbcSupplierOTIDID6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTIDID6.GotFocus
        'Dùng phím Enter
        tdbcSupplierOTIDID6.Tag = tdbcSupplierOTIDID6.Text
    End Sub

    Private Sub tdbcSupplierOTIDID6_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcSupplierOTIDID6.MouseDown
        'Di chuyển chuột
        tdbcSupplierOTIDID6.Tag = tdbcSupplierOTIDID6.Text
    End Sub

    Private Sub tdbcSupplierOTIDID6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTIDID6.SelectedValueChanged
        clsFilterCombo.LoadtdbcObjectID(tdbcSupplierIDID6, dtObjectID, ReturnValueC1Combo(tdbcSupplierOTIDID6))
        tdbcSupplierIDID6.Text = ""
        VisibleColumnsObjectID(tdbcSupplierIDID6, ReturnValueC1Combo(tdbcSupplierOTIDID6))
    End Sub

    Private Sub tdbcSupplierOTIDID6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTIDID6.Validated
        'clsFilterCombo.FilterCombo(tdbcSupplierOTIDID6, e)
        'If tdbcSupplierOTIDID6.FindStringExact(tdbcSupplierOTIDID6.Text) = -1 OrElse tdbcSupplierOTIDID6.SelectedValue Is Nothing Then
        '    tdbcSupplierOTIDID6.Text = ""
        '    LoadtdbcObjectID(tdbcSupplierIDID6, "-1")
        '    tdbcSupplierIDID6.Text = ""
        '    txtSupplierNameID6.Text = ""
        '    Exit Sub
        'End If
        ''ID 78424 13/08/2015
        'clsFilterCombo.LoadtdbcObjectID(tdbcSupplierIDID6, dtObjectID, ReturnValueC1Combo(tdbcSupplierOTIDID6))
        'tdbcObjectID.SelectedValue = "%"
        'txtSupplierNameID6.Text = ""

        ' ''''''''''''''''''''
        '5/9/2018, id 112182-Lỗi mất thông tin phòng ban và đơn vị khi truy vấn mã CCDC tại D02
        If tdbcSupplierOTIDID6.Tag Is Nothing Then Exit Sub
        If tdbcSupplierOTIDID6.Tag.ToString = tdbcSupplierOTIDID6.Text Then
            If clsFilterCombo.IsNewFilter And tdbcSupplierOTIDID6.FindStringExact(tdbcSupplierOTIDID6.Text) = -1 Then
                clsFilterCombo.LoadtdbcObjectID(tdbcSupplierIDID6, dtObjectID, "-1")
                VisibleColumnsObjectID(tdbcSupplierIDID6, "-1")
            End If
            Exit Sub
        End If
        If tdbcSupplierOTIDID6.FindStringExact(tdbcSupplierOTIDID6.Text) = -1 Then
            tdbcSupplierOTIDID6.Text = ""
            clsFilterCombo.LoadtdbcObjectID(tdbcSupplierIDID6, dtObjectID, "-1")
            VisibleColumnsObjectID(tdbcSupplierIDID6, "-1")
            tdbcSupplierIDID6.Text = ""
            Exit Sub
        End If
        tdbcSupplierIDID6.Text = ""
    End Sub

    Private Sub tdbcSupplierIDID6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierIDID6.SelectedValueChanged
        If tdbcSupplierIDID6.SelectedValue Is Nothing Then
            txtSupplierNameID6.Text = ""
        Else
            txtSupplierNameID6.Text = tdbcSupplierIDID6.Columns("ObjectName").Value.ToString
        End If
    End Sub

    Private Sub tdbcSupplierIDID6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierIDID6.LostFocus
        If tdbcSupplierIDID6.FindStringExact(tdbcSupplierIDID6.Text) = -1 Then
            tdbcSupplierIDID6.Text = ""
            txtSupplierNameID6.Text = ""
        End If

    End Sub

#End Region

    Private Sub cneOQuantity_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cneOQuantity.LostFocus
        If cneOQuantity.Value.ToString <> "" Then
            txtCQuantity.Text = SQLNumber(cneOQuantity.Value, DxxFormat.D07_QuantityDecimals)
        Else
            txtCQuantity.Text = ""
        End If

    End Sub

    Private Sub tdbcEmployeeID_Validated(ByVal sender As Object, ByVal e As EventArgs) Handles tdbcEmployeeID.Validated
        clsFilterCombo.FilterCombo(tdbcEmployeeID, e)

        If tdbcEmployeeID.FindStringExact(tdbcEmployeeID.Text) = -1 Then
            tdbcEmployeeID.Text = ""
            txtEmployeeName.Text = ""
        End If
    End Sub

    Private Sub tdbcAssetAccountID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetAccountID.Validated
        clsFilterCombo.FilterCombo(tdbcAssetAccountID, e)
        If tdbcAssetAccountID.FindStringExact(tdbcAssetAccountID.Text) = -1 Then tdbcAssetAccountID.Text = ""
    End Sub

    Private Sub tdbcDepAccountID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcDepAccountID.Validated
        clsFilterCombo.FilterCombo(tdbcDepAccountID, e)
        If tdbcDepAccountID.FindStringExact(tdbcDepAccountID.Text) = -1 Then tdbcDepAccountID.Text = ""
    End Sub

    Private Sub tdbcAssetConditionName_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAssetConditionName.Validated
        clsFilterCombo.FilterCombo(tdbcAssetConditionName, e)
        If tdbcAssetConditionName.FindStringExact(tdbcAssetConditionName.Text) = -1 Then tdbcAssetConditionName.Text = ""
    End Sub

    Private Sub tdbcAcode01ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode01ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode01ID, e)
        If tdbcAcode01ID.FindStringExact(tdbcAcode01ID.Text) = -1 Then tdbcAcode01ID.Text = ""
        AddNewACode(tdbcAcode01ID, tdbcAcode01ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode02ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode02ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode02ID, e)
        If tdbcAcode02ID.FindStringExact(tdbcAcode02ID.Text) = -1 Then tdbcAcode02ID.Text = ""
        AddNewACode(tdbcAcode02ID, tdbcAcode02ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode03ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode03ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode03ID, e)
        If tdbcAcode03ID.FindStringExact(tdbcAcode03ID.Text) = -1 Then tdbcAcode03ID.Text = ""
        AddNewACode(tdbcAcode03ID, tdbcAcode03ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode04ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode04ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode04ID, e)
        If tdbcAcode04ID.FindStringExact(tdbcAcode04ID.Text) = -1 Then tdbcAcode04ID.Text = ""
        AddNewACode(tdbcAcode04ID, tdbcAcode04ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode05ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode05ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode05ID, e)
        If tdbcAcode05ID.FindStringExact(tdbcAcode05ID.Text) = -1 Then tdbcAcode05ID.Text = ""
        AddNewACode(tdbcAcode05ID, tdbcAcode05ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode06ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode06ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode06ID, e)
        If tdbcAcode06ID.FindStringExact(tdbcAcode06ID.Text) = -1 Then tdbcAcode06ID.Text = ""
        AddNewACode(tdbcAcode06ID, tdbcAcode06ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode07ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode07ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode07ID, e)
        If tdbcAcode07ID.FindStringExact(tdbcAcode07ID.Text) = -1 Then tdbcAcode07ID.Text = ""
        AddNewACode(tdbcAcode07ID, tdbcAcode07ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode08ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode08ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode08ID, e)
        If tdbcAcode08ID.FindStringExact(tdbcAcode08ID.Text) = -1 Then tdbcAcode08ID.Text = ""
        AddNewACode(tdbcAcode08ID, tdbcAcode08ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode09ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode09ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode09ID, e)
        If tdbcAcode09ID.FindStringExact(tdbcAcode09ID.Text) = -1 Then tdbcAcode09ID.Text = ""
        AddNewACode(tdbcAcode09ID, tdbcAcode09ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcAcode10ID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAcode10ID.Validated
        clsFilterCombo.FilterCombo(tdbcAcode10ID, e)
        If tdbcAcode10ID.FindStringExact(tdbcAcode10ID.Text) = -1 Then tdbcAcode10ID.Text = ""
        AddNewACode(tdbcAcode10ID, tdbcAcode10ID.Text) '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    End Sub

    Private Sub tdbcReceiverID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcReceiverID.Validated
        clsFilterCombo.FilterCombo(tdbcReceiverID, e)
        If tdbcReceiverID.FindStringExact(tdbcReceiverID.Text) = -1 Then
            tdbcReceiverID.Text = ""
            txtReceiverName.Text = ""
        End If
    End Sub

    Private Sub tdbcLocationIDID6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLocationIDID6.Validated
        clsFilterCombo.FilterCombo(tdbcLocationIDID6, e)
        If tdbcLocationIDID6.FindStringExact(tdbcLocationIDID6.Text) = -1 Then
            tdbcLocationIDID6.Text = ""
            txtLocationNameID6.Text = ""
        End If
    End Sub

    Private Sub tdbcUnitID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcUnitID.Validated
        clsFilterCombo.FilterCombo(tdbcUnitID, e)
    End Sub

    Private Sub tdbcAccountID_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcAccountID.Validated
        clsFilterCombo.FilterCombo(tdbcAccountID, e)
    End Sub

    Private Sub tdbcMethodIDCCDC_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcMethodIDCCDC.Validated
        clsFilterCombo.FilterCombo(tdbcMethodIDCCDC, e)
    End Sub

    Private Sub tdbcAssetConditionName_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetConditionName.KeyDown
        If clsFilterCombo.IsNewFilter Then
            Exit Sub ' TH filter dạng mới thì F2 gọi D99F5555 đã có sẵn
        End If
    End Sub

    Private Sub tdbcLocationID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcLocationID.KeyDown
        If clsFilterCombo.IsNewFilter Then
            Exit Sub ' TH filter dạng mới thì F2 gọi D99F5555 đã có sẵn
        End If
    End Sub

    Private Sub AddNewACode(tdbcACode As C1.Win.C1List.C1Combo, sACodeID As String)
        '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
        If sACodeID = "+" Then
            If ReturnPermission("D02F0041") < 2 Then
                D99C0008.MsgL3(rL3("Ban_khong_co_quyen_them_moi"), L3MessageBoxIcon.Information)
                tdbcACode.Text = ""
                Exit Sub
            End If

            Dim bSaved As Boolean = False
            Dim sKey As String = ""
            Dim frm As New D02F0042
            With frm
                .FormName = "D02F1031"
                .TypeCodeID ="A" + tdbcACode.Name.Substring(9, 2)
                .ShowDialog()
                bSaved = .SavedOK
                sKey = .ACodeID
            End With

            If bSaved = True Then
                If sKey.ToString <> "" Then
                    ReLoadTDBCACode(tdbcACode)
                    tdbcACode.Text = sKey.ToString
                Else
                    tdbcACode.Text = ""
                End If
            Else
                tdbcACode.Text = ""
            End If
           
        End If
    End Sub

    Private Sub LoadTDBCACodeID()
        '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
        Dim sSQL As String = ""
        Dim dt As DataTable

        sSQL &= "Select " & NewCode & " As ACodeID, " & NewName & " As Description, '' As TypeCodeID, 0 As DisplayOrder " & vbCrLf
        sSQL &= "Union All " & vbCrLf
        sSQL &= "Select ACodeID, Description" & sUnicode & " As Description, TypeCodeID, 1 As DisplayOrder " & vbCrLf
        sSQL &= "From D02T0041 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where Disabled=0 " & vbCrLf
        sSQL &= "Order by DisplayOrder, TypeCodeID, AcodeID" & vbCrLf
        dt = ReturnDataTable(sSQL)

        LoadDataSource(tdbcAcode01ID, ReturnTableFilter(dt, "TypeCodeID = 'A01' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode02ID, ReturnTableFilter(dt, "TypeCodeID = 'A02' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode03ID, ReturnTableFilter(dt, "TypeCodeID = 'A03' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode04ID, ReturnTableFilter(dt, "TypeCodeID = 'A04' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode05ID, ReturnTableFilter(dt, "TypeCodeID = 'A05' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode06ID, ReturnTableFilter(dt, "TypeCodeID = 'A06' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode07ID, ReturnTableFilter(dt, "TypeCodeID = 'A07' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode08ID, ReturnTableFilter(dt, "TypeCodeID = 'A08' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode09ID, ReturnTableFilter(dt, "TypeCodeID = 'A09' or ACodeID = '+'"), gbUnicode)
        LoadDataSource(tdbcAcode10ID, ReturnTableFilter(dt, "TypeCodeID = 'A10' or ACodeID = '+'"), gbUnicode)
    End Sub

    Private Sub ReLoadTDBCACode(tdbcACode As C1.Win.C1List.C1Combo)
        '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
        Dim sSQL As String = ""
        Dim dt As DataTable

        sSQL &= "Select " & NewCode & " As ACodeID, " & NewName & " As Description, '' As TypeCodeID, 0 As DisplayOrder " & vbCrLf
        sSQL &= "Union All " & vbCrLf
        sSQL &= "Select ACodeID, Description" & sUnicode & " As Description, TypeCodeID, 1 As DisplayOrder " & vbCrLf
        sSQL &= "From D02T0041 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where Disabled = 0 And TypeCodeID = " & SQLString("A" + tdbcACode.Name.Substring(9, 2)) & vbCrLf & vbCrLf
        sSQL &= "Order by DisplayOrder, AcodeID" & vbCrLf
        dt = ReturnDataTable(sSQL)

        LoadDataSource(tdbcACode, dt, gbUnicode)
    End Sub

    'Private Sub tdbcAcode01ID_SelectedValueChanged(sender As Object, e As EventArgs) Handles tdbcAcode01ID.SelectedValueChanged, tdbcAcode02ID.SelectedValueChanged, tdbcAcode03ID.SelectedValueChanged,
    '    tdbcAcode04ID.SelectedValueChanged, tdbcAcode05ID.SelectedValueChanged, tdbcAcode06ID.SelectedValueChanged, tdbcAcode07ID.SelectedValueChanged,
    '    tdbcAcode08ID.SelectedValueChanged, tdbcAcode09ID.SelectedValueChanged, tdbcAcode10ID.SelectedValueChanged
    '    '1/12/2021, Phạm Thị Thu:id 204655-Thêm tính năng thêm mới mã phân tích khi tạo mới tài sản cố định
    '    Dim tdbcAcodeID As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)

    '    If tdbcAcodeID.SelectedValue Is Nothing Then Exit Sub
    '    AddNewACode(tdbcAcodeID, tdbcAcodeID.Text)
    'End Sub


End Class