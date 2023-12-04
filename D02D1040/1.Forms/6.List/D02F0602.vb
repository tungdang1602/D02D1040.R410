Public Class D02F0602

#Region "Const of tdbgConvertedAmount"
    Private Const COLMODE1_TransactionTypeID As Integer = 0 ' Loại bút toán
    Private Const COLMODE1_VoucherTypeID As Integer = 1     ' Loại phiếu
    Private Const COLMODE1_VoucherNo As Integer = 2         ' Số phiếu
    Private Const COLMODE1_VoucherDate As Integer = 3       ' Ngày phiếu
    Private Const COLMODE1_SerialNo As Integer = 4          ' Số Sêri
    Private Const COLMODE1_RefNo As Integer = 5             ' Số hóa đơn
    Private Const COLMODE1_RefDate As Integer = 6           ' Ngày hóa đơn
    Private Const COLMODE1_ObjectTypeID As Integer = 7      ' Mã loại đối tượng
    Private Const COLMODE1_ObjectID As Integer = 8          ' Mã đối tượng
    Private Const COLMODE1_Description As Integer = 9       ' Diễn giải
    Private Const COLMODE1_DebitAccountID As Integer = 10   ' TK nợ
    Private Const COLMODE1_CreditAccountID As Integer = 11  ' TK Có
    Private Const COLMODE1_CurrencyID As Integer = 12       ' Loại tiền
    Private Const COLMODE1_ExchangeRate As Integer = 13     ' Tỷ giá
    Private Const COLMODE1_OriginalAmount As Integer = 14   ' Nguyên tệ
    Private Const COLMODE1_ConvertedAmount As Integer = 15  ' Quy đổi
    Private Const COLMODE1_SourceID As Integer = 16         ' Nguồn hình thành
    Private Const COLMODE1_Cost As Integer = 17             ' Nguyên giá TSCĐ
#End Region

#Region "Const of tdbgDepreciation"
    Private Const COLMODE2_Period As Integer = 0              ' Kỳ
    Private Const COLMODE2_VoucherTypeID As Integer = 1       ' Loại phiếu
    Private Const COLMODE2_VoucherNo As Integer = 2           ' Số phiếu
    Private Const COLMODE2_VoucherDate As Integer = 3         ' Ngày phiếu
    Private Const COLMODE2_AssignmentID As Integer = 4        ' Tiêu thức phân bổ
    Private Const COLMODE2_ObjectTypeID As Integer = 5        ' Mã loại đối tượng
    Private Const COLMODE2_ObjectID As Integer = 6            ' Mã đối tượng
    Private Const COLMODE2_Description As Integer = 7         ' Diễn giải
    Private Const COLMODE2_DebitAccountID As Integer = 8      ' TK nợ
    Private Const COLMODE2_CreditAccountID As Integer = 9     ' TK Có
    Private Const COLMODE2_ConvertedAmount As Integer = 10    ' Giá trị khấu hao
    Private Const COLMODE2_SourceID As Integer = 11           ' Nguồn hình thành
    Private Const COLMODE2_AmountDepreciation As Integer = 12 ' Khấu hao lũy kế
#End Region

#Region "Const of tdbg01 - Total of Columns: 7"
    Private Const COLMODE3_FromPeriod As Integer = 0     ' Từ kỳ
    Private Const COLMODE3_ToPeriod As Integer = 1       ' Đến kỳ
    Private Const COLMODE3_ChangeDate As Integer = 2     ' Ngày tác động
    Private Const COLMODE3_AssignmentID As Integer = 3   ' Tiêu thức phân bổ
    Private Const COLMODE3_AssignmentName As Integer = 4 ' Tên tiêu thức phân bổ
    Private Const COLMODE3_Percentage As Integer = 5     ' Tỷ lệ
    Private Const COLMODE3_CreateUserID As Integer = 6   ' Người nhập
#End Region

#Region "Const of tdbg04 - Total of Columns: 6"
    Private Const COL04_FromPeriod As Integer = 0     ' Từ kỳ
    Private Const COL04_ToPeriod As Integer = 1       ' Đến kỳ
    Private Const COL04_ChangeDate As Integer = 2     ' Ngày tác động
    Private Const COL04_AssetAccountID As Integer = 3 ' Tài khoản tài sản
    Private Const COL04_DepAccountID As Integer = 4   ' Tài khoản khấu hao
    Private Const COL04_CreateUserID As Integer = 5   ' Người nhập
#End Region

#Region "Property"
    Private mMode As Int16
    Public WriteOnly Property Mode() As Int16
        Set(ByVal value As Int16)
            mMode = value
        End Set
    End Property

    Private mAssetID As String
    Public WriteOnly Property AssetID() As String
        Set(ByVal value As String)
            mAssetID = value
        End Set
    End Property

    Private mAssetName As String
    Public WriteOnly Property AssetName() As String
        Set(ByVal value As String)
            mAssetName = value
        End Set
    End Property
#End Region
    
    Private Sub btnDepreciation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDepreciation.Click
        VisibleControlsDepreciation(True)
        LoadMode2()
        VisibleControlsConvertedAmount(False)
        VisibleControlsHistory(False)
    End Sub

    Public Sub VisibleControlsConvertedAmount(ByVal bVisible As Boolean)
        tdbgConvertedAmount.Visible = bVisible
        lblCost.Visible = bVisible
        txtCost.Visible = bVisible
    End Sub

    Public Sub VisibleControlsDepreciation(ByVal bVisible As Boolean)
        tdbgDepreciation.Visible = bVisible
        lblAmountDepreciation.Visible = bVisible
        txtAmountDepreciation.Visible = bVisible
    End Sub

    Public Sub VisibleControlsHistory(ByVal bVisible As Boolean)
        tab.Visible = bVisible
    End Sub

    Private Sub MakeSameSizeAndLocation()
        tdbgDepreciation.Location = tdbgConvertedAmount.Location
        tdbgDepreciation.Size = tdbgConvertedAmount.Size
        lblAmountDepreciation.Location = lblCost.Location
        txtAmountDepreciation.Location = txtCost.Location
        tab.Size = tdbgConvertedAmount.Size
        tab.Location = tdbgConvertedAmount.Location
        tdbg01.Size = tdbgConvertedAmount.Size
        tdbg02.Size = tdbgConvertedAmount.Size
        tdbg03.Size = tdbgConvertedAmount.Size
        tdbg04.Size = tdbgConvertedAmount.Size '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
    End Sub

    Private Sub btnConvertedAmount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConvertedAmount.Click
        VisibleControlsConvertedAmount(True)
        LoadMode1()
        VisibleControlsDepreciation(False)
        VisibleControlsHistory(False)
    End Sub

    'Private Sub tdbgConvertedAmount_NumberFormat()
    '    tdbgConvertedAmount.Columns(COLMODE1_ExchangeRate).NumberFormat = DxxFormat.ExchangeRateDecimals
    '    tdbgConvertedAmount.Columns(COLMODE1_OriginalAmount).NumberFormat = DxxFormat.DecimalPlaces
    '    tdbgConvertedAmount.Columns(COLMODE1_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    'End Sub

    Private Sub tdbgConvertedAmount_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbgConvertedAmount.Columns(COLMODE1_ExchangeRate).DataField, DxxFormat.ExchangeRateDecimals, 28, 8)
        AddDecimalColumns(arr, tdbgConvertedAmount.Columns(COLMODE1_OriginalAmount).DataField, DxxFormat.DecimalPlaces, 28, 8)
        AddDecimalColumns(arr, tdbgConvertedAmount.Columns(COLMODE1_ConvertedAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbgConvertedAmount, arr)
    End Sub

    Private Sub tdbgDepreciation_NumberFormat()
        tdbgDepreciation.Columns(COLMODE2_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    End Sub

    Private Sub tdbg01_NumberFormat()
        tdbg01.Columns(COLMODE3_Percentage).NumberFormat = DxxFormat.DefaultNumber2
    End Sub

    Private Sub D02F0602_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F0602_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        Loadlanguage()
        MakeSameSizeAndLocation()
        tdbgConvertedAmount_NumberFormat()
        tdbgDepreciation_NumberFormat()
        tdbg01_NumberFormat()
        LoadInfo()
        Select Case mMode
            Case 1
                LoadMode1()
            Case 2
                LoadMode2()
            Case 3
                LoadMode3()
        End Select
        InputbyUnicode(Me, gbUnicode)
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub LoadInfo()
        txtDivisionID.Text = gsDivisionID
        txtAssetID.Text = mAssetID
        txtAssetName.Text = mAssetName
        txtDivisionID.Enabled = False
        txtAssetID.Enabled = False
        txtAssetName.Enabled = False
    End Sub

    Private Sub LoadFooterMode1()
        tdbgConvertedAmount.Columns(COLMODE1_ConvertedAmount).FooterText = Format(Sum(COLMODE1_ConvertedAmount, tdbgConvertedAmount), DxxFormat.D90_ConvertedDecimals)
    End Sub

    Private Sub LoadMode1()
        mMode = 1
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0603()
        LoadDataSource(tdbgConvertedAmount, sSQL, gbUnicode)
        LoadFooterMode1()
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            txtCost.Text = dt.Rows(0)("Cost").ToString()
        Else
            txtCost.Text = "0"
        End If
        txtCost.Text = Format(CDbl(txtCost.Text), DxxFormat.D07_UnitCostDecimals)
        txtCost.TextAlign = HorizontalAlignment.Right
    End Sub

    Private Sub LoadFooterMode2()
        tdbgDepreciation.Columns(COLMODE2_ConvertedAmount).FooterText = Format(Sum(COLMODE2_ConvertedAmount, tdbgDepreciation), DxxFormat.D90_ConvertedDecimals)
    End Sub

    Private Sub LoadMode2()
        mMode = 2
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0603()
        LoadDataSource(tdbgDepreciation, sSQL, gbUnicode)
        LoadFooterMode2()
    End Sub

    Private Sub LoadMode3()
        mMode = 3
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0603()
        LoadDataSource(tdbg01, sSQL, gbUnicode)
    End Sub

    Private Sub LoadMode4()
        mMode = 4
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0603()
        LoadDataSource(tdbg02, sSQL, gbUnicode)
    End Sub

    Private Sub LoadMode5()
        mMode = 5
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0603()
        LoadDataSource(tdbg03, sSQL, gbUnicode)
    End Sub

    Private Sub LoadMode6()
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        mMode = 6
        Dim sSQL As String = ""
        sSQL = SQLStoreD02P0603()
        LoadDataSource(tdbg04, sSQL, gbUnicode)
    End Sub

    Private Sub btnEffectHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEffectHistory.Click
        VisibleControlsHistory(True)
        LoadMode3()
        VisibleControlsConvertedAmount(False)
        VisibleControlsDepreciation(False)
    End Sub

    Private Sub btnAssetInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAssetInfo.Click
        Me.Close()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0603
    '# Create User: Hoàng Đức Thịnh
    '# Create Date: 04/08/2006 02:28:30
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0603() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0603 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, VarChar[20], NOT NULL
        sSQL &= SQLString(mAssetID) & COMMA 'AssetID, VarChar[20], NOT NULL
        sSQL &= SQLNumber(mMode) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function


    Private Sub tab_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tab.SelectedIndexChanged
        Select Case tab.SelectedIndex
            Case 0
                LoadMode3()
            Case 1
                LoadMode4()
            Case 2
                LoadMode5()
            Case 3 '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
                LoadMode6()
        End Select
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Truy_van_thong_tin_tai_san_-_D02F0602") & UnicodeCaption(gbUnicode) 'Truy vÊn th¤ng tin tªi s¶n - D02F0602
        '================================================================ 
        lblDivisionID.Text = rL3("Ma_don_vi") 'Mã đơn vị
        lblAssetID.Text = rL3("Ma_TSCD") 'Mã TSCĐ
        lblCost.Text = rL3("Nguyen_gia_TSCD") 'Nguyên giá TSCĐ
        lblAmountDepreciation.Text = rL3("Khau_hao_luy_ke") 'Khấu hao lũy kế
        '================================================================ 
        btnAssetInfo.Text = rL3("Thong_tin_tai_san") 'Thông tin tài sản
        btnConvertedAmount.Text = rL3("Nguyen_gia") 'Nguyên giá
        btnDepreciation.Text = rL3("Khau_hao") 'Khấu hao
        btnEffectHistory.Text = rL3("Lich_su_tac_dong") 'Lịch sử tác động
        '================================================================ 
        tab01.Text = rL3("Tieu_thuc_phan_bo") 'Tiếu thức phân bổ
        tab02.Text = rL3("Bo_phan_tiep_nhan_quan_ly") 'Bộ phận tiếp nhận/ quản lý
        tab03.Text = rL3("Thoi_gian_khau_hao") 'Thời gian khấu hao
        tab04.Text = rL3("Tai_khoan_tai_sankhau_hao") 'Tài khoản tài sản/khấu hao
        '================================================================ 
        tdbgConvertedAmount.Columns("TransactionTypeID").Caption = rL3("Loai_but_toan") 'Loại bút toán
        tdbgConvertedAmount.Columns("VoucherTypeID").Caption = rL3("Loai_phieu") 'Loại phiếu
        tdbgConvertedAmount.Columns("VoucherNo").Caption = rL3("So_phieu") 'Số phiếu
        tdbgConvertedAmount.Columns("VoucherDate").Caption = rL3("Ngay_phieu") 'Ngày phiếu
        tdbgConvertedAmount.Columns("RefNo").Caption = rL3("So_hoa_don") 'Số hóa đơn
        tdbgConvertedAmount.Columns("SerialNo").Caption = rL3("So_Seri") 'Số Sêri
        tdbgConvertedAmount.Columns("RefDate").Caption = rL3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbgConvertedAmount.Columns("ObjectTypeID").Caption = rL3("Ma_loai_doi_tuong") 'Mã loại đối tượng
        tdbgConvertedAmount.Columns("ObjectID").Caption = rL3("Ma_doi_tuong") 'Mã đối tượng
        tdbgConvertedAmount.Columns("Description").Caption = rL3("Dien_giai") 'Diễn giải
        tdbgConvertedAmount.Columns("DebitAccountID").Caption = rL3("TK_no") 'TK nợ
        tdbgConvertedAmount.Columns("CreditAccountID").Caption = rL3("TK_co") 'TK Có
        tdbgConvertedAmount.Columns("CurrencyID").Caption = rL3("Loai_tien") 'Loại tiền
        tdbgConvertedAmount.Columns("ExchangeRate").Caption = rL3("Ty_gia") 'Tỷ giá
        tdbgConvertedAmount.Columns("OriginalAmount").Caption = rL3("Nguyen_te") 'Nguyên tệ
        tdbgConvertedAmount.Columns("ConvertedAmount").Caption = rL3("Quy_doi") 'Quy đổi
        tdbgConvertedAmount.Columns("SourceID").Caption = rL3("Nguon_hinh_thanh") 'Nguồn hình thành
        tdbgDepreciation.Columns("Period").Caption = rL3("Ky") 'Kỳ
        tdbgDepreciation.Columns("VoucherTypeID").Caption = rL3("Loai_phieu") 'Loại phiếu
        tdbgDepreciation.Columns("VoucherNo").Caption = rL3("So_phieu") 'Số phiếu
        tdbgDepreciation.Columns("VoucherDate").Caption = rL3("Ngay_phieu") 'Ngày phiếu
        tdbgDepreciation.Columns("AssignmentID").Caption = rL3("Tieu_thuc_phan_bo") 'Tiêu thức phân bổ
        tdbgDepreciation.Columns("ObjectTypeID").Caption = rL3("Ma_loai_doi_tuong") 'Mã loại đối tượng
        tdbgDepreciation.Columns("ObjectID").Caption = rL3("Ma_doi_tuong") 'Mã đối tượng
        tdbgDepreciation.Columns("Description").Caption = rL3("Dien_giai") 'Diễn giải
        tdbgDepreciation.Columns("DebitAccountID").Caption = rL3("TK_no") 'TK nợ
        tdbgDepreciation.Columns("CreditAccountID").Caption = rL3("TK_co") 'TK Có
        tdbgDepreciation.Columns("ConvertedAmount").Caption = rL3("Gia_tri_khau_hao") 'Giá trị khấu hao
        tdbgDepreciation.Columns("SourceID").Caption = rL3("Nguon_hinh_thanh") 'Nguồn hình thành
        tdbgConvertedAmount.Columns(COLMODE1_VoucherNo).FooterText = rL3("Tong_cong")  'Tổng cộng
        tdbgDepreciation.Columns(COLMODE2_VoucherNo).FooterText = rL3("Tong_cong")  'Tổng cộng

        '================================================================ 
        tdbg01.Columns("ChangeDate").Caption = rL3("Ngay_tac_dong") 'Ngày tác động
        tdbg01.Columns("FromPeriod").Caption = rL3("Tu_ky") 'Từ kỳ
        tdbg01.Columns("ToPeriod").Caption = rL3("Den_ky") 'Đến kỳ
        tdbg01.Columns("AssignmentID").Caption = rL3("Tieu_thuc_phan_bo") 'Tiêu thức phân bổ
        tdbg01.Columns("AssignmentName").Caption = rL3("Ten_tieu_thuc_phan_bo") 'Tên tiêu thức phân bổ
        tdbg01.Columns("Percentage").Caption = rL3("Ty_le") 'Tỷ lệ
        tdbg01.Columns("CreateUserID").Caption = rL3("Nguoi_nhap") 'Người nhập
        tdbg02.Columns("TypeName").Caption = rL3("Loai") 'Loại
        tdbg02.Columns("FromPeriod").Caption = rL3("Tu_ky") 'Từ kỳ
        tdbg02.Columns("ChangeDate").Caption = rL3("Ngay_tac_dong") 'Ngày tác động
        tdbg02.Columns("ToPeriod").Caption = rL3("Den_ky") 'Đến kỳ
        tdbg02.Columns("ObjectTypeID").Caption = rL3("Loai_doi_tuong") 'Loại đối tượng
        tdbg02.Columns("ObjectID").Caption = rL3("Doi_tuong") 'Đối tượng
        tdbg02.Columns("ObjectName").Caption = rL3("Ten_doi_tuong") 'Tên đối tượng
        tdbg02.Columns("CreateUserID").Caption = rL3("Nguoi_nhap") 'Người nhập
        tdbg03.Columns("FromPeriod").Caption = rL3("Tu_ky") 'Từ kỳ
        tdbg03.Columns("ChangeDate").Caption = rL3("Ngay_tac_dong") 'Ngày tác động
        tdbg03.Columns("ToPeriod").Caption = rL3("Den_ky") 'Đến kỳ
        tdbg03.Columns("ServiceLife").Caption = rL3("So_ky_khau_hao") 'Số kỳ khấu hao
        tdbg03.Columns("CreateUserID").Caption = rL3("Nguoi_nhap") 'Người nhập
        '================================================================ 
        tdbg04.Columns(COL04_FromPeriod).Caption = rL3("Tu_ky") 'Từ kỳ
        tdbg04.Columns(COL04_ToPeriod).Caption = rL3("Den_ky") 'Đến kỳ
        tdbg04.Columns(COL04_ChangeDate).Caption = rL3("Ngay_tac_dong") 'Ngày tác động
        tdbg04.Columns(COL04_AssetAccountID).Caption = rL3("Tai_khoan_tai_san") 'Tài khoản tài sản
        tdbg04.Columns(COL04_DepAccountID).Caption = rL3("Tai_khoan_khau_hao") 'Tài khoản khấu hao
        tdbg04.Columns(COL04_CreateUserID).Caption = rL3("Nguoi_nhap") 'Người nhập

       

    End Sub
End Class