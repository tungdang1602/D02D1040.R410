Imports System.Drawing
Imports System
Public Class D02F1030
	Dim report As D99C2003
	Private _formIDPermission As String = "D02F1030"
	Public WriteOnly Property FormIDPermission() As String
		Set(ByVal Value As String)
			       _formIDPermission = Value
		   End Set
	End Property


    '#Region "Const of tdbg - Total of Columns: 85"
    '    Private Const COL_IsPrinted As Integer = 0            ' Chọn
    '    Private Const COL_AssetID As Integer = 1              ' Mã tài sản
    '    Private Const COL_AssetName As Integer = 2            ' Tên tài sản
    '    Private Const COL_Notes As Integer = 3                ' Ghi chú
    '    Private Const COL_D54ProjectID As Integer = 4         ' Mã dự án
    '    Private Const COL_D27PropertyProductID As Integer = 5 ' Mã BĐS
    '    Private Const COL_FANum01 As Integer = 6              ' Num01
    '    Private Const COL_FANum02 As Integer = 7              ' Num02
    '    Private Const COL_FANum03 As Integer = 8              ' Num03
    '    Private Const COL_FANum04 As Integer = 9              ' Num04
    '    Private Const COL_FANum05 As Integer = 10             ' Num05
    '    Private Const COL_FANum06 As Integer = 11             ' Num06
    '    Private Const COL_FANum07 As Integer = 12             ' Num07
    '    Private Const COL_FANum08 As Integer = 13             ' Num08
    '    Private Const COL_FANum09 As Integer = 14             ' Num09
    '    Private Const COL_FANum10 As Integer = 15             ' Num10
    '    Private Const COL_FAString01 As Integer = 16          ' Str01
    '    Private Const COL_FAString02 As Integer = 17          ' Str02
    '    Private Const COL_FAString03 As Integer = 18          ' Str03
    '    Private Const COL_FAString04 As Integer = 19          ' Str04
    '    Private Const COL_FAString05 As Integer = 20          ' Str05
    '    Private Const COL_FAString06 As Integer = 21          ' Str06
    '    Private Const COL_FAString07 As Integer = 22          ' Str07
    '    Private Const COL_FAString08 As Integer = 23          ' Str08
    '    Private Const COL_FAString09 As Integer = 24          ' Str09
    '    Private Const COL_FAString10 As Integer = 25          ' Str10
    '    Private Const COL_FADate01 As Integer = 26            ' Date01
    '    Private Const COL_FADate02 As Integer = 27            ' Date02
    '    Private Const COL_FADate03 As Integer = 28            ' Date03
    '    Private Const COL_FADate04 As Integer = 29            ' Date04
    '    Private Const COL_FADate05 As Integer = 30            ' Date05
    '    Private Const COL_FADate06 As Integer = 31            ' Date06
    '    Private Const COL_FADate07 As Integer = 32            ' Date07
    '    Private Const COL_FADate08 As Integer = 33            ' Date08
    '    Private Const COL_FADate09 As Integer = 34            ' Date09
    '    Private Const COL_FADate10 As Integer = 35            ' Date10
    '    Private Const COL_ShortName As Integer = 36           ' Tên tắt
    '    Private Const COL_AssetTag As Integer = 37            ' Thẻ tài sản
    '    Private Const COL_ObjectID As Integer = 38            ' Mã bộ phận tiếp nhận
    '    Private Const COL_ObjectName As Integer = 39          ' Tên bộ phận tiếp nhận
    '    Private Const COL_AssetUserID As Integer = 40         ' Mã người tiếp nhận
    '    Private Const COL_FullName As Integer = 41            ' Tên người tiếp nhận
    '    Private Const COL_UsedDate As Integer = 42            ' Ngày sử dụng
    '    Private Const COL_LocationID As Integer = 43          ' Mã vị trí
    '    Private Const COL_NewLocationName As Integer = 44     ' Tên vị trí
    '    Private Const COL_ConvertedAmount As Integer = 45     ' Nguyên giá
    '    Private Const COL_NotDEPCurrentCost As Integer = 46   ' Giá trị đất
    '    Private Const COL_DEPCurrentCost As Integer = 47      ' Giá trị xây dựng
    '    Private Const COL_DepreciatedAmount As Integer = 48   ' Định mức khấu hao
    '    Private Const COL_AmountDepreciation As Integer = 49  ' Hao mòn lũy kế
    '    Private Const COL_RemainAmount As Integer = 50        ' Giá trị còn lại
    '    Private Const COL_AssetAccountID As Integer = 51      ' TK tài sản
    '    Private Const COL_DepAccountID As Integer = 52        ' TK khấu hao
    '    Private Const COL_ServiceLife As Integer = 53         ' Số kỳ khấu hao
    '    Private Const COL_NewServiceLife As Integer = 54      ' Số kỳ khấu hao gốc
    '    Private Const COL_DepreciatedPeriod As Integer = 55   ' Số kỳ đã khấu hao
    '    Private Const COL_AssetPeriod As Integer = 56         ' Kỳ nhập tài sản
    '    Private Const COL_ACode01ID As Integer = 57           ' Maõ phaân tích taøi saûn 1(Maõ)
    '    Private Const COL_ACode01Name As Integer = 58         ' Maõ phaân tích taøi saûn 1(Teân)
    '    Private Const COL_ACode02ID As Integer = 59           ' Maõ phaân tích taøi saûn 2(Maõ)
    '    Private Const COL_ACode02Name As Integer = 60         ' Maõ phaân tích taøi saûn 2(Teân)
    '    Private Const COL_ACode03ID As Integer = 61           ' Maõ phaân tích taøi saûn 3(Maõ)
    '    Private Const COL_ACode03Name As Integer = 62         ' Maõ phaân tích taøi saûn 3(Teân)
    '    Private Const COL_ACode04ID As Integer = 63           ' Maõ phaân tích taøi saûn 4(Maõ)
    '    Private Const COL_ACode04Name As Integer = 64         ' Maõ phaân tích taøi saûn 4(Teân)
    '    Private Const COL_ACode05ID As Integer = 65           ' Maõ phaân tích taøi saûn 5(Maõ)
    '    Private Const COL_ACode05Name As Integer = 66         ' Maõ phaân tích taøi saûn 5(Teân)
    '    Private Const COL_IsTools As Integer = 67             ' CCDC
    '    Private Const COL_IsCompleted As Integer = 68         ' Đã hình thành
    '    Private Const COL_IsLiquidated As Integer = 69        ' Đã thanh lý
    '    Private Const COL_IsPledgedD23 As Integer = 70        ' Đang thế chấp
    '    Private Const COL_Disabled As Integer = 71            ' KSD
    '    Private Const COL_CreateUserID As Integer = 72        ' CreateUserID
    '    Private Const COL_CreateDate As Integer = 73          ' CreateDate
    '    Private Const COL_LastModifyUserID As Integer = 74    ' LastModifyUserID
    '    Private Const COL_LastModifyDate As Integer = 75      ' LastModifyDate
    '    Private Const COL_UsePeriod As Integer = 76           ' Kỳ sử dụng
    '    Private Const COL_DeptPeriod As Integer = 77          ' Kỳ bắt đầu tính KH
    '    Private Const COL_DepDate As Integer = 78             ' Ngày bắt đầu tính KH
    '    Private Const COL_Percentage As Integer = 79          ' Tỷ lệ KH
    '    Private Const COL_PurchaseDate As Integer = 80        ' Ngày mua
    '    Private Const COL_SupplierID As Integer = 81          ' Nhà cung cấp
    '    Private Const COL_SupplierName As Integer = 82        ' Tên nhà cung cấp
    '    Private Const COL_StrRefNo As Integer = 83            ' Số HĐ
    '    Private Const COL_StrRefDate As Integer = 84          ' Ngày HĐ
    '#End Region


#Region "Const of tdbg - Total of Columns: 87"
    Private Const COL_IsPrinted As Integer = 0            ' Chọn
    Private Const COL_AssetID As Integer = 1              ' Mã tài sản
    Private Const COL_AssetName As Integer = 2            ' Tên tài sản
    Private Const COL_Notes As Integer = 3                ' Ghi chú
    Private Const COL_D54ProjectID As Integer = 4         ' Mã dự án
    Private Const COL_D27PropertyProductID As Integer = 5 ' Mã BĐS
    Private Const COL_FANum01 As Integer = 6              ' Num01
    Private Const COL_FANum02 As Integer = 7              ' Num02
    Private Const COL_FANum03 As Integer = 8              ' Num03
    Private Const COL_FANum04 As Integer = 9              ' Num04
    Private Const COL_FANum05 As Integer = 10             ' Num05
    Private Const COL_FANum06 As Integer = 11             ' Num06
    Private Const COL_FANum07 As Integer = 12             ' Num07
    Private Const COL_FANum08 As Integer = 13             ' Num08
    Private Const COL_FANum09 As Integer = 14             ' Num09
    Private Const COL_FANum10 As Integer = 15             ' Num10
    Private Const COL_FAString01 As Integer = 16          ' Str01
    Private Const COL_FAString02 As Integer = 17          ' Str02
    Private Const COL_FAString03 As Integer = 18          ' Str03
    Private Const COL_FAString04 As Integer = 19          ' Str04
    Private Const COL_FAString05 As Integer = 20          ' Str05
    Private Const COL_FAString06 As Integer = 21          ' Str06
    Private Const COL_FAString07 As Integer = 22          ' Str07
    Private Const COL_FAString08 As Integer = 23          ' Str08
    Private Const COL_FAString09 As Integer = 24          ' Str09
    Private Const COL_FAString10 As Integer = 25          ' Str10
    Private Const COL_FADate01 As Integer = 26            ' Date01
    Private Const COL_FADate02 As Integer = 27            ' Date02
    Private Const COL_FADate03 As Integer = 28            ' Date03
    Private Const COL_FADate04 As Integer = 29            ' Date04
    Private Const COL_FADate05 As Integer = 30            ' Date05
    Private Const COL_FADate06 As Integer = 31            ' Date06
    Private Const COL_FADate07 As Integer = 32            ' Date07
    Private Const COL_FADate08 As Integer = 33            ' Date08
    Private Const COL_FADate09 As Integer = 34            ' Date09
    Private Const COL_FADate10 As Integer = 35            ' Date10
    Private Const COL_ShortName As Integer = 36           ' Tên tắt
    Private Const COL_AssetTag As Integer = 37            ' Thẻ tài sản
    Private Const COL_ObjectID As Integer = 38            ' Mã bộ phận tiếp nhận
    Private Const COL_ObjectName As Integer = 39          ' Tên bộ phận tiếp nhận
    Private Const COL_AssetUserID As Integer = 40         ' Mã người tiếp nhận
    Private Const COL_FullName As Integer = 41            ' Tên người tiếp nhận
    Private Const COL_UsedDate As Integer = 42            ' Ngày sử dụng
    Private Const COL_LocationID As Integer = 43          ' Mã vị trí
    Private Const COL_NewLocationName As Integer = 44     ' Tên vị trí
    Private Const COL_ConvertedAmount As Integer = 45     ' Nguyên giá
    Private Const COL_NotDEPCurrentCost As Integer = 46   ' Giá trị đất
    Private Const COL_DEPCurrentCost As Integer = 47      ' Giá trị xây dựng
    Private Const COL_DepreciatedAmount As Integer = 48   ' Định mức khấu hao
    Private Const COL_AmountDepreciation As Integer = 49  ' Hao mòn lũy kế
    Private Const COL_RemainAmount As Integer = 50        ' Giá trị còn lại
    Private Const COL_AssetAccountID As Integer = 51      ' TK tài sản
    Private Const COL_DepAccountID As Integer = 52        ' TK khấu hao
    Private Const COL_ServiceLife As Integer = 53         ' Số kỳ khấu hao
    Private Const COL_NewServiceLife As Integer = 54      ' Số kỳ khấu hao gốc
    Private Const COL_DepreciatedPeriod As Integer = 55   ' Số kỳ đã khấu hao
    Private Const COL_AssetPeriod As Integer = 56         ' Kỳ nhập tài sản
    Private Const COL_ACode01ID As Integer = 57           ' Maõ phaân tích taøi saûn 1(Maõ)
    Private Const COL_ACode01Name As Integer = 58         ' Maõ phaân tích taøi saûn 1(Teân)
    Private Const COL_ACode02ID As Integer = 59           ' Maõ phaân tích taøi saûn 2(Maõ)
    Private Const COL_ACode02Name As Integer = 60         ' Maõ phaân tích taøi saûn 2(Teân)
    Private Const COL_ACode03ID As Integer = 61           ' Maõ phaân tích taøi saûn 3(Maõ)
    Private Const COL_ACode03Name As Integer = 62         ' Maõ phaân tích taøi saûn 3(Teân)
    Private Const COL_ACode04ID As Integer = 63           ' Maõ phaân tích taøi saûn 4(Maõ)
    Private Const COL_ACode04Name As Integer = 64         ' Maõ phaân tích taøi saûn 4(Teân)
    Private Const COL_ACode05ID As Integer = 65           ' Maõ phaân tích taøi saûn 5(Maõ)
    Private Const COL_ACode05Name As Integer = 66         ' Maõ phaân tích taøi saûn 5(Teân)
    Private Const COL_IsTools As Integer = 67             ' CCDC
    Private Const COL_IsCompleted As Integer = 68         ' Đã hình thành
    Private Const COL_IsLiquidated As Integer = 69        ' Đã thanh lý
    Private Const COL_IsPledgedD23 As Integer = 70        ' Đang thế chấp
    Private Const COL_Disabled As Integer = 71            ' KSD
    Private Const COL_CreateUserID As Integer = 72        ' CreateUserID
    Private Const COL_CreateDate As Integer = 73          ' CreateDate
    Private Const COL_LastModifyUserID As Integer = 74    ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 75      ' LastModifyDate
    Private Const COL_DepreciationStatus As Integer = 76  ' Tình trạng khấu hao
    Private Const COL_UsageStatus As Integer = 77         ' Tình trạng sử dụng
    Private Const COL_UsePeriod As Integer = 78           ' Kỳ sử dụng
    Private Const COL_DeptPeriod As Integer = 79          ' Kỳ bắt đầu tính KH
    Private Const COL_DepDate As Integer = 80             ' Ngày bắt đầu tính KH
    Private Const COL_Percentage As Integer = 81          ' Tỷ lệ KH
    Private Const COL_PurchaseDate As Integer = 82        ' Ngày mua
    Private Const COL_SupplierID As Integer = 83          ' Nhà cung cấp
    Private Const COL_SupplierName As Integer = 84        ' Tên nhà cung cấp
    Private Const COL_StrRefNo As Integer = 85            ' Số HĐ
    Private Const COL_StrRefDate As Integer = 86          ' Ngày HĐ
#End Region


    Private iColumns() As Integer = {COL_ConvertedAmount, COL_DepreciatedAmount, COL_AmountDepreciation, COL_RemainAmount, COL_Percentage, COL_DEPCurrentCost, COL_NotDEPCurrentCost}
    Private dtGrid, dtCaptionCols As DataTable
    Dim arrAcode(4) As Boolean
    Dim bUseACode As Boolean = False
    Dim bFilter As Boolean = False
    Dim bRefreshFilter As Boolean
    Dim sFilter As New System.Text.StringBuilder()
    Dim iperD02F5601 As Integer = -1
    Dim iperD02F5602 As Integer = -1
    Dim iperD02F1030 As Integer = -1

    Private Sub D02F1030_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
            Exit Sub
        End If
        If e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
            Exit Sub
        End If
        If e.KeyCode = Keys.F5 Then
            btnFilter_Click(Nothing, Nothing)
            Exit Sub
        End If
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.NumPad1, Keys.D1
                    If btnManagement.Enabled = True Then
                        btnManagement_Click(Nothing, e)
                    End If
                    Exit Sub
                Case Keys.NumPad2, Keys.D2
                    If btnFinancial.Enabled = True Then
                        btnFinancial_Click(Nothing, e)
                    End If
                    Exit Sub
                Case Keys.NumPad3, Keys.D3
                    If btnAnalysis.Enabled = True Then
                        btnAnalysis_Click(Nothing, e)
                    End If
                    Exit Sub
                Case Keys.NumPad4, Keys.D4
                    If btnAnalysis.Enabled = True Then
                        btnInformation_Click(Nothing, e)
                    End If
                    Exit Sub
            End Select
        End If
    End Sub

    Dim bDelAsset As Boolean = False

    Private Sub D02F1030_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
	LoadInfoGeneral()
        Me.Cursor = Cursors.WaitCursor
        SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1)
        InputDateInTrueDBGrid(tdbg, COL_FADate01, COL_FADate02, COL_FADate03, COL_FADate04, COL_FADate05, COL_FADate06, COL_FADate07, COL_FADate08, COL_FADate09, COL_FADate10, COL_DepDate, COL_PurchaseDate)
        Loadlanguage()
        iperD02F5601 = ReturnPermission("D02F5601")
        iperD02F5602 = ReturnPermission("D02F5602")
        iperD02F1030 = ReturnPermission(_formIDPermission)
        gbEnabledUseFind = False
        ResetColorGrid(tdbg, SPLIT0, SPLIT2)
        ResetSplitDividerSize(tdbg)
        LoadTDBGridInformationCaption(tdbg, COL_FANum01, 1, False, gbUnicode)
        tdbg_NumberFormat()
        SetColumnACode()
        btnManagement_Click(Nothing, Nothing)
        CheckIdTextBox(txtAssetID)
        InputbyUnicode(Me, gbUnicode)
        'Kỳ cuối mới có quyền
        bDelAsset = EnableDelAsset()
        ResetGrid()
        EnableMenu()
        '***************
        'An hien cot CCDC
        tdbg.Splits(2).DisplayColumns(COL_IsTools).Visible = D02Systems.IsAssetIDForD02D43 ' L3Bool(ReturnScalar("SELECT IsAssetIDForD02D43 FROM D02T0000"))
        tdbg.Splits(1).DisplayColumns(COL_D54ProjectID).Visible = D02Systems.CIPforPropertyProduct
        tdbg.Splits(1).DisplayColumns(COL_D27PropertyProductID).Visible = D02Systems.CIPforPropertyProduct


        '8/1/2018, Phạm Thị Thu: id 105528: Bổ sung tính năng "Import danh mục mã CCDC" tại màn hình D02F1030
        tsbImportTools.Visible = D02Systems.IsAssetIDForD02D43
        tsmImportTools.Visible = D02Systems.IsAssetIDForD02D43
        mnsImportTools.Visible = D02Systems.IsAssetIDForD02D43

        SetResolutionForm(Me, ContextMenuStrip1)
        Me.Cursor = Cursors.Default
    End Sub

    'Cap nhat ngay 25/12/2012 theo ID 52909
    Private Sub ReturnPermissionImport()
        mnsAsset.Enabled = iperD02F5601 >= 2
        tsbAsset.Enabled = mnsAsset.Enabled
        tsmAsset.Enabled = mnsAsset.Enabled
        tsbFixedAsset.Enabled = iperD02F5602 >= 2
        tsbFixedAsset.Enabled = tsbFixedAsset.Enabled
        tsmFixedAsset.Enabled = tsbFixedAsset.Enabled
        '8/1/2018, Phạm Thị Thu: id 105528: Bổ sung tính năng "Import danh mục mã CCDC" tại màn hình D02F1030
        tsbImportTools.Enabled = (iperD02F1030 >= 2)
        tsmImportTools.Enabled = (iperD02F1030 >= 2)
        mnsImportTools.Enabled = (iperD02F1030 >= 2)

    End Sub

    'Lấy kỳ đầu tiên dùng để phân quyền : màn hình D02F2000 (Số dư đầu kỳ)
    Private Function EnableDelAsset() As Boolean
        Dim dt As DataTable
        Dim sSQL As String
        sSQL = "Select  TranMonth , TranYear From D02T9999 WITH(NOLOCK)  " & vbCrLf
        sSQL = sSQL & "Where DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL = sSQL & "Order By TranYear , TranMonth "

        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then

            Dim giFirstTranMonth As Integer = CInt(dt.Rows(0)("TranMonth").ToString())
            Dim giFirstTranYear As Integer = CInt(dt.Rows(0)("TranYear").ToString())
            If giFirstTranMonth = giTranMonth And giFirstTranYear = giTranYear Then Return True
        End If
        Return False
    End Function

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_muc_tai_san_co_dinh_-_D02F1030") & UnicodeCaption(gbUnicode) 'Danh móc tªi s¶n cç ¢Ünh - D02F1030
        '================================================================ 
        lblAssetID.Text = rl3("Ma_tai_san_co_chua") 'Mã tài sản có chứa
        lblAssetName.Text = rl3("Ten_tai_san_co_chua") 'Tên tài sản có chứa
        '================================================================
        btnManagement.Text = "1. " & rl3("Thong_tin_quan_ly") '1. Thông tin quản lý
        btnFinancial.Text = "2. " & rl3("Thong_tin_tai_chinh") '2. Thông tin tài chính
        btnAnalysis.Text = "3. " & rl3("Ma_phan_tich") '3. Mã phân tích
        btnInformation.Text = "4. " & rl3("Thong_tin_phu")
        btnFilter.Text = rl3("_Loc") '&Lọc
        tsbAsset.Text = rl3("TSCD")
        tsmAsset.Text = tsbAsset.Text
        mnsAsset.Text = tsbAsset.Text
        tsbFixedAsset.Text = rl3("So_du_TSCD")
        tsmFixedAsset.Text = tsbFixedAsset.Text
        mnsFixedAsset.Text = tsbFixedAsset.Text
        tsmPledgedD23.Text = rL3("Xem_thong_tin_thue_cha_p")
        mnsPledgedD23.Text = tsmPledgedD23.Text
        '================================================================ 
        chkShowAsset.Text = rl3("Hien_nhung_tai_san_da_thanh_ly") 'Hiện những tài sản đã thanh lý
        chkShowDisabled.Text = rL3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng 
        chkIsCompleted.Text = rL3("Hien_thi_tai_san_chua_hinh_thanh") 'Hiển thị tài sản chưa hình thành
        '================================================================ 
        grpFilter.Text = rl3("Tieu_thuc_loc") 'Tiêu thức lọc
        '================================================================ 
        tdbg.Columns(COL_IsPrinted).Caption = rL3("Chon") 'Mã tài sản
        tdbg.Columns(COL_AssetID).Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
        tdbg.Columns(COL_Notes).Caption = rl3("Ghi_chu")
        tdbg.Columns("ShortName").Caption = rl3("Ten_tat") 'Tên tắt
        tdbg.Columns("AssetTag").Caption = rl3("The_tai_san") 'Thẻ tài sản
        tdbg.Columns("ObjectID").Caption = rL3("Ma_bo_phan_tiep_nhan") 'Mã bộ phận tiếp nhận
        tdbg.Columns("ObjectName").Caption = rL3("Ten_bo_phan_tiep_nhan") 'Tên bộ phận tiếp nhận
        tdbg.Columns("AssetUserID").Caption = rl3("Ma_nguoi_tiep_nhan") 'Mã người tiếp nhận
        tdbg.Columns("FullName").Caption = rl3("Ten_nguoi_tiep_nhan") 'Tên người tiếp nhận
        tdbg.Columns("ConvertedAmount").Caption = rL3("Nguyen_gia") 'Nguyên giá
        tdbg.Columns("NotDEPCurrentCost").Caption = rL3("Gia_tri_dat")
        tdbg.Columns("DEPCurrentCost").Caption = rL3("Gia_tri_xay_dung")
        tdbg.Columns("DepreciatedAmount").Caption = rl3("Dinh_muc_khau_hao") 'Định mức khấu hao
        tdbg.Columns("AmountDepreciation").Caption = rl3("Hao_mon_luy_ke") 'Hao mòn lũy kế
        tdbg.Columns("RemainAmount").Caption = rl3("Gia_tri_con_lai") 'Giá trị còn lại
        tdbg.Columns("AssetAccountID").Caption = rl3("TK_tai_san") 'TK tài sản
        tdbg.Columns("DepAccountID").Caption = rl3("TK_khau_hao") 'TK khấu hao
        tdbg.Columns("ServiceLife").Caption = rl3("So_ky_khau_hao") 'Số kỳ khấu hao
        tdbg.Columns("NewServiceLife").Caption = rl3("So_ky_khau_hao_goc") 'Số kỳ khấu hao gốc
        tdbg.Columns("DepreciatedPeriod").Caption = rl3("So_ky_da_khau_hao") 'Số kỳ đã khấu hao
        tdbg.Columns("AssetPeriod").Caption = rl3("Ky_nhap_tai_san") 'Kỳ nhập tài sản
        tdbg.Columns("ACode01ID").Caption = rl3("Ma_phan_tich_tai_san") & " 1" 'Mã phân tích tài sản 1
        tdbg.Columns("ACode01Name").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("ACode02ID").Caption = rl3("Ma_phan_tich_tai_san") & " 2" 'Mã phân tích tài sản 2
        tdbg.Columns("ACode02Name").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("ACode03ID").Caption = rl3("Ma_phan_tich_tai_san") & " 3" 'Mã phân tích tài sản 3
        tdbg.Columns("ACode03Name").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("ACode04ID").Caption = rl3("Ma_phan_tich_tai_san") & " 4" 'Mã phân tích tài sản 4
        tdbg.Columns("ACode04Name").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("ACode05ID").Caption = rl3("Ma_phan_tich_tai_san") & " 5" 'Mã phân tích tài sản 5
        tdbg.Columns("ACode05Name").Caption = rL3("Dien_giai") 'Diễn giải
        tdbg.Columns(COL_IsTools).Caption = rL3("CCDC") 'CCDC
        tdbg.Columns("IsCompleted").Caption = rl3("Da_hinh_thanh") 'Đã hình thành
        tdbg.Columns("IsLiquidated").Caption = rL3("Da_thanh_ly") 'Đã thanh lý
        tdbg.Columns("IsPledgedD23").Caption = rL3("Dang_the_chap")
        tdbg.Columns("Disabled").Caption = rl3("KSD") 'KSD

        tdbg.Columns("UsePeriod").Caption = rl3("Ky_su_dung") 'Kỳ sử dụng
        tdbg.Columns("DeptPeriod").Caption = rl3("Ky_bat_dau_tinh_KH") 'Kỳ bắt đầu tính KH
        tdbg.Columns("DepDate").Caption = rl3("Ngay_bat_dau_tinh_KH") 'Ngày bắt đầu tính KH
        tdbg.Columns("Percentage").Caption = rL3("Ty_le_KH") 'Tỷ lệ KH
        tdbg.Columns(COL_PurchaseDate).Caption = rL3("Ngay_mua") 'Ngày mua
        tdbg.Columns(COL_SupplierID).Caption = rL3("Nha_cung_cap") 'Nhà cung cấp
        tdbg.Columns(COL_SupplierName).Caption = rL3("Ten_nha_cung_cap") 'Tên nhà cung cấp
        tdbg.Columns(COL_StrRefNo).Caption = rL3("So_HD") 'Số HĐ
        tdbg.Columns(COL_StrRefDate).Caption = rL3("Ngay_HD") 'Ngày HĐ
        tdbg.Columns("D54ProjectID").Caption = rL3("Ma_du_an")
        tdbg.Columns("D27PropertyProductID").Caption = rL3("Ma_BDS")
        tdbg.Columns(COL_UsedDate).Caption = rL3("Ngay_su_dung")

        tdbg.Columns(COL_LocationID).Caption = rL3("Ma_vi_tri") 'Mã vị trí
        tdbg.Columns(COL_NewLocationName).Caption = rL3("Ten_vi_tri") 'Tên vị trí
        '================================================================

        tsmInputFixedAsset.Text = rl3("Phieu_nhap_TSCD") 'Phiếu nhập TSCĐ
        tsmListFixedAsset.Text = rL3("Danh_muc_TSCD") 'Danh mục TSCĐ

        tsbListAttachFixedAsset.Text = rL3("Danh_sach_thiet_bi_dinh_kem_TSCD") 'Danh sách thiết bị đính kèm TSCĐ
        tsmListAttachFixedAsset.Text = rL3("Danh_sach_thiet_bi_dinh_kem_TSCD") 'Danh sách thiết bị đính kèm TSCĐ
        mnsListAttachFixedAsset.Text = rL3("Danh_sach_thiet_bi_dinh_kem_TSCD") 'Danh sách thiết bị đính kèm TSCĐ

        tsmDelAsset.Text = rl3("Xoa_TSCD") & " (" & "Lem&onTool" & ")"
        mnsDelAsset.Text = tsmDelAsset.Text

        tsbImportTools.Text = rL3("Cong_cu_dung_cu") 'Công cụ dụng cụ
        tsmImportTools.Text = tsbImportTools.Text 'Công cụ dụng cụ
        mnsImportTools.Text = tsbImportTools.Text 'Công cụ dụng cụ



        '================================================================ 

        tdbg.Columns(COL_DepreciationStatus).Caption = rL3("Tinh_trang_khau_hao") 'Tình trạng khấu hao
        tdbg.Columns(COL_UsageStatus).Caption = rL3("Tinh_trang_su_dung") 'Tình trạng sử dụng

    End Sub



    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        sFind = ""
        LoadTDBGrid(True)
        CallD09U1111_Button(Not bFilter)
        bFilter = True
    End Sub

    Private Sub CallD09U1111_Button(ByVal bLoadFirst As Boolean)
        'CHÚ Ý: Luôn luôn để đúng thứ tự Split và nút nhấn trên lưới
        Dim arrMaster As New ArrayList
        If bLoadFirst = True Then
            'Những cột bắt buộc nhập
            Dim arrColObligatory() As Integer = {COL_AssetID}
            '-----------------------------------
            'Các cột ở SPLIT0
            AddColVisible(tdbg, SPLIT0, arrMaster, arrColObligatory, , , gbUnicode)
            '-----------------------------------
            'Các cột ở SPLIT1
            'Nút 1
            VisibleColumns(1)
            AddColVisible(tdbg, SPLIT1, arrMaster, arrColObligatory, , , gbUnicode)
            'Nút 2
            VisibleColumns(2)
            AddColVisible(tdbg, SPLIT1, arrMaster, arrColObligatory, , , gbUnicode)
            'Nút 3
            VisibleColumns(3)
            AddColVisible(tdbg, SPLIT1, arrMaster, arrColObligatory, , , gbUnicode)
            'Them ngay 17/8/2012 theo incident 50615 cua THIHUAN boi VANVINH
            'Nút 4
            VisibleColumns(4)
            AddColVisible(tdbg, SPLIT1, arrMaster, arrColObligatory, , , gbUnicode)
            '-----------------------------------
            AddColVisible(tdbg, SPLIT2, arrMaster, arrColObligatory, , , gbUnicode)

            'Bật lại Nút 1 để trở về trạng thái ban đầu
            VisibleColumns(1)

            dtCaptionCols = CreateTableForExcelOnly(tdbg, arrMaster)
        End If
    End Sub

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String
        sSQL = SQLStoreD02P0612()
        dtGrid = ReturnDataTable(sSQL)

        gbEnabledUseFind = dtGrid.Rows.Count > 0

        If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFind = ""
        End If

        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("AssetID = " & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0))
        End If

        If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString

        If Not chkShowDisabled.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "Disabled = 0"
        End If

        'If chkShowAsset.Checked Then
        '    dtGrid.DefaultView.RowFilter = strFind
        'Else
        '    If strFind <> "" Then
        '        dtGrid.DefaultView.RowFilter = strFind & " And IsLiquidated = 0"
        '    Else
        '        dtGrid.DefaultView.RowFilter = "IsLiquidated = 0"
        '    End If
        'End If

        'If chkShowAsset.Checked Then
        '    If strFind <> "" Then strFind &= " And "
        '    strFind &= "IsLiquidated = 1"
        'End If

        If Not chkShowAsset.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "IsLiquidated = 0"
        End If

        If Not chkIsCompleted.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "IsCompleted = 1"
        End If
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        '   SetColumnACode()
        CheckMenu(_formIDPermission, TableToolStrip, tdbg.RowCount, gbEnabledUseFind, False, ContextMenuStrip1, , _formIDPermission)
        tsmDelAsset.Enabled = bDelAsset And tsmDelete.Enabled
        mnsDelAsset.Enabled = tsmDelAsset.Enabled

        tsbListAttachFixedAsset.Enabled = (tdbg.RowCount > 0)
        tsmListAttachFixedAsset.Enabled = (tdbg.RowCount > 0)
        mnsListAttachFixedAsset.Enabled = (tdbg.RowCount > 0)


        '*Thêm ngày 19/10/2012 theo incident 50831 của Bảo Trân bởi Văn Vinh
        'tsbImportData.Enabled = tdbg.RowCount > 0 'ReturnPermission("D02F5602") >= 2 And ReturnPermission(Me.Name) > 0
        'tsmImportData.Enabled = tsbImportData.Enabled
        'mnsImportData.Enabled = tsbImportData.Enabled
        'tsbFixedAsset.Enabled = (ReturnPermission("D02F5602") >= 2)
        'tsmFixedAsset.Enabled = tsbFixedAsset.Enabled
        'mnsFixedAsset.Enabled = tsbFixedAsset.Enabled
        ReturnPermissionImport()
        '****************************************
        FooterSum(tdbg, iColumns, , True)
        FooterTotalGrid(tdbg, COL_AssetID)
        EnableMenu()
    End Sub

    Private Sub EnableMenu()
        mnsPledgedD23.Enabled = ReturnPermission("D23F1010") >= EnumPermission.View AndAlso tdbg.RowCount > 0
        tsmPledgedD23.Enabled = mnsPledgedD23.Enabled
        'D99C0008.Msg((ReturnPermission("D23F1010").ToString))
    End Sub


    Private Sub SetColumnACode()
        Dim sSQL As String
        sSQL = "SELECT * FROM D02T0040 WITH(NOLOCK) "
        sSQL &= "WHERE Type = 'A' AND RIGHT(TypeCodeID, 2) IN ('01','02','03','04','05') "
        sSQL &= "ORDER BY TypeCodeID "
        Dim dtACode As DataTable
        dtACode = ReturnDataTable(sSQL)
        If dtACode.Rows.Count > 0 Then
            For i As Integer = 0 To 4
                ''4/8/2021, id 178264-Lỗi font chữ màn hình danh mục tài sản cố định
                tdbg.Splits(1).DisplayColumns(COL_ACode01ID + i * 2).HeadingStyle.Font = FontUnicode(True)
                tdbg.Splits(1).DisplayColumns(COL_ACode01ID + i * 2 + 1).HeadingStyle.Font = FontUnicode(True)

                If gsLanguage = "84" Then
                    tdbg.Columns(COL_ACode01ID + i * 2).Caption = dtACode.Rows(i).Item("VieTypeCodeNameU").ToString & " (" & rL3("Ma") & ")" 'rL3("Ma_VNI")
                    tdbg.Columns(COL_ACode01ID + i * 2 + 1).Caption = dtACode.Rows(i).Item("VieTypeCodeNameU").ToString & " (" & rL3("Ten") & ")" 'rL3("Ten_VNI")
                Else
                    tdbg.Columns(COL_ACode01ID + i * 2).Caption = dtACode.Rows(i).Item("EngTypeCodeNameU").ToString & " (" & rL3("Ma") & ")" 'rL3("Ma_VNI")
                    tdbg.Columns(COL_ACode01ID + i * 2 + 1).Caption = dtACode.Rows(i).Item("EngTypeCodeNameU").ToString & " (" & rL3("Ten") & ")" 'rL3("Ten_VNI")
                End If
                arrAcode(i) = Not CType(dtACode.Rows(i).Item("Disabled").ToString, Boolean)
                If Not bUseACode And Not CType(dtACode.Rows(i).Item("Disabled").ToString, Boolean) Then
                    bUseACode = True
                End If
            Next
        End If
    End Sub

#Region "Menu bar"

   
    Dim bFormD02F0087 As Boolean = False
    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim sAssetID As String = ""
        Dim sMethodID As String = ""
        Dim sSQLD91T1001_SaveLastKey As String = ""

        If D02Systems.AssetAuto = 2 And D02Systems.IsShowFormAutoCreate = True Then
            Dim frm As New D02F0087
            With frm
                .ShowDialog()
                .Dispose()
                If Not .bChoose Then Exit Sub
                sAssetID = .sAssetID
                sMethodID = .sMethodID
                sSQLD91T1001_SaveLastKey = .sSQLD91T1001_SaveLastKey '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
                bFormD02F0087 = True
            End With
        End If
        Dim f As New D02F1031
        With f
            .AssetID = ""
            .bFormD02F0087 = bFormD02F0087
            .sAssetID = sAssetID
            .sMethodID = sMethodID
            .sSQLD91T1001_SaveLastKey = sSQLD91T1001_SaveLastKey '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If .SavedOK Then LoadTDBGrid(True, .AssetID_D02F1031)
            .Dispose()
        End With
        sAssetID = ""
        sMethodID = ""
        bFormD02F0087 = False
    End Sub

    Private Sub tsbInherit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbInherit.Click, tsmInherit.Click, mnsInherit.Click
        Dim sAssetID As String = ""
        Dim sMethodID As String = ""
        Dim sSQLD91T1001_SaveLastKey As String = ""

        If D02Systems.AssetAuto = 2 And D02Systems.IsShowFormAutoCreate = True Then
            Dim frm As New D02F0087
            With frm
                .ShowDialog()
                .Dispose()
                If Not .bChoose Then Exit Sub
                sAssetID = .sAssetID
                sMethodID = .sMethodID
                sSQLD91T1001_SaveLastKey = .sSQLD91T1001_SaveLastKey '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
                bFormD02F0087 = True
            End With
        End If
        Dim f As New D02F1031
        With f
            .AssetID = tdbg.Columns(COL_AssetID).Text
            .bFormD02F0087 = bFormD02F0087 '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
            .sAssetID = sAssetID
            .sMethodID = sMethodID
            .sSQLD91T1001_SaveLastKey = sSQLD91T1001_SaveLastKey '13/6/2019, Nguyễn Thị Tuyết My:id 120539-Lỗi sinh mã tự động khi chưa lưu
            .LocationID = tdbg.Columns(COL_LocationID).Text
            .IsTools = L3Bool(tdbg.Columns(COL_IsTools).Text)
            .FormState = EnumFormState.FormCopy
            .ShowDialog()
            If .SavedOK Then LoadTDBGrid(True, .AssetID_D02F1031)
            .Dispose()
        End With
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim f As New D02F1031()
        With f
            .AssetID = tdbg.Columns(COL_AssetID).Text
            .Completed = CBool(tdbg.Columns(COL_IsCompleted).Text)
            .LocationID = tdbg.Columns(COL_LocationID).Text
            .IsTools = L3Bool(tdbg.Columns(COL_IsTools).Text)
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        Dim f As New D02F1031
        With f
            .AssetID = tdbg.Columns(COL_AssetID).Text
            .Completed = CBool(tdbg.Columns(COL_IsCompleted).Text)
            .LocationID = tdbg.Columns(COL_LocationID).Text
            .IsTools = L3Bool(tdbg.Columns(COL_IsTools).Text)
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
        End With
        If f.SavedOK Then LoadTDBGrid(False, tdbg.Columns(COL_AssetID).Text)
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub

        Dim sAssetID As String = tdbg.Columns(COL_AssetID).Text
        Dim sAssetName As String = tdbg.Columns(COL_AssetName).Text

        If Not CheckStore("Exec D02P0010 " & SQLString(sAssetID)) Then Exit Sub

        Dim sSQL As String = ""
        If L3Bool(tdbg.Columns(COL_IsTools).Text) Then
            sSQL &= "Delete D02T1001 Where AssetID=" & SQLString(sAssetID) & " And DivisionID = " & SQLString(gsDivisionID) & vbCrLf
            sSQL &= "DELETE D07T0002 WHERE InventoryID = " & SQLString(sAssetID) & " AND IsD19 = 1" & vbCrLf
        Else
            sSQL = "Delete D02T0001 Where AssetID=" & SQLString(sAssetID) & vbCrLf
        End If
        sSQL &= "Delete D02T4001 Where DivisionID=" & SQLString(gsDivisionID) & " And AssetID=" & SQLString(sAssetID) & vbCrLf
        sSQL &= "Delete D02T0004 Where AssetID=" & SQLString(sAssetID) & " And DivisionID=" & SQLString(gsDivisionID)
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL)

        If bRunSQL = True Then
            'ExecuteAuditLog("Assets", "03", sAssetID, sAssetName, "", "", "")
            Lemon3.D91.RunAuditLog("02", "Assets", "03", tdbg.Columns(COL_AssetID).Text, tdbg.Columns(COL_AssetName).Text)
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
            ResetGrid()
            DeleteOK()
        Else
            DeleteNotOK()
        End If
    End Sub

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(Me, tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbExportToExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExportToExcel.Click, tsmExportToExcel.Click, mnsExportToExcel.Click
        '       Dim frm As New D99F2222
        'Gọi form Xuất Excel như sau:
        ResetTableForExcel(tdbg, dtCaptionCols)
        CallShowD99F2222(Me, dtCaptionCols, dtGrid, gsGroupColumns)
        '       ResetTableForExcel(tdbg, dtCaptionCols)
        '       With frm
        '           .FormID = Me.Name
        '           .UseUnicode = gbUnicode
        '           .dtLoadGrid = dtCaptionCols
        '           .GroupColumns = gsGroupColumns
        '           .dtExportTable = dtGrid
        '           .ShowDialog()
        '           .Dispose()
        '       End With
    End Sub

    Private Sub tsbInputFixedAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbInputFixedAsset.Click, tsmInputFixedAsset.Click, mnsInputFixedAsset.Click
        'Dim report As New D99C1003
        'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
        If Not AllowNewD99C2003(report, Me) Then Exit Sub
        '************************************()
        Me.Cursor = Cursors.WaitCursor
        Dim sReportName As String = "D02R3011"
        Dim sSubReportName As String = "D02R0000"
        Dim sReportCaption As String = ""
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        sReportCaption = rL3("Phieu_nhap_TSCDF")
        sPathReport = UnicodeGetReportPath(gbUnicode, D02Options.ReportLanguage, "") & sReportName & ".rpt"

        sSQLSub = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
        UnicodeSubReport(sSubReportName, sSQLSub, , gbUnicode)

        sSQL = SQLStoreD02P1033()
        Dim dtPrint As DataTable = ReturnDataTable(sSQL)
        Dim sSQLList As String = ""
        For i As Integer = 0 To tdbg.RowCount - 1
            sSQLList &= SQLString(tdbg(i, COL_AssetID).ToString) & ", "
        Next
        sSQLList = sSQLList.Substring(0, sSQLList.Length - 2)
        dtPrint = ReturnTableFilter(dtPrint, "AssetID IN (" & sSQLList & ")")

        With report
            .OpenConnection(conn)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(dtPrint)
            .PrintReport(sPathReport, sReportCaption & " - " & sReportName)
        End With

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tsbListFixedAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbListFixedAsset.Click, tsmListFixedAsset.Click, mnsListFixedAsset.Click
        '      'Dim report As New D99C1003
        '      'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
        'If Not AllowNewD99C2003(report, Me) Then Exit Sub
        ''************************************()
        '      Me.Cursor = Cursors.WaitCursor
        '      Dim conn As New SqlConnection(gsConnectionString)
        '      Dim sReportName As String = "D02R3020"
        '      Dim sSubReportName As String = "D02R0000"
        '      Dim sReportCaption As String = ""
        '      Dim sPathReport As String = ""
        '      Dim sSQLList As String = ""
        '      Dim sSQL As String = ""
        '      Dim sSQLSub As String = ""

        '      ExecuteSQL(SQLStoreD02P3020())

        '      sReportCaption = rl3("Danh_sach_TSCD")
        '      sPathReport = UnicodeGetReportPath(gbUnicode, D02Options.ReportLanguage, "") & sReportName & ".rpt"

        '      sSQLSub = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
        '      UnicodeSubReport(sSubReportName, sSQLSub, , gbUnicode)
        '      For i As Integer = 0 To tdbg.RowCount - 1
        '          sSQLList &= SQLString(tdbg(i, COL_AssetID).ToString) & ", "
        '      Next
        '      sSQLList = sSQLList.Substring(0, sSQLList.Length - 2)
        '      sSQL = "Select * "
        '      sSQL &= "From   D02V3020 Where UserID = " & SQLString(gsUserID) & vbCrLf
        '      sSQL &= "And AssetID In (" & sSQLList & ") " & vbCrLf
        '      sSQL &= "Order by GroupID, AssetID"
        '      With report
        '          .OpenConnection(conn)
        '          .AddSub(sSQLSub, sSubReportName & ".rpt")
        '          .AddMain(sSQL)
        '          .PrintReport(sPathReport, sReportCaption & " - " & sReportName)
        '      End With
        '      Me.Cursor = Cursors.Default
        If Not AllowPrint() Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        Print(Me, Me.Name)
        Me.Cursor = Cursors.Default
    End Sub

    Private Function AllowPrint() As Boolean
        tdbg.UpdateData()
        Dim sFil As String = "IsPrinted = True"
        If dtGrid.DefaultView.RowFilter <> "" Then
            sFil &= " AND " & dtGrid.DefaultView.RowFilter
        End If
        Dim dr() As DataRow = dtGrid.Select(sFil)
        If dr.Length <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub printReport(ByVal sReportName As String, ByVal sReportPath As String, ByVal sReportCaption As String, ByVal sSQL As String)
        If Not AllowNewD99C2003(report, Me) Then Exit Sub
        Dim conn As New SqlConnection(gsConnectionString)
        With report
            .OpenConnection(conn)
            Dim sSQLSub As String = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
            Dim sSubReport As String = "D02R0000"
            UnicodeSubReport(sSubReport, sSQLSub, gsDivisionID, gbUnicode)
            .AddSub(sSQLSub, sSubReport & ".rpt")
            .AddMain(sSQL)
            .PrintReport(sReportPath, sReportCaption & " - " & sReportName)
        End With
    End Sub

    Private Sub Print(ByVal form As Form, ByVal sReportTypeID As String, Optional ByVal ModuleID As String = "02")
        Dim sReportName As String = "D02R3020"
        Dim sReportPath As String = ""
        Dim sReportTitle As String = rL3("Danh_sach_TSCD")
        Dim sCustomReport As String = ""
        Dim file As String = D99D0541.GetReportPathNew(ModuleID, sReportTypeID, sReportName, sCustomReport, sReportPath, sReportTitle)
        If sReportName = "" Then Exit Sub

        '  ExecuteSQL(SQLStoreD02P3020())
        Dim sSQL As String = ""
        Dim sSQLList As String = ""
        For i As Integer = 0 To tdbg.RowCount - 1
            If L3Bool(tdbg(i, COL_IsPrinted).ToString) Then
                sSQLList &= SQLString(tdbg(i, COL_AssetID).ToString) & ", "
            End If
        Next
        sSQLList = sSQLList.Substring(0, sSQLList.Length - 2)
        sSQL = "Select * "
        sSQL &= "From   D02V3020 Where UserID = " & SQLString(gsUserID) & vbCrLf
        sSQL &= "And AssetID In (" & sSQLList & ") " & vbCrLf
        sSQL &= "Order by GroupID, AssetID"

        tdbg.UpdateData()
        Dim sFil As String = "IsPrinted = True"
        If dtGrid.DefaultView.RowFilter <> "" Then
            sFil &= " AND " & dtGrid.DefaultView.RowFilter
        End If
        Dim dtTable As DataTable = ReturnTableFilter(dtGrid, sFil, True)
        Select Case file.ToLower
            Case "rpt"
                printReport(sReportName, sReportPath, sReportTitle, SQLStoreD02P3020() & vbCrLf & sSQL)
            Case Else
                D99D0541.PrintOfficeType(sReportTypeID, sReportName, sReportPath, file, dtTable)
        End Select
    End Sub

    Private Sub chkShowDisabled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()
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
            ReLoadTDBGrid() 'Làm giống sự kiện Finder_FindClick. Ví dụ đối với form Báo cáo thường gọi btnPrint_Click(Nothing, Nothing): sFind = "
        End Set
    End Property


    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        ShowFindDialogClient(Finder, dtCaptionCols, Me, "0", gbUnicode)
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Grid"

    Private Sub tdbg_NumberFormat()
        tdbg.Columns(COL_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_DepreciatedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_AmountDepreciation).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_RemainAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_DEPCurrentCost).NumberFormat = DxxFormat.D90_ConvertedDecimals
        tdbg.Columns(COL_NotDEPCurrentCost).NumberFormat = DxxFormat.D90_ConvertedDecimals
        'tdbg.Columns(COL_ServiceLife).NumberFormat = D02Format.DefaultNumber0
        'tdbg.Columns(COL_NewServiceLife).NumberFormat = D02Format.DefaultNumber0
        'tdbg.Columns(COL_DepreciatedPeriod).NumberFormat = D02Format.DefaultNumber0
    End Sub

    Private Sub tdbg_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        Select Case e.ColIndex
            Case COL_Percentage
                If Not L3IsNumeric(tdbg.Columns(e.ColIndex).Text, EnumDataType.Number) Then e.Cancel = True
        End Select
    End Sub

    Private Sub tdbg_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.AfterColUpdate
        tdbg.UpdateData()
    End Sub

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If tdbg.FilterActive Then Exit Sub
        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If

    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        '--- Chỉ cho nhập số
        Select Case tdbg.Col
            Case COL_FADate01, COL_FADate02, COL_FADate03, COL_FADate04, COL_FADate05, COL_FADate06, COL_FADate07, COL_FADate08, COL_FADate09, COL_FADate10
                e.Handled = CheckKeyPress(e.KeyChar)
            Case COL_ConvertedAmount, COL_FANum01, COL_FANum02, COL_FANum03, COL_FANum04, COL_FANum05, COL_FANum06, COL_FANum07, COL_FANum08, COL_FANum09, COL_FANum10
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_DepreciatedAmount
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_AmountDepreciation
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_RemainAmount
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_ServiceLife
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_NewServiceLife
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_DepreciatedPeriod
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.Control And e.KeyCode = Keys.S Then HeadClick(tdbg.Col)
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
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

#End Region

#Region "Store"

    'Private Function SQLStringD02R3011() As String
    '    Dim sSQL As String = ""
    '    sSQL = "Select SupplierName" & UnicodeJoin(gbUnicode) & " As SupplierName, "
    '    sSQL &= "OBJECTADDRESS" & UnicodeJoin(gbUnicode) & " As OBJECTADDRESS,TELNO, VOUCHERNO, "
    '    sSQL &= "VOUCHERDATE, SERINO,Description" & UnicodeJoin(gbUnicode) & " As Description, "
    '    sSQL &= "RECIEVIEDDIVISION" & UnicodeJoin(gbUnicode) & " As RECIEVIEDDIVISION, "
    '    sSQL &= "AssetID, AssetName" & UnicodeJoin(gbUnicode) & " As AssetName, "
    '    sSQL &= "CONVERTEDAMOUNT, TAXAMOUNT "
    '    sSQL &= "From D02V1111 "
    '    sSQL &= "Where DivisionID=" & SQLString(gsDivisionID)
    '    sSQL &= " And TranMonth + TranYear * 100 <= " & giTranMonth + giTranYear * 100
    '    sSQL &= " And AssetID=" & SQLString(tdbg.Columns(COL_AssetID).Text)
    '    sSQL &= " Order by SupplierName Asc"
    '    Return sSQL
    'End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1033
    '# Created User: HUỲNH KHANH
    '# Created Date: 25/03/2016 04:20:51
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1033() As String
        Dim sSQL As String = ""
        sSQL &= ("-- In phieu nhap tai san" & vbCrlf)
        sSQL &= "Exec D02P1033 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[20], NOT NULL
        sSQL &= SQLString(Me.Name) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_AssetID).Text) 'AssetID, varchar[20], NOT NULL
        Return sSQL
    End Function



    Private Function SQLStoreD02P3020() As String
        Dim sSQL As String = ""
        ''Huỳnh Edit 24/06/2010
        'Dim dtTmp As DataTable
        'sSQL = "Select Top 1 TranMonth, TranYear From D90T9999 Where DivisionID = " & SQLString(gsDivisionID) & " Order By TranYear * 100 + TranMonth"
        'dtTmp = ReturnDataTable(sSQL)

        sSQL = "Exec D02P3020 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, VarChar[20], NOT NULL
        sSQL &= SQLString("%") & COMMA 'GroupTypeID, VarChar[20], NOT NULL
        sSQL &= SQLString("%") & COMMA 'TypeCodeID, VarChar[20], NOT NULL
        sSQL &= SQLString("%") & COMMA 'FromAssetID, VarChar[20], NOT NULL
        sSQL &= SQLString("%") & COMMA 'ToAssetID, VarChar[20], NOT NULL
        'If dtTmp.Rows.Count > 0 Then
        '    sSQL &= SQLMoney(dtTmp.Rows(0).Item("TranMonth").ToString) & COMMA 'FromMonth, Money, NOT NULL
        '    sSQL &= SQLMoney(dtTmp.Rows(0).Item("TranYear").ToString) & COMMA 'FromYear, Money, NOT NULL
        'Else
        '    sSQL &= SQLMoney(giTranMonth) & COMMA 'FromMonth, Money, NOT NULL
        '    sSQL &= SQLMoney(giTranYear) & COMMA 'FromYear, Money, NOT NULL
        'End If
        sSQL &= SQLMoney(1) & COMMA 'FromMonth, Money, NOT NULL
        sSQL &= SQLMoney(1900) & COMMA 'FromYear, Money, NOT NULL
        sSQL &= SQLMoney(giTranMonth) & COMMA 'ToMonth, Money, NOT NULL
        sSQL &= SQLMoney(giTranYear) & COMMA 'ToYear, Money, NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'ReportTypeID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'ReportID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL

        Return sSQL
    End Function

    Private Function SQLStoreD02P0612() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0612 "
        sSQL &= SQLString(gsDivisionID) & COMMA
        sSQL &= SQLNumber(giTranMonth) & COMMA
        sSQL &= SQLNumber(giTranYear) & COMMA
        sSQL &= SQLString("") & COMMA 'sFind
        sSQL &= SQLNumber(1) & COMMA 'Mode
        sSQL &= SQLString(txtAssetID.Text) & COMMA 'AssetID
        sSQL &= "N" & SQLString(txtAssetName.Text) & COMMA 'AssetName, varchar[100], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function
#End Region

    Private Sub btnManagement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManagement.Click
        tdbg.Focus()
        VisibleColumns(1)
    End Sub

    Private Sub btnFinancial_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinancial.Click
        tdbg.Focus()
        VisibleColumns(2)
    End Sub

    Private Sub btnAnalysis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAnalysis.Click
        tdbg.Focus()
        VisibleColumns(3)
    End Sub

    Private Sub VisibleColumns(ByVal btn As Integer)
        Select Case btn
            Case 1
                btnManagement.Enabled = False : btnFinancial.Enabled = True : btnAnalysis.Enabled = bUseACode : btnInformation.Enabled = True
                With tdbg.Splits(1).DisplayColumns
                    .Item(COL_ShortName).Visible = True : .Item(COL_AssetTag).Visible = True : .Item(COL_ObjectID).Visible = True : .Item(COL_ObjectName).Visible = True : .Item(COL_AssetUserID).Visible = True : .Item(COL_FullName).Visible = True : .Item(COL_LocationID).Visible = True : .Item(COL_NewLocationName).Visible = True : .Item(COL_DepreciationStatus).Visible = True : .Item(COL_UsageStatus).Visible = True
                    .Item(COL_ConvertedAmount).Visible = False : .Item(COL_DepreciatedAmount).Visible = False : .Item(COL_AmountDepreciation).Visible = False : .Item(COL_RemainAmount).Visible = False : .Item(COL_AssetAccountID).Visible = False : .Item(COL_DepAccountID).Visible = False : .Item(COL_ServiceLife).Visible = False : .Item(COL_NewServiceLife).Visible = False : .Item(COL_DepreciatedPeriod).Visible = False : .Item(COL_AssetPeriod).Visible = False
                    .Item(COL_ACode01ID).Visible = False : .Item(COL_ACode02ID).Visible = False : .Item(COL_ACode03ID).Visible = False : .Item(COL_ACode04ID).Visible = False : .Item(COL_ACode05ID).Visible = False
                    .Item(COL_ACode01Name).Visible = False : .Item(COL_ACode02Name).Visible = False : .Item(COL_ACode03Name).Visible = False : .Item(COL_ACode04Name).Visible = False : .Item(COL_ACode05Name).Visible = False : .Item(COL_D54ProjectID).Visible = False : .Item(COL_D27PropertyProductID).Visible = False

                    '1/2/2021, Trần Trí Thông:id 153065-DAPHACO_Lỗi hiện thị sai ngày sử dụng tai màn hinh Danh mục Tài sản cố định
                    tdbg.Splits(1).DisplayColumns(COL_UsedDate).Visible = D02Systems.IsAssetIDForD02D43

                    For i As Integer = COL_FANum01 To COL_FADate10
                        .Item(i).Visible = False
                    Next
                    'Them ngay 20/82012 theo incident 50615 cua THIHUAN boi VANVINh
                    .Item(COL_UsePeriod).Visible = False
                    .Item(COL_DeptPeriod).Visible = False
                    .Item(COL_DepDate).Visible = False
                    .Item(COL_Percentage).Visible = False
                    .Item(COL_PurchaseDate).Visible = False
                    .Item(COL_SupplierID).Visible = False
                    .Item(COL_SupplierName).Visible = False
                    .Item(COL_StrRefNo).Visible = False
                    .Item(COL_StrRefDate).Visible = False
                End With
                With tdbg.Splits(0).DisplayColumns
                    For i As Integer = COL_FANum01 To COL_FADate10
                        .Item(i).Visible = False
                    Next
                End With
                tdbg.Focus()
                tdbg.SplitIndex = SPLIT1
                tdbg.Col = COL_ShortName

                tdbg.Splits(SPLIT1).DisplayColumns(COL_DEPCurrentCost).Visible = False
                tdbg.Splits(SPLIT1).DisplayColumns(COL_NotDEPCurrentCost).Visible = False
            Case 2
                btnManagement.Enabled = True : btnFinancial.Enabled = False : btnAnalysis.Enabled = bUseACode : btnInformation.Enabled = True
                With tdbg.Splits(1).DisplayColumns
                    .Item(COL_ShortName).Visible = False : .Item(COL_AssetTag).Visible = False : .Item(COL_ObjectID).Visible = False : .Item(COL_ObjectName).Visible = False : .Item(COL_AssetUserID).Visible = False : .Item(COL_FullName).Visible = False : .Item(COL_UsedDate).Visible = False : .Item(COL_LocationID).Visible = False : .Item(COL_NewLocationName).Visible = False : .Item(COL_DepreciationStatus).Visible = False : .Item(COL_UsageStatus).Visible = False
                    .Item(COL_ConvertedAmount).Visible = True : .Item(COL_DepreciatedAmount).Visible = True : .Item(COL_AmountDepreciation).Visible = True : .Item(COL_RemainAmount).Visible = True : .Item(COL_AssetAccountID).Visible = True : .Item(COL_DepAccountID).Visible = True : .Item(COL_ServiceLife).Visible = True : .Item(COL_NewServiceLife).Visible = True : .Item(COL_DepreciatedPeriod).Visible = True : .Item(COL_AssetPeriod).Visible = True
                    .Item(COL_ACode01ID).Visible = False : .Item(COL_ACode02ID).Visible = False : .Item(COL_ACode03ID).Visible = False : .Item(COL_ACode04ID).Visible = False : .Item(COL_ACode05ID).Visible = False
                    .Item(COL_ACode01Name).Visible = False : .Item(COL_ACode02Name).Visible = False : .Item(COL_ACode03Name).Visible = False : .Item(COL_ACode04Name).Visible = False : .Item(COL_ACode05Name).Visible = False : .Item(COL_D54ProjectID).Visible = False : .Item(COL_D27PropertyProductID).Visible = False
                    For i As Integer = COL_FANum01 To COL_FADate10
                        .Item(i).Visible = False
                    Next
                    'Them ngay 20/82012 theo incident 50615 cua THIHUAN boi VANVINh
                    .Item(COL_UsePeriod).Visible = True
                    .Item(COL_DeptPeriod).Visible = True
                    .Item(COL_DepDate).Visible = True
                    .Item(COL_Percentage).Visible = True
                    .Item(COL_PurchaseDate).Visible = True
                    .Item(COL_SupplierID).Visible = True
                    .Item(COL_SupplierName).Visible = True
                    .Item(COL_StrRefNo).Visible = True
                    .Item(COL_StrRefDate).Visible = True
                    '
                End With
                With tdbg.Splits(0).DisplayColumns
                    For i As Integer = COL_FANum01 To COL_FADate10
                        .Item(i).Visible = False
                    Next
                End With
                tdbg.Focus()
                tdbg.SplitIndex = SPLIT1
                tdbg.Col = COL_ConvertedAmount
                tdbg.Splits(SPLIT1).DisplayColumns(COL_DEPCurrentCost).Visible = D02Systems.UseProperty
                tdbg.Splits(SPLIT1).DisplayColumns(COL_NotDEPCurrentCost).Visible = D02Systems.UseProperty
            Case 3
                btnManagement.Enabled = True : btnFinancial.Enabled = True : btnAnalysis.Enabled = False : btnInformation.Enabled = True
                With tdbg.Splits(1).DisplayColumns
                    .Item(COL_ShortName).Visible = False : .Item(COL_AssetTag).Visible = False : .Item(COL_ObjectID).Visible = False : .Item(COL_ObjectName).Visible = False : .Item(COL_AssetUserID).Visible = False : .Item(COL_FullName).Visible = False : .Item(COL_UsedDate).Visible = False : .Item(COL_LocationID).Visible = False : .Item(COL_NewLocationName).Visible = False : .Item(COL_DepreciationStatus).Visible = False : .Item(COL_UsageStatus).Visible = False
                    .Item(COL_ConvertedAmount).Visible = False : .Item(COL_DepreciatedAmount).Visible = False : .Item(COL_AmountDepreciation).Visible = False : .Item(COL_RemainAmount).Visible = False : .Item(COL_AssetAccountID).Visible = False : .Item(COL_DepAccountID).Visible = False : .Item(COL_ServiceLife).Visible = False : .Item(COL_NewServiceLife).Visible = False : .Item(COL_DepreciatedPeriod).Visible = False : .Item(COL_AssetPeriod).Visible = False : .Item(COL_D54ProjectID).Visible = False : .Item(COL_D27PropertyProductID).Visible = False
                    For i As Integer = 0 To 4
                        .Item(COL_ACode01ID + i * 2).Visible = arrAcode(i)
                        .Item(COL_ACode01ID + i * 2 + 1).Visible = arrAcode(i)
                    Next
                    For i As Integer = COL_FANum01 To COL_FADate10
                        .Item(i).Visible = False
                    Next
                    'Them ngay 20/82012 theo incident 50615 cua THIHUAN boi VANVINh
                    .Item(COL_UsePeriod).Visible = False
                    .Item(COL_DeptPeriod).Visible = False
                    .Item(COL_DepDate).Visible = False
                    .Item(COL_Percentage).Visible = False
                    .Item(COL_PurchaseDate).Visible = False
                    .Item(COL_SupplierID).Visible = False
                    .Item(COL_SupplierName).Visible = False
                    .Item(COL_StrRefNo).Visible = False
                    .Item(COL_StrRefDate).Visible = False
                    '
                End With
                With tdbg.Splits(0).DisplayColumns
                    For i As Integer = COL_FANum01 To COL_FADate10
                        .Item(i).Visible = False
                    Next
                End With

                tdbg.Focus()
                tdbg.SplitIndex = SPLIT1
                tdbg.Col = COL_ACode01ID
                tdbg.Splits(SPLIT1).DisplayColumns(COL_DEPCurrentCost).Visible = False
                tdbg.Splits(SPLIT1).DisplayColumns(COL_NotDEPCurrentCost).Visible = False
            Case 4
                btnManagement.Enabled = True : btnFinancial.Enabled = True : btnAnalysis.Enabled = bUseACode : btnInformation.Enabled = False
                With tdbg.Splits(1).DisplayColumns
                    .Item(COL_AssetName).Visible = False
                    .Item(COL_AssetID).Visible = False : .Item(COL_D54ProjectID).Visible = True AndAlso D02Systems.CIPforPropertyProduct : .Item(COL_D27PropertyProductID).Visible = True AndAlso D02Systems.CIPforPropertyProduct

                    For i As Integer = COL_FANum01 To COL_FADate10
                        .Item(i).Visible = ArrSpecVisiable(i - COL_FANum01)
                    Next
                    For i As Integer = COL_FADate10 + 1 To tdbg.Columns.Count - 1
                        .Item(i).Visible = False
                    Next
                    'Them ngay 20/82012 theo incident 50615 cua THIHUAN boi VANVINH
                    .Item(COL_UsePeriod).Visible = False
                    .Item(COL_DeptPeriod).Visible = False
                    .Item(COL_DepDate).Visible = False
                    .Item(COL_Percentage).Visible = False
                    .Item(COL_PurchaseDate).Visible = False
                    .Item(COL_SupplierID).Visible = False
                    .Item(COL_SupplierName).Visible = False
                    .Item(COL_StrRefNo).Visible = False
                    .Item(COL_StrRefDate).Visible = False
                End With
                With tdbg.Splits(0).DisplayColumns
                    For i As Integer = COL_FANum01 To COL_FADate10
                        .Item(i).Visible = False
                    Next
                End With
                tdbg.Focus()
                tdbg.SplitIndex = SPLIT1
                tdbg.Col = COL_ACode01ID
                tdbg.Splits(SPLIT1).DisplayColumns(COL_DEPCurrentCost).Visible = False
                tdbg.Splits(SPLIT1).DisplayColumns(COL_NotDEPCurrentCost).Visible = False
        End Select


    End Sub


    Private Sub txtAssetID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAssetID.KeyPress
        If e.KeyChar = "'" Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub txtAssetName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAssetName.KeyPress
        If e.KeyChar = "'" Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub btnInformation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInformation.Click
        tdbg.Focus()
        VisibleColumns(4)
    End Sub

    Private ArrSpecVisiable(30) As Boolean
    Private Function LoadTDBGridInformationCaption(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_Spec01ID As Integer, ByVal Split As Integer, Optional ByVal IsVisibleColumn As Boolean = False, Optional ByVal bUnicode As Boolean = False) As Boolean
        Dim bUseSpec As Boolean = False
        Dim sSQL As String = SQLStoreD02P0015()
        Dim dt As New DataTable
        dt = ReturnDataTable(sSQL)
        Dim iIndex As Integer = COL_Spec01ID
        Dim i As Integer
        If dt.Rows.Count > 0 Then
            For i = 0 To 29
                If (geLanguage = EnumLanguage.Vietnamese) Then
                    tdbg.Columns(iIndex).Caption = dt.Rows(i).Item("Data84").ToString
                Else
                    tdbg.Columns(iIndex).Caption = dt.Rows(i).Item("Data01").ToString
                End If
                tdbg.Columns(iIndex).Tag = (Convert.ToBoolean(dt.Rows(i).Item("Disabled")))
                If (i < 11) Then
                    tdbg.Columns(iIndex).NumberFormat = InsertFormat(L3Int(dt.Rows(i).Item("DecimalNum")))
                End If
                ArrSpecVisiable(iIndex - COL_Spec01ID) = Convert.ToBoolean(tdbg.Columns(iIndex).Tag)
                If Not bUseSpec And Convert.ToBoolean(tdbg.Columns(iIndex).Tag) = True Then
                    bUseSpec = True
                End If
                tdbg.Splits(Split).DisplayColumns(iIndex).HeadingStyle.Font = FontUnicode(bUnicode, tdbg.Splits(Split).DisplayColumns(iIndex).HeadingStyle.Font.Style) 'New System.Drawing.Font("Lemon3", 8.249999!)
                If IsVisibleColumn Then ' Dùng cho lưới có nhiều nút
                    tdbg.Splits(Split).DisplayColumns(iIndex).Visible = Convert.ToBoolean(tdbg.Columns(iIndex).Tag)
                End If
                iIndex += 1
            Next
        End If
        dt = Nothing
        Return bUseSpec
    End Function

    Private Function InsertFormat(ByVal ONumber As Object) As String
        Dim iNumber As Int16 = Convert.ToInt16(ONumber)
        Dim sRet As String = "#,##0"
        If iNumber = 0 Then
        Else
            sRet &= "." & Strings.StrDup(iNumber, "0")
        End If
        Return sRet
    End Function
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

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1030
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 14/06/2012 03:54:09
    '# Modified User: 
    '# Modified Date: 
    '# Description: Xóa TSCĐ
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1030() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1030 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsLanguage) 'Language, varchar[20], NOT NULL
        Return sSQL
    End Function

    Private Sub tsmDelAsset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsmDelAsset.Click, mnsDelAsset.Click
        If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub
        Dim dtTemp As DataTable = ReturnDataTable(SQLStoreD02P1030)
        If dtTemp.Rows.Count > 0 Then
            If L3Int(dtTemp.Rows(0).Item("Status")) = 0 Then
                DeleteOK()
            Else
                '  MessageBox.Show(dtTemp.Rows(0).Item("Message").ToString, rl3("Thong_bao"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                D99C0008.MsgL3(ConvertVietwareFToUnicode(dtTemp.Rows(0).Item("Message").ToString))
            End If
        End If
        dtTemp.Dispose()
        Dim row As Integer = tdbg.Row
        LoadTDBGrid()
        tdbg.Row = row
    End Sub
    'Thêm ngày 19/10/2012 theo incident 50831 của Bảo Trân bởi Văn Vinh
    Private Sub tsmAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmAsset.Click, tsbAsset.Click, mnsAsset.Click
        '       Me.Cursor = Cursors.WaitCursor
        '       .bSaved = False
        '       Dim frm As New D80F2090
        'Gọi form Import Data như sau:
        '       If CallShowDialogD80F2090(D02, "D02F5601", Me.Name) Then
        '           'Load lại dữ liệu
        '       End If
        '       With frm
        '           .FormActive = "D80F2090"
        '           .FormPermission = "D02F5601"
        '           .ModuleID = D02
        '           .TransTypeID = Me.Name  'Theo TL phân tích
        '           .sFont = IIf(gbUnicode, "UNICODE", "VNI").ToString 'VNI-UNICODE: Theo TL phân tích
        '           .ShowDialog()
        '           If .OutPut01 Then .bSaved = .OutPut01
        '           .Dispose()
        '       End With

        '       If .bSaved Then
        '           'Load lại dữ liệu
        '           LoadTDBGrid()
        '       End If
        '       Me.Cursor = Cursors.Default
        Me.Cursor = Cursors.WaitCursor
        If CallShowDialogD80F2090(D02, "D02F5601", Me.Name) Then
            'Load lại dữ liệu
            LoadTDBGrid()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tsmFixedAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmFixedAsset.Click, tsbFixedAsset.Click, mnsFixedAsset.Click
        '.bSaved = False
        'Dim frm As New D80F2090
        'Gọi form Import Data như sau:
        'With frm
        '    .FormActive = "D80F2090"
        '    .FormPermission = "D02F5601"
        '    .ModuleID = D02
        '    .TransTypeID = "D02F1030A"  'Theo TL phân tích
        '    .sFont = IIf(gbUnicode, "UNICODE", "VNI").ToString 'VNI-UNICODE: Theo TL phân tích
        '    .ShowDialog()
        '    If .OutPut01 Then .bSaved = .OutPut01
        '    .Dispose()
        'End With
        'If .bSaved Then
        '    'Load lại dữ liệu
        '    LoadTDBGrid()
        'End If
        Me.Cursor = Cursors.WaitCursor
        If CallShowDialogD80F2090(D02, "D02F5601", "D02F1030A") Then
            'Load lại dữ liệu
            LoadTDBGrid()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    '8/1/2018, Phạm Thị Thu: id 105528: Bổ sung tính năng "Import danh mục mã CCDC" tại màn hình D02F1030
    'Import công cụ dụng cụ
    Private Sub tsbImportTools_Click(sender As Object, e As EventArgs) Handles tsbImportTools.Click, tsmImportTools.Click, mnsImportTools.Click
        Me.Cursor = Cursors.WaitCursor
        If CallShowDialogD80F2090(D02, "D02F1030", "D02F1030B") Then
            'Load lại dữ liệu
            LoadTDBGrid()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub mnsExportToCustomExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim sReportTypeID As String = "D02F1030"
        Dim sReportName As String = "" '"DXXRXXXX"
        Dim sReportPath As String = ""
        Dim sReportTitle As String = "" 'Thêm biến
        Dim sCustomReport As String = "" '= tdbcTranTypeID.Columns("InvoiceForm").Text
        Try
            Dim file As String = GetReportPathNew("02", sReportTypeID, sReportName, sCustomReport, sReportPath, sReportTitle)
            If sReportName = "" Then Exit Sub
            'MessageBox.Show("DLL D99D0541")
            Select Case file.ToLower
                '            Case "rpt"
                '                printReport(sReportName, sReportPath)
                Case "xls", "xlsx"
                    Me.Cursor = Cursors.WaitCursor
                    Dim sPathFile As String = GetObjectFile(sReportTypeID, sReportName, file, sReportPath)
                    If sPathFile = "" Then Exit Select
                    MyExcel(dtGrid, sPathFile, file, False)
            End Select
        Catch ex As Exception

        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub mnsPledgedD23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsPledgedD23.Click
        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "PledgeItemID", tdbg.Columns(COL_AssetID).Text)
        CallFormShowDialog("D23D1240", "D23F1210", arrPro)
    End Sub

    Private Sub chkIsCompleted_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIsCompleted.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()

    End Sub

    Private Sub chkShowAsset_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAsset.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

    Dim bSelect As Boolean = False 'Mặc định Uncheck - tùy thuộc dữ liệu database
    Private Sub HeadClick(ByVal iCol As Integer)
        If tdbg.RowCount <= 0 Then Exit Sub
        Select Case iCol
            Case COL_IsPrinted, COL_IsTools, COL_IsCompleted, COL_IsLiquidated, COL_IsPledgedD23, COL_Disabled
                L3HeadClick(tdbg, iCol, bSelect)
            Case COL_AssetID, COL_AssetName, COL_Notes, COL_FADate01, COL_FADate02, COL_FADate03, COL_FADate04, COL_FADate05, COL_FADate06, COL_FADate07, COL_FADate08, COL_FADate09, COL_FADate10, COL_ShortName, COL_AssetTag, COL_ObjectID, COL_ObjectName, COL_AssetUserID, COL_FullName, COL_UsedDate, COL_LocationID, COL_NewLocationName, COL_UsePeriod, COL_DeptPeriod, COL_DepDate, COL_Percentage, COL_PurchaseDate, COL_SupplierID, COL_SupplierName, COL_StrRefNo, COL_StrRefDate
                tdbg.AllowSort = False
                'Copy 1 cột
                'CopyColumns(tdbg, iCol, tdbg.Columns(iCol).Text, tdbg.Bookmark)
                '****************************************************
                'Copy nhiều cột
                'Dim iColRelative() As Integer = {COL_XXXXX}
                'CopyColumnArr(tdbg, iCol, iColRelative)
                '****************************************************
                'Copy nhiều cột phụ thuộc cột COL_IsUsed
                'Dim iColRelative() As Integer = {COL_XXXXX}
                'CopyColumnsArr(tdbg, iCol, COL_IsUsed, iColRelative)
                '****************************************************
            Case Else
                tdbg.AllowSort = True
        End Select
    End Sub

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.HeadClick
        HeadClick(e.ColIndex)
    End Sub


    '7/4/2017, 	Phạm Thị Thu: id 96093-[CDS] Thẻ TSCĐ - Danh mục TSCĐ theo chủng loại
    'Danh sách thiết bị đính kèm TSCĐ
    Private Sub tsbListAttachFixedAsset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbListAttachFixedAsset.Click, tsmListAttachFixedAsset.Click, mnsListAttachFixedAsset.Click
        Me.Cursor = Cursors.WaitCursor
        PrintListAttachFixedAsset("D02F1030B")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub PrintListAttachFixedAsset(ByVal sReportTypeID As String, Optional ByVal ModuleID As String = "02")
        Dim sReportName As String = "D02R3020B"
        Dim sReportPath As String = ""
        Dim sReportTitle As String = rL3("Danh_sach_thiet_bi_dinh_kem_TSCD_F")
        Dim sCustomReport As String = ""
        Dim file As String = D99D0541.GetReportPathNew(ModuleID, sReportTypeID, sReportName, sCustomReport, sReportPath, sReportTitle)
        If sReportName = "" Then Exit Sub

        Dim sSQL As String = SQLStoreD02P1034()
        
        Dim dtTable As DataTable = ReturnDataTable(sSQL)
        Select Case file.ToLower
            Case "rpt"
                If Not AllowNewD99C2003(report, Me) Then Exit Sub
                Dim conn As New SqlConnection(gsConnectionString)
                With report
                    .OpenConnection(conn)

                    Dim sSQLSub As String = ""
                    sSQLSub = "SELECT CompanyName, CompanyPhone, CompanyFax, AddressLine1, AddressLine2,AddressLine3, AddressLine4, AddressLine5, CompanyAddress" & vbCrLf
                    sSQLSub &= "FROM D91V0016 WHERE DivisionID = " & SQLString(gsDivisionID)
                    Dim sSubReport As String = "D91R0000"
                    UnicodeSubReport(sSubReport, sSQLSub, gsDivisionID, gbUnicode)

                    .AddSub(sSQLSub, sSubReport & ".rpt")

                    .AddMain(sSQL)
                    .PrintReport(sReportPath, sReportTitle & " - " & sReportName)
                End With
            Case Else
                D99D0541.PrintOfficeType(sReportTypeID, sReportName, sReportPath, file, dtTable)
        End Select

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1034
    '# Created User: NGOCTHOAI
    '# Created Date: 07/04/2017 02:07:31
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1034() As String
        Dim sSQL As String = ""
        sSQL &= ("-- do nguon in Danh sach thiet bi dinh kem TSCD " & vbCrlf)
        sSQL &= "Exec D02P1034 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber(0) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLString("") 'AssetID, varchar[20], NOT NULL
        Return sSQL
    End Function


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AnchorForControl(EnumAnchorStyles.TopLeftRight, grpFilter, txtAssetName)
        AnchorForControl(EnumAnchorStyles.TopRight, btnFilter, btnManagement, btnFinancial, btnAnalysis, btnInformation)
        AnchorResizeColumnsGrid(EnumAnchorStyles.TopLeftRightBottom, tdbg)
        AnchorForControl(EnumAnchorStyles.BottomLeft, chkShowDisabled, chkShowAsset)

    End Sub

 
End Class