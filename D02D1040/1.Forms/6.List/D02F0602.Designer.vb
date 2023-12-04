<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F0602
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F0602))
        Me.txtDivisionID = New System.Windows.Forms.TextBox()
        Me.lblDivisionID = New System.Windows.Forms.Label()
        Me.txtAssetID = New System.Windows.Forms.TextBox()
        Me.lblAssetID = New System.Windows.Forms.Label()
        Me.txtAssetName = New System.Windows.Forms.TextBox()
        Me.btnAssetInfo = New System.Windows.Forms.Button()
        Me.btnConvertedAmount = New System.Windows.Forms.Button()
        Me.btnDepreciation = New System.Windows.Forms.Button()
        Me.btnEffectHistory = New System.Windows.Forms.Button()
        Me.tdbgConvertedAmount = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.txtCost = New System.Windows.Forms.TextBox()
        Me.lblCost = New System.Windows.Forms.Label()
        Me.tdbgDepreciation = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tab = New System.Windows.Forms.TabControl()
        Me.tab01 = New System.Windows.Forms.TabPage()
        Me.tdbg01 = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tab02 = New System.Windows.Forms.TabPage()
        Me.tdbg02 = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tab03 = New System.Windows.Forms.TabPage()
        Me.tdbg03 = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tab04 = New System.Windows.Forms.TabPage()
        Me.tdbg04 = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.txtAmountDepreciation = New System.Windows.Forms.TextBox()
        Me.lblAmountDepreciation = New System.Windows.Forms.Label()
        CType(Me.tdbgConvertedAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbgDepreciation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tab.SuspendLayout()
        Me.tab01.SuspendLayout()
        CType(Me.tdbg01, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tab02.SuspendLayout()
        CType(Me.tdbg02, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tab03.SuspendLayout()
        CType(Me.tdbg03, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tab04.SuspendLayout()
        CType(Me.tdbg04, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDivisionID
        '
        Me.txtDivisionID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtDivisionID.Location = New System.Drawing.Point(78, 6)
        Me.txtDivisionID.Name = "txtDivisionID"
        Me.txtDivisionID.Size = New System.Drawing.Size(128, 20)
        Me.txtDivisionID.TabIndex = 0
        '
        'lblDivisionID
        '
        Me.lblDivisionID.AutoSize = True
        Me.lblDivisionID.Location = New System.Drawing.Point(3, 10)
        Me.lblDivisionID.Name = "lblDivisionID"
        Me.lblDivisionID.Size = New System.Drawing.Size(55, 13)
        Me.lblDivisionID.TabIndex = 1
        Me.lblDivisionID.Text = "Mã đơn vị"
        Me.lblDivisionID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtAssetID
        '
        Me.txtAssetID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtAssetID.Location = New System.Drawing.Point(284, 6)
        Me.txtAssetID.Name = "txtAssetID"
        Me.txtAssetID.Size = New System.Drawing.Size(128, 20)
        Me.txtAssetID.TabIndex = 2
        '
        'lblAssetID
        '
        Me.lblAssetID.AutoSize = True
        Me.lblAssetID.Location = New System.Drawing.Point(216, 10)
        Me.lblAssetID.Name = "lblAssetID"
        Me.lblAssetID.Size = New System.Drawing.Size(54, 13)
        Me.lblAssetID.TabIndex = 3
        Me.lblAssetID.Text = "Mã TSCĐ"
        Me.lblAssetID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtAssetName
        '
        Me.txtAssetName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtAssetName.Location = New System.Drawing.Point(415, 6)
        Me.txtAssetName.Name = "txtAssetName"
        Me.txtAssetName.Size = New System.Drawing.Size(354, 20)
        Me.txtAssetName.TabIndex = 4
        '
        'btnAssetInfo
        '
        Me.btnAssetInfo.Location = New System.Drawing.Point(337, 34)
        Me.btnAssetInfo.Name = "btnAssetInfo"
        Me.btnAssetInfo.Size = New System.Drawing.Size(100, 22)
        Me.btnAssetInfo.TabIndex = 5
        Me.btnAssetInfo.Text = "Thông tin tài sản"
        Me.btnAssetInfo.UseVisualStyleBackColor = True
        '
        'btnConvertedAmount
        '
        Me.btnConvertedAmount.Location = New System.Drawing.Point(442, 34)
        Me.btnConvertedAmount.Name = "btnConvertedAmount"
        Me.btnConvertedAmount.Size = New System.Drawing.Size(100, 22)
        Me.btnConvertedAmount.TabIndex = 6
        Me.btnConvertedAmount.Text = "Nguyên giá"
        Me.btnConvertedAmount.UseVisualStyleBackColor = True
        '
        'btnDepreciation
        '
        Me.btnDepreciation.Location = New System.Drawing.Point(547, 34)
        Me.btnDepreciation.Name = "btnDepreciation"
        Me.btnDepreciation.Size = New System.Drawing.Size(100, 22)
        Me.btnDepreciation.TabIndex = 7
        Me.btnDepreciation.Text = "Khấu hao"
        Me.btnDepreciation.UseVisualStyleBackColor = True
        '
        'btnEffectHistory
        '
        Me.btnEffectHistory.Location = New System.Drawing.Point(653, 34)
        Me.btnEffectHistory.Name = "btnEffectHistory"
        Me.btnEffectHistory.Size = New System.Drawing.Size(116, 22)
        Me.btnEffectHistory.TabIndex = 8
        Me.btnEffectHistory.Text = "Lịch sử tác động"
        Me.btnEffectHistory.UseVisualStyleBackColor = True
        '
        'tdbgConvertedAmount
        '
        Me.tdbgConvertedAmount.AllowColMove = False
        Me.tdbgConvertedAmount.AllowColSelect = False
        Me.tdbgConvertedAmount.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbgConvertedAmount.AllowSort = False
        Me.tdbgConvertedAmount.AllowUpdate = False
        Me.tdbgConvertedAmount.AlternatingRows = True
        Me.tdbgConvertedAmount.ColumnFooters = True
        Me.tdbgConvertedAmount.EmptyRows = True
        Me.tdbgConvertedAmount.ExtendRightColumn = True
        Me.tdbgConvertedAmount.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbgConvertedAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbgConvertedAmount.Images.Add(CType(resources.GetObject("tdbgConvertedAmount.Images"), System.Drawing.Image))
        Me.tdbgConvertedAmount.Location = New System.Drawing.Point(6, 63)
        Me.tdbgConvertedAmount.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.HighlightRowRaiseCell
        Me.tdbgConvertedAmount.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbgConvertedAmount.Name = "tdbgConvertedAmount"
        Me.tdbgConvertedAmount.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbgConvertedAmount.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbgConvertedAmount.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbgConvertedAmount.PrintInfo.PageSettings = CType(resources.GetObject("tdbgConvertedAmount.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbgConvertedAmount.Size = New System.Drawing.Size(763, 383)
        Me.tdbgConvertedAmount.TabAcrossSplits = True
        Me.tdbgConvertedAmount.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbgConvertedAmount.TabIndex = 9
        Me.tdbgConvertedAmount.Tag = "COLMODE1"
        Me.tdbgConvertedAmount.PropBag = resources.GetString("tdbgConvertedAmount.PropBag")
        '
        'txtCost
        '
        Me.txtCost.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtCost.Location = New System.Drawing.Point(641, 462)
        Me.txtCost.Name = "txtCost"
        Me.txtCost.Size = New System.Drawing.Size(128, 20)
        Me.txtCost.TabIndex = 10
        '
        'lblCost
        '
        Me.lblCost.AutoSize = True
        Me.lblCost.Location = New System.Drawing.Point(533, 466)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.Size = New System.Drawing.Size(93, 13)
        Me.lblCost.TabIndex = 11
        Me.lblCost.Text = "Nguyên giá TSCĐ"
        Me.lblCost.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tdbgDepreciation
        '
        Me.tdbgDepreciation.AllowColMove = False
        Me.tdbgDepreciation.AllowColSelect = False
        Me.tdbgDepreciation.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbgDepreciation.AllowSort = False
        Me.tdbgDepreciation.AllowUpdate = False
        Me.tdbgDepreciation.AlternatingRows = True
        Me.tdbgDepreciation.ColumnFooters = True
        Me.tdbgDepreciation.EmptyRows = True
        Me.tdbgDepreciation.ExtendRightColumn = True
        Me.tdbgDepreciation.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbgDepreciation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbgDepreciation.Images.Add(CType(resources.GetObject("tdbgDepreciation.Images"), System.Drawing.Image))
        Me.tdbgDepreciation.Location = New System.Drawing.Point(127, 134)
        Me.tdbgDepreciation.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.HighlightRowRaiseCell
        Me.tdbgDepreciation.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbgDepreciation.Name = "tdbgDepreciation"
        Me.tdbgDepreciation.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbgDepreciation.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbgDepreciation.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbgDepreciation.PrintInfo.PageSettings = CType(resources.GetObject("tdbgDepreciation.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbgDepreciation.Size = New System.Drawing.Size(75, 23)
        Me.tdbgDepreciation.TabAcrossSplits = True
        Me.tdbgDepreciation.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbgDepreciation.TabIndex = 12
        Me.tdbgDepreciation.Tag = "COLMODE2"
        Me.tdbgDepreciation.Visible = False
        Me.tdbgDepreciation.PropBag = resources.GetString("tdbgDepreciation.PropBag")
        '
        'tab
        '
        Me.tab.Controls.Add(Me.tab01)
        Me.tab.Controls.Add(Me.tab02)
        Me.tab.Controls.Add(Me.tab03)
        Me.tab.Controls.Add(Me.tab04)
        Me.tab.Location = New System.Drawing.Point(45, 174)
        Me.tab.Name = "tab"
        Me.tab.SelectedIndex = 0
        Me.tab.Size = New System.Drawing.Size(343, 100)
        Me.tab.TabIndex = 13
        Me.tab.Visible = False
        '
        'tab01
        '
        Me.tab01.Controls.Add(Me.tdbg01)
        Me.tab01.Location = New System.Drawing.Point(4, 22)
        Me.tab01.Name = "tab01"
        Me.tab01.Padding = New System.Windows.Forms.Padding(3)
        Me.tab01.Size = New System.Drawing.Size(335, 74)
        Me.tab01.TabIndex = 0
        Me.tab01.Text = "Tiếu thức phân bổ"
        Me.tab01.UseVisualStyleBackColor = True
        '
        'tdbg01
        '
        Me.tdbg01.AllowColMove = False
        Me.tdbg01.AllowColSelect = False
        Me.tdbg01.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg01.AllowSort = False
        Me.tdbg01.AllowUpdate = False
        Me.tdbg01.AlternatingRows = True
        Me.tdbg01.ColumnFooters = True
        Me.tdbg01.EmptyRows = True
        Me.tdbg01.ExtendRightColumn = True
        Me.tdbg01.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg01.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg01.Images.Add(CType(resources.GetObject("tdbg01.Images"), System.Drawing.Image))
        Me.tdbg01.Location = New System.Drawing.Point(0, 3)
        Me.tdbg01.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.HighlightRowRaiseCell
        Me.tdbg01.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg01.Name = "tdbg01"
        Me.tdbg01.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg01.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg01.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg01.PrintInfo.PageSettings = CType(resources.GetObject("tdbg01.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg01.Size = New System.Drawing.Size(75, 23)
        Me.tdbg01.TabAcrossSplits = True
        Me.tdbg01.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg01.TabIndex = 0
        Me.tdbg01.Tag = "COLMODE3"
        Me.tdbg01.PropBag = resources.GetString("tdbg01.PropBag")
        '
        'tab02
        '
        Me.tab02.Controls.Add(Me.tdbg02)
        Me.tab02.Location = New System.Drawing.Point(4, 22)
        Me.tab02.Name = "tab02"
        Me.tab02.Padding = New System.Windows.Forms.Padding(3)
        Me.tab02.Size = New System.Drawing.Size(335, 74)
        Me.tab02.TabIndex = 1
        Me.tab02.Text = "Bộ phận tiếp nhận/ quản lý"
        Me.tab02.UseVisualStyleBackColor = True
        '
        'tdbg02
        '
        Me.tdbg02.AllowColMove = False
        Me.tdbg02.AllowColSelect = False
        Me.tdbg02.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg02.AllowSort = False
        Me.tdbg02.AllowUpdate = False
        Me.tdbg02.AlternatingRows = True
        Me.tdbg02.ColumnFooters = True
        Me.tdbg02.EmptyRows = True
        Me.tdbg02.ExtendRightColumn = True
        Me.tdbg02.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg02.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg02.Images.Add(CType(resources.GetObject("tdbg02.Images"), System.Drawing.Image))
        Me.tdbg02.Location = New System.Drawing.Point(0, 3)
        Me.tdbg02.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.HighlightRowRaiseCell
        Me.tdbg02.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg02.Name = "tdbg02"
        Me.tdbg02.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg02.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg02.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg02.PrintInfo.PageSettings = CType(resources.GetObject("tdbg02.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg02.Size = New System.Drawing.Size(75, 23)
        Me.tdbg02.TabAcrossSplits = True
        Me.tdbg02.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg02.TabIndex = 0
        Me.tdbg02.Tag = "COL"
        Me.tdbg02.PropBag = resources.GetString("tdbg02.PropBag")
        '
        'tab03
        '
        Me.tab03.Controls.Add(Me.tdbg03)
        Me.tab03.Location = New System.Drawing.Point(4, 22)
        Me.tab03.Name = "tab03"
        Me.tab03.Padding = New System.Windows.Forms.Padding(3)
        Me.tab03.Size = New System.Drawing.Size(335, 74)
        Me.tab03.TabIndex = 2
        Me.tab03.Text = "Thời gian khấu hao"
        Me.tab03.UseVisualStyleBackColor = True
        '
        'tdbg03
        '
        Me.tdbg03.AllowColMove = False
        Me.tdbg03.AllowColSelect = False
        Me.tdbg03.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg03.AllowSort = False
        Me.tdbg03.AllowUpdate = False
        Me.tdbg03.AlternatingRows = True
        Me.tdbg03.ColumnFooters = True
        Me.tdbg03.EmptyRows = True
        Me.tdbg03.ExtendRightColumn = True
        Me.tdbg03.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg03.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg03.Images.Add(CType(resources.GetObject("tdbg03.Images"), System.Drawing.Image))
        Me.tdbg03.Location = New System.Drawing.Point(0, 3)
        Me.tdbg03.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.HighlightRowRaiseCell
        Me.tdbg03.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg03.Name = "tdbg03"
        Me.tdbg03.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg03.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg03.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg03.PrintInfo.PageSettings = CType(resources.GetObject("tdbg03.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg03.Size = New System.Drawing.Size(75, 23)
        Me.tdbg03.TabAcrossSplits = True
        Me.tdbg03.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg03.TabIndex = 0
        Me.tdbg03.PropBag = resources.GetString("tdbg03.PropBag")
        '
        'tab04
        '
        Me.tab04.Controls.Add(Me.tdbg04)
        Me.tab04.Location = New System.Drawing.Point(4, 22)
        Me.tab04.Name = "tab04"
        Me.tab04.Padding = New System.Windows.Forms.Padding(3)
        Me.tab04.Size = New System.Drawing.Size(335, 74)
        Me.tab04.TabIndex = 3
        Me.tab04.Text = "Tài khoản tài sản/khấu hao"
        Me.tab04.UseVisualStyleBackColor = True
        '
        'tdbg04
        '
        Me.tdbg04.AllowColMove = False
        Me.tdbg04.AllowColSelect = False
        Me.tdbg04.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg04.AllowSort = False
        Me.tdbg04.AllowUpdate = False
        Me.tdbg04.AlternatingRows = True
        Me.tdbg04.ColumnFooters = True
        Me.tdbg04.EmptyRows = True
        Me.tdbg04.ExtendRightColumn = True
        Me.tdbg04.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg04.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg04.Images.Add(CType(resources.GetObject("tdbg04.Images"), System.Drawing.Image))
        Me.tdbg04.Location = New System.Drawing.Point(0, 3)
        Me.tdbg04.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.HighlightRowRaiseCell
        Me.tdbg04.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg04.Name = "tdbg04"
        Me.tdbg04.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg04.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg04.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg04.PrintInfo.PageSettings = CType(resources.GetObject("tdbg04.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg04.Size = New System.Drawing.Size(75, 23)
        Me.tdbg04.TabAcrossSplits = True
        Me.tdbg04.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg04.TabIndex = 1
        Me.tdbg04.Tag = "COL04"
        Me.tdbg04.PropBag = resources.GetString("tdbg04.PropBag")
        '
        'txtAmountDepreciation
        '
        Me.txtAmountDepreciation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtAmountDepreciation.Location = New System.Drawing.Point(284, 463)
        Me.txtAmountDepreciation.Name = "txtAmountDepreciation"
        Me.txtAmountDepreciation.Size = New System.Drawing.Size(128, 20)
        Me.txtAmountDepreciation.TabIndex = 14
        '
        'lblAmountDepreciation
        '
        Me.lblAmountDepreciation.AutoSize = True
        Me.lblAmountDepreciation.Location = New System.Drawing.Point(190, 468)
        Me.lblAmountDepreciation.Name = "lblAmountDepreciation"
        Me.lblAmountDepreciation.Size = New System.Drawing.Size(84, 13)
        Me.lblAmountDepreciation.TabIndex = 15
        Me.lblAmountDepreciation.Text = "Khấu hao lũy kế"
        Me.lblAmountDepreciation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'D02F0602
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(773, 498)
        Me.Controls.Add(Me.txtAmountDepreciation)
        Me.Controls.Add(Me.tab)
        Me.Controls.Add(Me.tdbgDepreciation)
        Me.Controls.Add(Me.txtCost)
        Me.Controls.Add(Me.tdbgConvertedAmount)
        Me.Controls.Add(Me.btnEffectHistory)
        Me.Controls.Add(Me.btnDepreciation)
        Me.Controls.Add(Me.btnConvertedAmount)
        Me.Controls.Add(Me.btnAssetInfo)
        Me.Controls.Add(Me.txtAssetName)
        Me.Controls.Add(Me.txtAssetID)
        Me.Controls.Add(Me.txtDivisionID)
        Me.Controls.Add(Me.lblDivisionID)
        Me.Controls.Add(Me.lblAssetID)
        Me.Controls.Add(Me.lblCost)
        Me.Controls.Add(Me.lblAmountDepreciation)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F0602"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Truy vÊn th¤ng tin tªi s¶n - D02F0602"
        CType(Me.tdbgConvertedAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbgDepreciation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tab.ResumeLayout(False)
        Me.tab01.ResumeLayout(False)
        CType(Me.tdbg01, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tab02.ResumeLayout(False)
        CType(Me.tdbg02, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tab03.ResumeLayout(False)
        CType(Me.tdbg03, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tab04.ResumeLayout(False)
        CType(Me.tdbg04, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents txtDivisionID As System.Windows.Forms.TextBox
    Private WithEvents lblDivisionID As System.Windows.Forms.Label
    Private WithEvents txtAssetID As System.Windows.Forms.TextBox
    Private WithEvents lblAssetID As System.Windows.Forms.Label
    Private WithEvents txtAssetName As System.Windows.Forms.TextBox
    Private WithEvents btnAssetInfo As System.Windows.Forms.Button
    Private WithEvents btnConvertedAmount As System.Windows.Forms.Button
    Private WithEvents btnDepreciation As System.Windows.Forms.Button
    Private WithEvents btnEffectHistory As System.Windows.Forms.Button
    Private WithEvents tdbgConvertedAmount As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents txtCost As System.Windows.Forms.TextBox
    Private WithEvents lblCost As System.Windows.Forms.Label
    Private WithEvents tdbgDepreciation As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents tab As System.Windows.Forms.TabControl
    Friend WithEvents tab01 As System.Windows.Forms.TabPage
    Friend WithEvents tab02 As System.Windows.Forms.TabPage
    Private WithEvents txtAmountDepreciation As System.Windows.Forms.TextBox
    Private WithEvents lblAmountDepreciation As System.Windows.Forms.Label
    Private WithEvents tdbg01 As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents tdbg02 As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents tab03 As System.Windows.Forms.TabPage
    Private WithEvents tdbg03 As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents tab04 As System.Windows.Forms.TabPage
    Private WithEvents tdbg04 As C1.Win.C1TrueDBGrid.C1TrueDBGrid
End Class