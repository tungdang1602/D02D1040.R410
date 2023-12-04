<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F0042
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F0042))
        Me.grp1 = New System.Windows.Forms.GroupBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.chkDisabled = New System.Windows.Forms.CheckBox()
        Me.txtACodeID = New System.Windows.Forms.TextBox()
        Me.tdbcTypeCodeID = New C1.Win.C1List.C1Combo()
        Me.lblTypeCodeID = New System.Windows.Forms.Label()
        Me.txtTypeCodeName = New System.Windows.Forms.TextBox()
        Me.lblACodeID = New System.Windows.Forms.Label()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.grp1.SuspendLayout()
        CType(Me.tdbcTypeCodeID, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grp1
        '
        Me.grp1.Controls.Add(Me.txtDescription)
        Me.grp1.Controls.Add(Me.chkDisabled)
        Me.grp1.Controls.Add(Me.txtACodeID)
        Me.grp1.Controls.Add(Me.tdbcTypeCodeID)
        Me.grp1.Controls.Add(Me.lblTypeCodeID)
        Me.grp1.Controls.Add(Me.txtTypeCodeName)
        Me.grp1.Controls.Add(Me.lblACodeID)
        Me.grp1.Controls.Add(Me.lblDescription)
        Me.grp1.Location = New System.Drawing.Point(6, 0)
        Me.grp1.Name = "grp1"
        Me.grp1.Size = New System.Drawing.Size(512, 100)
        Me.grp1.TabIndex = 0
        Me.grp1.TabStop = False
        '
        'txtDescription
        '
        Me.txtDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtDescription.Location = New System.Drawing.Point(108, 69)
        Me.txtDescription.MaxLength = 250
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(398, 20)
        Me.txtDescription.TabIndex = 7
        '
        'chkDisabled
        '
        Me.chkDisabled.AutoSize = True
        Me.chkDisabled.Location = New System.Drawing.Point(278, 43)
        Me.chkDisabled.Name = "chkDisabled"
        Me.chkDisabled.Size = New System.Drawing.Size(109, 19)
        Me.chkDisabled.TabIndex = 5
        Me.chkDisabled.Text = "Không sử dụng"
        Me.chkDisabled.UseVisualStyleBackColor = True
        '
        'txtACodeID
        '
        Me.txtACodeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtACodeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtACodeID.Location = New System.Drawing.Point(108, 40)
        Me.txtACodeID.MaxLength = 20
        Me.txtACodeID.Name = "txtACodeID"
        Me.txtACodeID.Size = New System.Drawing.Size(128, 20)
        Me.txtACodeID.TabIndex = 4
        '
        'tdbcTypeCodeID
        '
        Me.tdbcTypeCodeID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcTypeCodeID.AllowColMove = False
        Me.tdbcTypeCodeID.AllowSort = False
        Me.tdbcTypeCodeID.AlternatingRows = True
        Me.tdbcTypeCodeID.AutoCompletion = True
        Me.tdbcTypeCodeID.AutoDropDown = True
        Me.tdbcTypeCodeID.Caption = ""
        Me.tdbcTypeCodeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcTypeCodeID.ColumnWidth = 100
        Me.tdbcTypeCodeID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcTypeCodeID.DisplayMember = "TypeCodeID"
        Me.tdbcTypeCodeID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcTypeCodeID.DropDownWidth = 500
        Me.tdbcTypeCodeID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcTypeCodeID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcTypeCodeID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcTypeCodeID.EmptyRows = True
        Me.tdbcTypeCodeID.ExtendRightColumn = True
        Me.tdbcTypeCodeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcTypeCodeID.Images.Add(CType(resources.GetObject("tdbcTypeCodeID.Images"), System.Drawing.Image))
        Me.tdbcTypeCodeID.Location = New System.Drawing.Point(108, 10)
        Me.tdbcTypeCodeID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcTypeCodeID.MaxDropDownItems = CType(8, Short)
        Me.tdbcTypeCodeID.MaxLength = 32767
        Me.tdbcTypeCodeID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcTypeCodeID.Name = "tdbcTypeCodeID"
        Me.tdbcTypeCodeID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcTypeCodeID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcTypeCodeID.Size = New System.Drawing.Size(128, 21)
        Me.tdbcTypeCodeID.TabIndex = 1
        Me.tdbcTypeCodeID.ValueMember = "TypeCodeID"
        Me.tdbcTypeCodeID.PropBag = resources.GetString("tdbcTypeCodeID.PropBag")
        '
        'lblTypeCodeID
        '
        Me.lblTypeCodeID.AutoSize = True
        Me.lblTypeCodeID.Location = New System.Drawing.Point(4, 15)
        Me.lblTypeCodeID.Name = "lblTypeCodeID"
        Me.lblTypeCodeID.Size = New System.Drawing.Size(101, 15)
        Me.lblTypeCodeID.TabIndex = 0
        Me.lblTypeCodeID.Text = "Mã loại phân tích"
        Me.lblTypeCodeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTypeCodeName
        '
        Me.txtTypeCodeName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.txtTypeCodeName.Location = New System.Drawing.Point(241, 10)
        Me.txtTypeCodeName.MaxLength = 250
        Me.txtTypeCodeName.Name = "txtTypeCodeName"
        Me.txtTypeCodeName.ReadOnly = True
        Me.txtTypeCodeName.Size = New System.Drawing.Size(265, 20)
        Me.txtTypeCodeName.TabIndex = 2
        Me.txtTypeCodeName.TabStop = False
        '
        'lblACodeID
        '
        Me.lblACodeID.AutoSize = True
        Me.lblACodeID.Location = New System.Drawing.Point(4, 45)
        Me.lblACodeID.Name = "lblACodeID"
        Me.lblACodeID.Size = New System.Drawing.Size(89, 15)
        Me.lblACodeID.TabIndex = 3
        Me.lblACodeID.Text = "Mã khoản mục"
        Me.lblACodeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblDescription
        '
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Location = New System.Drawing.Point(4, 74)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(56, 15)
        Me.lblDescription.TabIndex = 6
        Me.lblDescription.Text = "Diễn giải"
        Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(278, 110)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(76, 22)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "&Lưu"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(360, 110)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(76, 22)
        Me.btnNext.TabIndex = 2
        Me.btnNext.Text = "Nhập &tiếp"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(442, 110)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 22)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'D02F0042
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(524, 142)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.grp1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F0042"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CËp nhËt mº ph¡n tÛch - D02F0042"
        Me.grp1.ResumeLayout(False)
        Me.grp1.PerformLayout()
        CType(Me.tdbcTypeCodeID, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grp1 As System.Windows.Forms.GroupBox
    Private WithEvents tdbcTypeCodeID As C1.Win.C1List.C1Combo
    Private WithEvents txtDescription As System.Windows.Forms.TextBox
    Private WithEvents chkDisabled As System.Windows.Forms.CheckBox
    Private WithEvents txtACodeID As System.Windows.Forms.TextBox
    Private WithEvents lblTypeCodeID As System.Windows.Forms.Label
    Private WithEvents txtTypeCodeName As System.Windows.Forms.TextBox
    Private WithEvents lblACodeID As System.Windows.Forms.Label
    Private WithEvents lblDescription As System.Windows.Forms.Label
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnNext As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
End Class