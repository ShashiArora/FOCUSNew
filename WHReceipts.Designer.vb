<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WHReceipts
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.lblScan = New System.Windows.Forms.Label()
        Me.txtGRNscan = New System.Windows.Forms.TextBox()
        Me.btnImageClear = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtRceiptNo = New System.Windows.Forms.TextBox()
        Me.GroupBoxEdit = New System.Windows.Forms.GroupBox()
        Me.DataGridViewReceipts = New System.Windows.Forms.DataGridView()
        Me.btnRecAccept = New System.Windows.Forms.Button()
        Me.lblstkroom = New System.Windows.Forms.Label()
        Me.CheckBoxupdate = New System.Windows.Forms.CheckBox()
        Me.ComboBoxBINS = New System.Windows.Forms.ComboBox()
        Me.txtStockRoom = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpReceiptdt = New System.Windows.Forms.DateTimePicker()
        Me.txtGRNNO = New System.Windows.Forms.TextBox()
        Me.txtGRNDate = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtPONumber = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtVendorID = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtRemarks = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxSelAll = New System.Windows.Forms.CheckBox()
        Me.ComboBoxMatType = New System.Windows.Forms.ComboBox()
        Me.CheckBoxLoad = New System.Windows.Forms.CheckBox()
        Me.GroupBoxMenu = New Focus.myGroupBox()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.BtnView = New System.Windows.Forms.Button()
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.GroupBoxEdit.SuspendLayout()
        CType(Me.DataGridViewReceipts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBoxMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblScan
        '
        Me.lblScan.AutoSize = True
        Me.lblScan.Location = New System.Drawing.Point(4, 15)
        Me.lblScan.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblScan.Name = "lblScan"
        Me.lblScan.Size = New System.Drawing.Size(77, 13)
        Me.lblScan.TabIndex = 0
        Me.lblScan.Text = "GRN  Barcode"
        '
        'txtGRNscan
        '
        Me.txtGRNscan.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGRNscan.Location = New System.Drawing.Point(88, 11)
        Me.txtGRNscan.Margin = New System.Windows.Forms.Padding(2)
        Me.txtGRNscan.Name = "txtGRNscan"
        Me.txtGRNscan.Size = New System.Drawing.Size(418, 28)
        Me.txtGRNscan.TabIndex = 1
        '
        'btnImageClear
        '
        Me.btnImageClear.Enabled = False
        Me.btnImageClear.Location = New System.Drawing.Point(243, 49)
        Me.btnImageClear.Margin = New System.Windows.Forms.Padding(2)
        Me.btnImageClear.Name = "btnImageClear"
        Me.btnImageClear.Size = New System.Drawing.Size(65, 22)
        Me.btnImageClear.TabIndex = 3
        Me.btnImageClear.Text = "&OK"
        Me.btnImageClear.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 55)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "GRN No."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(29, 33)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(23, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "OR"
        '
        'txtRceiptNo
        '
        Me.txtRceiptNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRceiptNo.Location = New System.Drawing.Point(904, 11)
        Me.txtRceiptNo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRceiptNo.Name = "txtRceiptNo"
        Me.txtRceiptNo.Size = New System.Drawing.Size(143, 24)
        Me.txtRceiptNo.TabIndex = 10
        '
        'GroupBoxEdit
        '
        Me.GroupBoxEdit.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GroupBoxEdit.Controls.Add(Me.DataGridViewReceipts)
        Me.GroupBoxEdit.Controls.Add(Me.btnRecAccept)
        Me.GroupBoxEdit.Location = New System.Drawing.Point(20, 195)
        Me.GroupBoxEdit.Name = "GroupBoxEdit"
        Me.GroupBoxEdit.Size = New System.Drawing.Size(1057, 441)
        Me.GroupBoxEdit.TabIndex = 244
        Me.GroupBoxEdit.TabStop = False
        Me.GroupBoxEdit.Text = "Item Details"
        '
        'DataGridViewReceipts
        '
        Me.DataGridViewReceipts.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.DataGridViewReceipts.AllowUserToAddRows = False
        Me.DataGridViewReceipts.AllowUserToDeleteRows = False
        Me.DataGridViewReceipts.AllowUserToResizeRows = False
        Me.DataGridViewReceipts.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.DataGridViewReceipts.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewReceipts.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewReceipts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewReceipts.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewReceipts.GridColor = System.Drawing.Color.DimGray
        Me.DataGridViewReceipts.Location = New System.Drawing.Point(9, 18)
        Me.DataGridViewReceipts.Name = "DataGridViewReceipts"
        Me.DataGridViewReceipts.RowHeadersWidth = 56
        Me.DataGridViewReceipts.Size = New System.Drawing.Size(1035, 388)
        Me.DataGridViewReceipts.TabIndex = 245
        Me.DataGridViewReceipts.Visible = False
        '
        'btnRecAccept
        '
        Me.btnRecAccept.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRecAccept.Location = New System.Drawing.Point(982, 411)
        Me.btnRecAccept.Margin = New System.Windows.Forms.Padding(2)
        Me.btnRecAccept.Name = "btnRecAccept"
        Me.btnRecAccept.Size = New System.Drawing.Size(66, 21)
        Me.btnRecAccept.TabIndex = 12
        Me.btnRecAccept.Text = "Accept"
        Me.btnRecAccept.UseVisualStyleBackColor = True
        '
        'lblstkroom
        '
        Me.lblstkroom.AutoSize = True
        Me.lblstkroom.Location = New System.Drawing.Point(496, 168)
        Me.lblstkroom.Name = "lblstkroom"
        Me.lblstkroom.Size = New System.Drawing.Size(51, 13)
        Me.lblstkroom.TabIndex = 250
        Me.lblstkroom.Text = "StkRoom"
        '
        'CheckBoxupdate
        '
        Me.CheckBoxupdate.AutoSize = True
        Me.CheckBoxupdate.Location = New System.Drawing.Point(748, 166)
        Me.CheckBoxupdate.Name = "CheckBoxupdate"
        Me.CheckBoxupdate.Size = New System.Drawing.Size(61, 17)
        Me.CheckBoxupdate.TabIndex = 7
        Me.CheckBoxupdate.Text = "Update"
        Me.CheckBoxupdate.UseVisualStyleBackColor = True
        '
        'ComboBoxBINS
        '
        Me.ComboBoxBINS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxBINS.FormattingEnabled = True
        Me.ComboBoxBINS.Location = New System.Drawing.Point(609, 167)
        Me.ComboBoxBINS.MaxLength = 12
        Me.ComboBoxBINS.Name = "ComboBoxBINS"
        Me.ComboBoxBINS.Size = New System.Drawing.Size(133, 21)
        Me.ComboBoxBINS.TabIndex = 6
        '
        'txtStockRoom
        '
        Me.txtStockRoom.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtStockRoom.Location = New System.Drawing.Point(553, 168)
        Me.txtStockRoom.MaxLength = 6
        Me.txtStockRoom.Name = "txtStockRoom"
        Me.txtStockRoom.Size = New System.Drawing.Size(50, 20)
        Me.txtStockRoom.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(836, 15)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 13)
        Me.Label3.TabIndex = 245
        Me.Label3.Text = "Receipt No."
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(846, 54)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 246
        Me.Label4.Text = "Date"
        '
        'dtpReceiptdt
        '
        Me.dtpReceiptdt.Location = New System.Drawing.Point(904, 50)
        Me.dtpReceiptdt.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpReceiptdt.Name = "dtpReceiptdt"
        Me.dtpReceiptdt.Size = New System.Drawing.Size(143, 20)
        Me.dtpReceiptdt.TabIndex = 11
        '
        'txtGRNNO
        '
        Me.txtGRNNO.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtGRNNO.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGRNNO.Location = New System.Drawing.Point(88, 50)
        Me.txtGRNNO.Margin = New System.Windows.Forms.Padding(2)
        Me.txtGRNNO.Name = "txtGRNNO"
        Me.txtGRNNO.Size = New System.Drawing.Size(143, 24)
        Me.txtGRNNO.TabIndex = 2
        '
        'txtGRNDate
        '
        Me.txtGRNDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGRNDate.Location = New System.Drawing.Point(704, 15)
        Me.txtGRNDate.Margin = New System.Windows.Forms.Padding(2)
        Me.txtGRNDate.Name = "txtGRNDate"
        Me.txtGRNDate.Size = New System.Drawing.Size(128, 24)
        Me.txtGRNDate.TabIndex = 251
        Me.txtGRNDate.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(638, 15)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 13)
        Me.Label6.TabIndex = 250
        Me.Label6.Text = "GRN Date"
        Me.Label6.Visible = False
        '
        'txtPONumber
        '
        Me.txtPONumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPONumber.Location = New System.Drawing.Point(704, 32)
        Me.txtPONumber.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPONumber.Name = "txtPONumber"
        Me.txtPONumber.Size = New System.Drawing.Size(128, 24)
        Me.txtPONumber.TabIndex = 253
        Me.txtPONumber.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(638, 39)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(62, 13)
        Me.Label7.TabIndex = 252
        Me.Label7.Text = "PO Number"
        Me.Label7.Visible = False
        '
        'txtVendorID
        '
        Me.txtVendorID.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorID.Location = New System.Drawing.Point(704, 47)
        Me.txtVendorID.Margin = New System.Windows.Forms.Padding(2)
        Me.txtVendorID.Name = "txtVendorID"
        Me.txtVendorID.Size = New System.Drawing.Size(128, 24)
        Me.txtVendorID.TabIndex = 255
        Me.txtVendorID.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(638, 56)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(55, 13)
        Me.Label8.TabIndex = 254
        Me.Label8.Text = "Vendor ID"
        Me.Label8.Visible = False
        '
        'txtRemarks
        '
        Me.txtRemarks.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRemarks.Location = New System.Drawing.Point(88, 82)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRemarks.MaxLength = 150
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(959, 24)
        Me.txtRemarks.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 86)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 256
        Me.Label5.Text = "Remarks"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtVendorID)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtRemarks)
        Me.GroupBox1.Controls.Add(Me.txtPONumber)
        Me.GroupBox1.Controls.Add(Me.lblScan)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtGRNDate)
        Me.GroupBox1.Controls.Add(Me.txtGRNscan)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.dtpReceiptdt)
        Me.GroupBox1.Controls.Add(Me.txtGRNNO)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.btnImageClear)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtRceiptNo)
        Me.GroupBox1.Location = New System.Drawing.Point(25, 54)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(1052, 112)
        Me.GroupBox1.TabIndex = 258
        Me.GroupBox1.TabStop = False
        '
        'CheckBoxSelAll
        '
        Me.CheckBoxSelAll.AutoSize = True
        Me.CheckBoxSelAll.Location = New System.Drawing.Point(97, 173)
        Me.CheckBoxSelAll.Name = "CheckBoxSelAll"
        Me.CheckBoxSelAll.Size = New System.Drawing.Size(55, 17)
        Me.CheckBoxSelAll.TabIndex = 259
        Me.CheckBoxSelAll.Text = "Sel All"
        Me.CheckBoxSelAll.UseVisualStyleBackColor = True
        '
        'ComboBoxMatType
        '
        Me.ComboBoxMatType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxMatType.FormattingEnabled = True
        Me.ComboBoxMatType.Location = New System.Drawing.Point(833, 167)
        Me.ComboBoxMatType.MaxLength = 12
        Me.ComboBoxMatType.Name = "ComboBoxMatType"
        Me.ComboBoxMatType.Size = New System.Drawing.Size(184, 21)
        Me.ComboBoxMatType.TabIndex = 8
        '
        'CheckBoxLoad
        '
        Me.CheckBoxLoad.AutoSize = True
        Me.CheckBoxLoad.Location = New System.Drawing.Point(1023, 166)
        Me.CheckBoxLoad.Name = "CheckBoxLoad"
        Me.CheckBoxLoad.Size = New System.Drawing.Size(50, 17)
        Me.CheckBoxLoad.TabIndex = 9
        Me.CheckBoxLoad.Text = "Load"
        Me.CheckBoxLoad.UseVisualStyleBackColor = True
        '
        'GroupBoxMenu
        '
        Me.GroupBoxMenu.BackColor = System.Drawing.Color.SteelBlue
        Me.GroupBoxMenu.BorderColor = System.Drawing.Color.Black
        Me.GroupBoxMenu.Controls.Add(Me.btnDelete)
        Me.GroupBoxMenu.Controls.Add(Me.BtnView)
        Me.GroupBoxMenu.Controls.Add(Me.BtnAdd)
        Me.GroupBoxMenu.Location = New System.Drawing.Point(25, 10)
        Me.GroupBoxMenu.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxMenu.Name = "GroupBoxMenu"
        Me.GroupBoxMenu.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxMenu.Size = New System.Drawing.Size(1052, 39)
        Me.GroupBoxMenu.TabIndex = 249
        Me.GroupBoxMenu.TabStop = False
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(132, 8)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(2)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(56, 27)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'BtnView
        '
        Me.BtnView.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnView.Location = New System.Drawing.Point(71, 8)
        Me.BtnView.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnView.Name = "BtnView"
        Me.BtnView.Size = New System.Drawing.Size(56, 27)
        Me.BtnView.TabIndex = 1
        Me.BtnView.Text = "View"
        Me.BtnView.UseVisualStyleBackColor = True
        '
        'BtnAdd
        '
        Me.BtnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAdd.Location = New System.Drawing.Point(4, 7)
        Me.BtnAdd.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(62, 28)
        Me.BtnAdd.TabIndex = 0
        Me.BtnAdd.Text = "New"
        Me.BtnAdd.UseVisualStyleBackColor = True
        '
        'WHReceipts
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1092, 639)
        Me.Controls.Add(Me.CheckBoxLoad)
        Me.Controls.Add(Me.ComboBoxMatType)
        Me.Controls.Add(Me.CheckBoxSelAll)
        Me.Controls.Add(Me.CheckBoxupdate)
        Me.Controls.Add(Me.lblstkroom)
        Me.Controls.Add(Me.ComboBoxBINS)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtStockRoom)
        Me.Controls.Add(Me.GroupBoxMenu)
        Me.Controls.Add(Me.GroupBoxEdit)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "WHReceipts"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Receipts"
        Me.GroupBoxEdit.ResumeLayout(False)
        CType(Me.DataGridViewReceipts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBoxMenu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblScan As System.Windows.Forms.Label
    Friend WithEvents txtGRNscan As System.Windows.Forms.TextBox
    Friend WithEvents btnImageClear As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtRceiptNo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBoxEdit As System.Windows.Forms.GroupBox
    Friend WithEvents btnRecAccept As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpReceiptdt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtGRNNO As System.Windows.Forms.TextBox
    Friend WithEvents BtnAdd As System.Windows.Forms.Button
    Friend WithEvents GroupBoxMenu As Focus.myGroupBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnView As System.Windows.Forms.Button
    Friend WithEvents DataGridViewReceipts As System.Windows.Forms.DataGridView
    Friend WithEvents txtGRNDate As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPONumber As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtVendorID As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblstkroom As System.Windows.Forms.Label
    Friend WithEvents CheckBoxupdate As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxBINS As System.Windows.Forms.ComboBox
    Friend WithEvents txtStockRoom As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxSelAll As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxMatType As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxLoad As System.Windows.Forms.CheckBox
End Class
