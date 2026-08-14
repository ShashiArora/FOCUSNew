<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WHMaterialReq
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.txtMO = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.DTPReqDt = New System.Windows.Forms.DateTimePicker()
        Me.btnReqSave = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.ComboBoxPur = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ComboBoxUYN = New System.Windows.Forms.ComboBox()
        Me.LabelDetQty = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtSlNo = New System.Windows.Forms.TextBox()
        Me.DataGridViewMaterialReq = New System.Windows.Forms.DataGridView()
        Me.ComboBoxSD = New System.Windows.Forms.ComboBox()
        Me.ComboBoxCell = New System.Windows.Forms.ComboBox()
        Me.ComboBoxdept = New System.Windows.Forms.ComboBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtPartNumber = New System.Windows.Forms.TextBox()
        Me.txtQty = New System.Windows.Forms.TextBox()
        Me.txtDetailRemark = New System.Windows.Forms.TextBox()
        Me.txtReqNo = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnSaveChanges = New System.Windows.Forms.Button()
        Me.lblStk = New System.Windows.Forms.Label()
        Me.LblStockavble = New System.Windows.Forms.Label()
        Me.DataGridViewWHPartNumbers = New System.Windows.Forms.DataGridView()
        Me.ComboBoxuom = New System.Windows.Forms.ComboBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.ComboBoxIssueType = New System.Windows.Forms.ComboBox()
        Me.RBFSItem = New System.Windows.Forms.RadioButton()
        Me.RBNonFSItem = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.DataGridViewMatReqEdit = New System.Windows.Forms.DataGridView()
        Me.GroupBoxFS = New System.Windows.Forms.GroupBox()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBoxMenu = New Focus.myGroupBox()
        Me.btnApproval = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.BtnEdit = New System.Windows.Forms.Button()
        Me.BtnAdd = New System.Windows.Forms.Button()
        CType(Me.DataGridViewMaterialReq, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridViewWHPartNumbers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridViewMatReqEdit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxFS.SuspendLayout()
        Me.GroupBoxMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtMO
        '
        Me.txtMO.Location = New System.Drawing.Point(87, 90)
        Me.txtMO.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMO.MaxLength = 50
        Me.txtMO.Name = "txtMO"
        Me.txtMO.Size = New System.Drawing.Size(217, 20)
        Me.txtMO.TabIndex = 3
        Me.txtMO.TabStop = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(9, 90)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(59, 13)
        Me.Label14.TabIndex = 258
        Me.Label14.Text = "MoNumber"
        '
        'DTPReqDt
        '
        Me.DTPReqDt.Location = New System.Drawing.Point(87, 58)
        Me.DTPReqDt.Margin = New System.Windows.Forms.Padding(2)
        Me.DTPReqDt.Name = "DTPReqDt"
        Me.DTPReqDt.Size = New System.Drawing.Size(126, 20)
        Me.DTPReqDt.TabIndex = 2
        Me.DTPReqDt.TabStop = False
        '
        'btnReqSave
        '
        Me.btnReqSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReqSave.Location = New System.Drawing.Point(970, 382)
        Me.btnReqSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnReqSave.Name = "btnReqSave"
        Me.btnReqSave.Size = New System.Drawing.Size(90, 21)
        Me.btnReqSave.TabIndex = 29
        Me.btnReqSave.Text = "Add Lines"
        Me.btnReqSave.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(14, 379)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(49, 13)
        Me.Label13.TabIndex = 253
        Me.Label13.Text = "Remarks"
        '
        'ComboBoxPur
        '
        Me.ComboBoxPur.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPur.FormattingEnabled = True
        Me.ComboBoxPur.Location = New System.Drawing.Point(929, 345)
        Me.ComboBoxPur.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBoxPur.Name = "ComboBoxPur"
        Me.ComboBoxPur.Size = New System.Drawing.Size(131, 21)
        Me.ComboBoxPur.TabIndex = 17
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(926, 329)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(46, 13)
        Me.Label12.TabIndex = 251
        Me.Label12.Text = "Purpose"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(794, 330)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 13)
        Me.Label11.TabIndex = 250
        Me.Label11.Text = "Used Part[Y/N]"
        '
        'ComboBoxUYN
        '
        Me.ComboBoxUYN.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxUYN.FormattingEnabled = True
        Me.ComboBoxUYN.Items.AddRange(New Object() {"Yes", "No"})
        Me.ComboBoxUYN.Location = New System.Drawing.Point(797, 346)
        Me.ComboBoxUYN.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBoxUYN.Name = "ComboBoxUYN"
        Me.ComboBoxUYN.Size = New System.Drawing.Size(128, 21)
        Me.ComboBoxUYN.TabIndex = 16
        '
        'LabelDetQty
        '
        Me.LabelDetQty.AutoSize = True
        Me.LabelDetQty.Location = New System.Drawing.Point(638, 330)
        Me.LabelDetQty.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LabelDetQty.Name = "LabelDetQty"
        Me.LabelDetQty.Size = New System.Drawing.Size(23, 13)
        Me.LabelDetQty.TabIndex = 248
        Me.LabelDetQty.Text = "Qty"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(351, 329)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(60, 13)
        Me.Label9.TabIndex = 247
        Me.Label9.Text = "Description"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(117, 329)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(66, 13)
        Me.Label8.TabIndex = 246
        Me.Label8.Text = "Part Number"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(14, 329)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(36, 13)
        Me.Label7.TabIndex = 245
        Me.Label7.Text = "Sl.No."
        '
        'txtSlNo
        '
        Me.txtSlNo.Location = New System.Drawing.Point(16, 346)
        Me.txtSlNo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSlNo.Name = "txtSlNo"
        Me.txtSlNo.Size = New System.Drawing.Size(28, 20)
        Me.txtSlNo.TabIndex = 10
        '
        'DataGridViewMaterialReq
        '
        Me.DataGridViewMaterialReq.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.DataGridViewMaterialReq.AllowUserToAddRows = False
        Me.DataGridViewMaterialReq.AllowUserToDeleteRows = False
        Me.DataGridViewMaterialReq.AllowUserToResizeRows = False
        Me.DataGridViewMaterialReq.BackgroundColor = System.Drawing.Color.White
        Me.DataGridViewMaterialReq.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewMaterialReq.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridViewMaterialReq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewMaterialReq.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewMaterialReq.GridColor = System.Drawing.Color.DimGray
        Me.DataGridViewMaterialReq.Location = New System.Drawing.Point(15, 17)
        Me.DataGridViewMaterialReq.Name = "DataGridViewMaterialReq"
        Me.DataGridViewMaterialReq.RowHeadersWidth = 56
        Me.DataGridViewMaterialReq.Size = New System.Drawing.Size(1040, 294)
        Me.DataGridViewMaterialReq.TabIndex = 243
        '
        'ComboBoxSD
        '
        Me.ComboBoxSD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxSD.FormattingEnabled = True
        Me.ComboBoxSD.Location = New System.Drawing.Point(851, 52)
        Me.ComboBoxSD.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBoxSD.Name = "ComboBoxSD"
        Me.ComboBoxSD.Size = New System.Drawing.Size(204, 21)
        Me.ComboBoxSD.TabIndex = 8
        '
        'ComboBoxCell
        '
        Me.ComboBoxCell.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxCell.FormattingEnabled = True
        Me.ComboBoxCell.Location = New System.Drawing.Point(851, 82)
        Me.ComboBoxCell.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBoxCell.Name = "ComboBoxCell"
        Me.ComboBoxCell.Size = New System.Drawing.Size(204, 21)
        Me.ComboBoxCell.TabIndex = 9
        '
        'ComboBoxdept
        '
        Me.ComboBoxdept.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxdept.FormattingEnabled = True
        Me.ComboBoxdept.Location = New System.Drawing.Point(851, 23)
        Me.ComboBoxdept.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBoxdept.Name = "ComboBoxdept"
        Me.ComboBoxdept.Size = New System.Drawing.Size(204, 21)
        Me.ComboBoxdept.TabIndex = 7
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(354, 347)
        Me.txtDescription.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(268, 20)
        Me.txtDescription.TabIndex = 13
        '
        'txtPartNumber
        '
        Me.txtPartNumber.Location = New System.Drawing.Point(117, 347)
        Me.txtPartNumber.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPartNumber.Name = "txtPartNumber"
        Me.txtPartNumber.Size = New System.Drawing.Size(233, 20)
        Me.txtPartNumber.TabIndex = 12
        '
        'txtQty
        '
        Me.txtQty.Location = New System.Drawing.Point(630, 348)
        Me.txtQty.Margin = New System.Windows.Forms.Padding(2)
        Me.txtQty.Name = "txtQty"
        Me.txtQty.Size = New System.Drawing.Size(67, 20)
        Me.txtQty.TabIndex = 14
        '
        'txtDetailRemark
        '
        Me.txtDetailRemark.Location = New System.Drawing.Point(67, 379)
        Me.txtDetailRemark.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDetailRemark.MaxLength = 100
        Me.txtDetailRemark.Name = "txtDetailRemark"
        Me.txtDetailRemark.Size = New System.Drawing.Size(558, 20)
        Me.txtDetailRemark.TabIndex = 18
        '
        'txtReqNo
        '
        Me.txtReqNo.Location = New System.Drawing.Point(87, 31)
        Me.txtReqNo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtReqNo.Name = "txtReqNo"
        Me.txtReqNo.Size = New System.Drawing.Size(126, 20)
        Me.txtReqNo.TabIndex = 1
        Me.txtReqNo.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(780, 81)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Cell"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(780, 55)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Sub Division"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(780, 26)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Department"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 62)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Date"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 31)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Request No."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnSaveChanges)
        Me.GroupBox1.Controls.Add(Me.lblStk)
        Me.GroupBox1.Controls.Add(Me.LblStockavble)
        Me.GroupBox1.Controls.Add(Me.DataGridViewWHPartNumbers)
        Me.GroupBox1.Controls.Add(Me.ComboBoxuom)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.txtDescription)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.txtPartNumber)
        Me.GroupBox1.Controls.Add(Me.ComboBoxIssueType)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.LabelDetQty)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.ComboBoxUYN)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.txtSlNo)
        Me.GroupBox1.Controls.Add(Me.DataGridViewMaterialReq)
        Me.GroupBox1.Controls.Add(Me.txtDetailRemark)
        Me.GroupBox1.Controls.Add(Me.ComboBoxPur)
        Me.GroupBox1.Controls.Add(Me.btnReqSave)
        Me.GroupBox1.Controls.Add(Me.txtQty)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 212)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(1074, 409)
        Me.GroupBox1.TabIndex = 260
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Item Details"
        '
        'btnSaveChanges
        '
        Me.btnSaveChanges.Enabled = False
        Me.btnSaveChanges.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveChanges.Location = New System.Drawing.Point(951, 311)
        Me.btnSaveChanges.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSaveChanges.Name = "btnSaveChanges"
        Me.btnSaveChanges.Size = New System.Drawing.Size(109, 21)
        Me.btnSaveChanges.TabIndex = 267
        Me.btnSaveChanges.Text = "Save Changes"
        Me.btnSaveChanges.UseVisualStyleBackColor = True
        '
        'lblStk
        '
        Me.lblStk.AutoSize = True
        Me.lblStk.Location = New System.Drawing.Point(635, 386)
        Me.lblStk.Name = "lblStk"
        Me.lblStk.Size = New System.Drawing.Size(26, 13)
        Me.lblStk.TabIndex = 266
        Me.lblStk.Text = "Stk:"
        Me.lblStk.Visible = False
        '
        'LblStockavble
        '
        Me.LblStockavble.AutoSize = True
        Me.LblStockavble.Location = New System.Drawing.Point(686, 386)
        Me.LblStockavble.Name = "LblStockavble"
        Me.LblStockavble.Size = New System.Drawing.Size(23, 13)
        Me.LblStockavble.TabIndex = 265
        Me.LblStockavble.Text = "Stk"
        Me.LblStockavble.Visible = False
        '
        'DataGridViewWHPartNumbers
        '
        Me.DataGridViewWHPartNumbers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewWHPartNumbers.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridViewWHPartNumbers.Location = New System.Drawing.Point(372, 28)
        Me.DataGridViewWHPartNumbers.Margin = New System.Windows.Forms.Padding(2)
        Me.DataGridViewWHPartNumbers.Name = "DataGridViewWHPartNumbers"
        Me.DataGridViewWHPartNumbers.RowTemplate.Height = 28
        Me.DataGridViewWHPartNumbers.Size = New System.Drawing.Size(545, 261)
        Me.DataGridViewWHPartNumbers.TabIndex = 264
        Me.DataGridViewWHPartNumbers.Visible = False
        '
        'ComboBoxuom
        '
        Me.ComboBoxuom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxuom.FormattingEnabled = True
        Me.ComboBoxuom.Items.AddRange(New Object() {"EA", "KG", "SET", "MTR", "FT", "LTR"})
        Me.ComboBoxuom.Location = New System.Drawing.Point(706, 346)
        Me.ComboBoxuom.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBoxuom.Name = "ComboBoxuom"
        Me.ComboBoxuom.Size = New System.Drawing.Size(87, 21)
        Me.ComboBoxuom.TabIndex = 15
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(716, 330)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(32, 13)
        Me.Label15.TabIndex = 263
        Me.Label15.Text = "UOM"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(54, 329)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(59, 13)
        Me.Label10.TabIndex = 256
        Me.Label10.Text = "Issue Type"
        '
        'ComboBoxIssueType
        '
        Me.ComboBoxIssueType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxIssueType.FormattingEnabled = True
        Me.ComboBoxIssueType.Items.AddRange(New Object() {"Issue", "Return"})
        Me.ComboBoxIssueType.Location = New System.Drawing.Point(52, 346)
        Me.ComboBoxIssueType.Margin = New System.Windows.Forms.Padding(2)
        Me.ComboBoxIssueType.Name = "ComboBoxIssueType"
        Me.ComboBoxIssueType.Size = New System.Drawing.Size(61, 21)
        Me.ComboBoxIssueType.TabIndex = 11
        '
        'RBFSItem
        '
        Me.RBFSItem.AutoSize = True
        Me.RBFSItem.Location = New System.Drawing.Point(8, 17)
        Me.RBFSItem.Margin = New System.Windows.Forms.Padding(2)
        Me.RBFSItem.Name = "RBFSItem"
        Me.RBFSItem.Size = New System.Drawing.Size(98, 17)
        Me.RBFSItem.TabIndex = 5
        Me.RBFSItem.TabStop = True
        Me.RBFSItem.Text = "FS Item / Tools"
        Me.RBFSItem.UseVisualStyleBackColor = True
        '
        'RBNonFSItem
        '
        Me.RBNonFSItem.AutoSize = True
        Me.RBNonFSItem.Location = New System.Drawing.Point(9, 46)
        Me.RBNonFSItem.Margin = New System.Windows.Forms.Padding(2)
        Me.RBNonFSItem.Name = "RBNonFSItem"
        Me.RBNonFSItem.Size = New System.Drawing.Size(84, 17)
        Me.RBNonFSItem.TabIndex = 6
        Me.RBNonFSItem.TabStop = True
        Me.RBNonFSItem.Text = "Non FS Item"
        Me.RBNonFSItem.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DataGridViewMatReqEdit)
        Me.GroupBox2.Controls.Add(Me.GroupBoxFS)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.txtReqNo)
        Me.GroupBox2.Controls.Add(Me.DTPReqDt)
        Me.GroupBox2.Controls.Add(Me.txtNotes)
        Me.GroupBox2.Controls.Add(Me.ComboBoxdept)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.txtMO)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.ComboBoxCell)
        Me.GroupBox2.Controls.Add(Me.ComboBoxSD)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 51)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(1074, 158)
        Me.GroupBox2.TabIndex = 261
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Material Request "
        '
        'DataGridViewMatReqEdit
        '
        Me.DataGridViewMatReqEdit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewMatReqEdit.Location = New System.Drawing.Point(641, 17)
        Me.DataGridViewMatReqEdit.Margin = New System.Windows.Forms.Padding(2)
        Me.DataGridViewMatReqEdit.Name = "DataGridViewMatReqEdit"
        Me.DataGridViewMatReqEdit.RowTemplate.Height = 28
        Me.DataGridViewMatReqEdit.Size = New System.Drawing.Size(107, 43)
        Me.DataGridViewMatReqEdit.TabIndex = 265
        Me.DataGridViewMatReqEdit.Visible = False
        '
        'GroupBoxFS
        '
        Me.GroupBoxFS.Controls.Add(Me.RBNonFSItem)
        Me.GroupBoxFS.Controls.Add(Me.RBFSItem)
        Me.GroupBoxFS.Location = New System.Drawing.Point(510, 26)
        Me.GroupBoxFS.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxFS.Name = "GroupBoxFS"
        Me.GroupBoxFS.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxFS.Size = New System.Drawing.Size(110, 69)
        Me.GroupBoxFS.TabIndex = 261
        Me.GroupBoxFS.TabStop = False
        Me.GroupBoxFS.Text = "Item"
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(87, 118)
        Me.txtNotes.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNotes.MaxLength = 150
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(968, 20)
        Me.txtNotes.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 118)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Remarks"
        '
        'GroupBoxMenu
        '
        Me.GroupBoxMenu.BackColor = System.Drawing.Color.SteelBlue
        Me.GroupBoxMenu.BorderColor = System.Drawing.Color.Black
        Me.GroupBoxMenu.Controls.Add(Me.btnApproval)
        Me.GroupBoxMenu.Controls.Add(Me.btnDelete)
        Me.GroupBoxMenu.Controls.Add(Me.BtnEdit)
        Me.GroupBoxMenu.Controls.Add(Me.BtnAdd)
        Me.GroupBoxMenu.Location = New System.Drawing.Point(16, 10)
        Me.GroupBoxMenu.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxMenu.Name = "GroupBoxMenu"
        Me.GroupBoxMenu.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxMenu.Size = New System.Drawing.Size(1074, 37)
        Me.GroupBoxMenu.TabIndex = 250
        Me.GroupBoxMenu.TabStop = False
        '
        'btnApproval
        '
        Me.btnApproval.BackColor = System.Drawing.Color.LightGray
        Me.btnApproval.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnApproval.Location = New System.Drawing.Point(193, 6)
        Me.btnApproval.Margin = New System.Windows.Forms.Padding(2)
        Me.btnApproval.Name = "btnApproval"
        Me.btnApproval.Size = New System.Drawing.Size(139, 27)
        Me.btnApproval.TabIndex = 3
        Me.btnApproval.Text = "Send for Approval"
        Me.btnApproval.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.LightGray
        Me.btnDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(132, 6)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(2)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(56, 27)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'BtnEdit
        '
        Me.BtnEdit.BackColor = System.Drawing.Color.LightGray
        Me.BtnEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnEdit.Location = New System.Drawing.Point(71, 6)
        Me.BtnEdit.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(56, 27)
        Me.BtnEdit.TabIndex = 1
        Me.BtnEdit.Text = "Edit"
        Me.BtnEdit.UseVisualStyleBackColor = False
        '
        'BtnAdd
        '
        Me.BtnAdd.BackColor = System.Drawing.Color.LightGray
        Me.BtnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAdd.Location = New System.Drawing.Point(4, 5)
        Me.BtnAdd.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(62, 28)
        Me.BtnAdd.TabIndex = 0
        Me.BtnAdd.Text = "New"
        Me.BtnAdd.UseVisualStyleBackColor = False
        '
        'WHMaterialReq
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1097, 627)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBoxMenu)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "WHMaterialReq"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Material Request"
        CType(Me.DataGridViewMaterialReq, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridViewWHPartNumbers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DataGridViewMatReqEdit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxFS.ResumeLayout(False)
        Me.GroupBoxFS.PerformLayout()
        Me.GroupBoxMenu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBoxMenu As Focus.myGroupBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnAdd As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RBFSItem As System.Windows.Forms.RadioButton
    Friend WithEvents RBNonFSItem As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBoxSD As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxCell As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxdept As System.Windows.Forms.ComboBox
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtPartNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtQty As System.Windows.Forms.TextBox
    Friend WithEvents txtDetailRemark As System.Windows.Forms.TextBox
    Friend WithEvents txtReqNo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtSlNo As System.Windows.Forms.TextBox
    Friend WithEvents DataGridViewMaterialReq As System.Windows.Forms.DataGridView
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxPur As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxUYN As System.Windows.Forms.ComboBox
    Friend WithEvents LabelDetQty As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnApproval As System.Windows.Forms.Button
    Friend WithEvents btnReqSave As System.Windows.Forms.Button
    Friend WithEvents DTPReqDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtMO As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBoxFS As System.Windows.Forms.GroupBox
    Friend WithEvents txtNotes As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxIssueType As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxuom As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents DataGridViewWHPartNumbers As System.Windows.Forms.DataGridView
    Friend WithEvents LblStockavble As System.Windows.Forms.Label
    Friend WithEvents lblStk As System.Windows.Forms.Label
    Friend WithEvents DataGridViewMatReqEdit As System.Windows.Forms.DataGridView
    Friend WithEvents btnSaveChanges As System.Windows.Forms.Button
End Class
