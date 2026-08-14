<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WHMaterialIssue
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.datagridDC = New System.Windows.Forms.DataGridView()
        Me.datagridReqPending = New System.Windows.Forms.DataGridView()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtMatReqNo = New System.Windows.Forms.TextBox()
        Me.txtHeaderNotes = New System.Windows.Forms.TextBox()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.RBNonFSItem = New System.Windows.Forms.RadioButton()
        Me.RBFSItem = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtMatIssNo = New System.Windows.Forms.TextBox()
        Me.DTPIssDt = New System.Windows.Forms.DateTimePicker()
        Me.txtRemarks = New System.Windows.Forms.TextBox()
        Me.ComboBoxdept = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtMO = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ComboBoxCell = New System.Windows.Forms.ComboBox()
        Me.ComboBoxSD = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblNote = New System.Windows.Forms.Label()
        Me.DataGridViewMaterialReq = New System.Windows.Forms.DataGridView()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.MyGroupBox1 = New Focus.myGroupBox()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.GroupBoxMenu = New Focus.myGroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.BtnEdit = New System.Windows.Forms.Button()
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.GroupBox2.SuspendLayout()
        CType(Me.datagridDC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.datagridReqPending, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridViewMaterialReq, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MyGroupBox1.SuspendLayout()
        Me.GroupBoxMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.datagridDC)
        Me.GroupBox2.Controls.Add(Me.datagridReqPending)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.txtMatReqNo)
        Me.GroupBox2.Controls.Add(Me.txtHeaderNotes)
        Me.GroupBox2.Controls.Add(Me.lblNotes)
        Me.GroupBox2.Controls.Add(Me.GroupBox4)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.txtMatIssNo)
        Me.GroupBox2.Controls.Add(Me.DTPIssDt)
        Me.GroupBox2.Controls.Add(Me.txtRemarks)
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
        Me.GroupBox2.Location = New System.Drawing.Point(31, 55)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2, 1, 2, 1)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2, 1, 2, 1)
        Me.GroupBox2.Size = New System.Drawing.Size(1039, 162)
        Me.GroupBox2.TabIndex = 264
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Material Issue"
        '
        'datagridDC
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.datagridDC.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.datagridDC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.datagridDC.DefaultCellStyle = DataGridViewCellStyle2
        Me.datagridDC.Location = New System.Drawing.Point(394, 34)
        Me.datagridDC.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.datagridDC.Name = "datagridDC"
        Me.datagridDC.Size = New System.Drawing.Size(98, 46)
        Me.datagridDC.TabIndex = 268
        Me.datagridDC.Visible = False
        '
        'datagridReqPending
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.datagridReqPending.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.datagridReqPending.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.datagridReqPending.DefaultCellStyle = DataGridViewCellStyle4
        Me.datagridReqPending.Location = New System.Drawing.Point(468, 10)
        Me.datagridReqPending.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.datagridReqPending.Name = "datagridReqPending"
        Me.datagridReqPending.Size = New System.Drawing.Size(151, 18)
        Me.datagridReqPending.TabIndex = 264
        Me.datagridReqPending.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(645, 11)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(68, 13)
        Me.Label8.TabIndex = 266
        Me.Label8.Text = "Mat Req No."
        '
        'txtMatReqNo
        '
        Me.txtMatReqNo.Location = New System.Drawing.Point(818, 11)
        Me.txtMatReqNo.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtMatReqNo.Name = "txtMatReqNo"
        Me.txtMatReqNo.Size = New System.Drawing.Size(203, 20)
        Me.txtMatReqNo.TabIndex = 265
        '
        'txtHeaderNotes
        '
        Me.txtHeaderNotes.Location = New System.Drawing.Point(105, 138)
        Me.txtHeaderNotes.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtHeaderNotes.MaxLength = 150
        Me.txtHeaderNotes.Name = "txtHeaderNotes"
        Me.txtHeaderNotes.Size = New System.Drawing.Size(914, 20)
        Me.txtHeaderNotes.TabIndex = 263
        '
        'lblNotes
        '
        Me.lblNotes.AutoSize = True
        Me.lblNotes.Location = New System.Drawing.Point(9, 140)
        Me.lblNotes.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(35, 13)
        Me.lblNotes.TabIndex = 262
        Me.lblNotes.Text = "Notes"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.RBNonFSItem)
        Me.GroupBox4.Controls.Add(Me.RBFSItem)
        Me.GroupBox4.Location = New System.Drawing.Point(510, 31)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(2, 1, 2, 1)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(2, 1, 2, 1)
        Me.GroupBox4.Size = New System.Drawing.Size(109, 69)
        Me.GroupBox4.TabIndex = 261
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Item"
        '
        'RBNonFSItem
        '
        Me.RBNonFSItem.AutoSize = True
        Me.RBNonFSItem.Location = New System.Drawing.Point(9, 46)
        Me.RBNonFSItem.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.RBNonFSItem.Name = "RBNonFSItem"
        Me.RBNonFSItem.Size = New System.Drawing.Size(84, 17)
        Me.RBNonFSItem.TabIndex = 2
        Me.RBNonFSItem.TabStop = True
        Me.RBNonFSItem.Text = "Non FS Item"
        Me.RBNonFSItem.UseVisualStyleBackColor = True
        '
        'RBFSItem
        '
        Me.RBFSItem.AutoSize = True
        Me.RBFSItem.Location = New System.Drawing.Point(9, 17)
        Me.RBFSItem.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.RBFSItem.Name = "RBFSItem"
        Me.RBFSItem.Size = New System.Drawing.Size(98, 17)
        Me.RBFSItem.TabIndex = 3
        Me.RBFSItem.TabStop = True
        Me.RBFSItem.Text = "FS Item / Tools"
        Me.RBFSItem.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 31)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Issue / DC No."
        '
        'txtMatIssNo
        '
        Me.txtMatIssNo.Location = New System.Drawing.Point(105, 31)
        Me.txtMatIssNo.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtMatIssNo.Name = "txtMatIssNo"
        Me.txtMatIssNo.Size = New System.Drawing.Size(125, 20)
        Me.txtMatIssNo.TabIndex = 10
        '
        'DTPIssDt
        '
        Me.DTPIssDt.Location = New System.Drawing.Point(105, 59)
        Me.DTPIssDt.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.DTPIssDt.Name = "DTPIssDt"
        Me.DTPIssDt.Size = New System.Drawing.Size(125, 20)
        Me.DTPIssDt.TabIndex = 257
        '
        'txtRemarks
        '
        Me.txtRemarks.Location = New System.Drawing.Point(105, 116)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(914, 20)
        Me.txtRemarks.TabIndex = 12
        '
        'ComboBoxdept
        '
        Me.ComboBoxdept.FormattingEnabled = True
        Me.ComboBoxdept.Location = New System.Drawing.Point(818, 38)
        Me.ComboBoxdept.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.ComboBoxdept.Name = "ComboBoxdept"
        Me.ComboBoxdept.Size = New System.Drawing.Size(201, 21)
        Me.ComboBoxdept.TabIndex = 17
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
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(647, 38)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Department"
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
        'txtMO
        '
        Me.txtMO.Location = New System.Drawing.Point(105, 90)
        Me.txtMO.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtMO.Name = "txtMO"
        Me.txtMO.Size = New System.Drawing.Size(217, 20)
        Me.txtMO.TabIndex = 259
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(9, 90)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(62, 13)
        Me.Label14.TabIndex = 258
        Me.Label14.Text = "Mo Number"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(647, 83)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Cell"
        '
        'ComboBoxCell
        '
        Me.ComboBoxCell.FormattingEnabled = True
        Me.ComboBoxCell.Location = New System.Drawing.Point(818, 87)
        Me.ComboBoxCell.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.ComboBoxCell.Name = "ComboBoxCell"
        Me.ComboBoxCell.Size = New System.Drawing.Size(201, 21)
        Me.ComboBoxCell.TabIndex = 18
        '
        'ComboBoxSD
        '
        Me.ComboBoxSD.FormattingEnabled = True
        Me.ComboBoxSD.Location = New System.Drawing.Point(818, 62)
        Me.ComboBoxSD.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.ComboBoxSD.Name = "ComboBoxSD"
        Me.ComboBoxSD.Size = New System.Drawing.Size(201, 21)
        Me.ComboBoxSD.TabIndex = 19
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(647, 62)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Sub Division"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblNote)
        Me.GroupBox1.Controls.Add(Me.DataGridViewMaterialReq)
        Me.GroupBox1.Controls.Add(Me.btnSave)
        Me.GroupBox1.Location = New System.Drawing.Point(31, 221)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 1, 2, 1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 1, 2, 1)
        Me.GroupBox1.Size = New System.Drawing.Size(1039, 376)
        Me.GroupBox1.TabIndex = 263
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Item Details"
        '
        'lblNote
        '
        Me.lblNote.AutoSize = True
        Me.lblNote.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblNote.ForeColor = System.Drawing.Color.Black
        Me.lblNote.Location = New System.Drawing.Point(9, 351)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.Size = New System.Drawing.Size(19, 13)
        Me.lblNote.TabIndex = 270
        Me.lblNote.Text = "La"
        Me.lblNote.Visible = False
        '
        'DataGridViewMaterialReq
        '
        Me.DataGridViewMaterialReq.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.DataGridViewMaterialReq.AllowUserToAddRows = False
        Me.DataGridViewMaterialReq.AllowUserToDeleteRows = False
        Me.DataGridViewMaterialReq.AllowUserToResizeRows = False
        Me.DataGridViewMaterialReq.BackgroundColor = System.Drawing.Color.White
        Me.DataGridViewMaterialReq.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewMaterialReq.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.DataGridViewMaterialReq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewMaterialReq.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewMaterialReq.GridColor = System.Drawing.Color.DimGray
        Me.DataGridViewMaterialReq.Location = New System.Drawing.Point(15, 17)
        Me.DataGridViewMaterialReq.Name = "DataGridViewMaterialReq"
        Me.DataGridViewMaterialReq.RowHeadersWidth = 56
        Me.DataGridViewMaterialReq.Size = New System.Drawing.Size(1006, 328)
        Me.DataGridViewMaterialReq.TabIndex = 243
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(945, 350)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(90, 21)
        Me.btnSave.TabIndex = 254
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'MyGroupBox1
        '
        Me.MyGroupBox1.BackColor = System.Drawing.Color.SteelBlue
        Me.MyGroupBox1.BorderColor = System.Drawing.Color.Black
        Me.MyGroupBox1.Controls.Add(Me.btnPrint)
        Me.MyGroupBox1.Controls.Add(Me.btnCancel)
        Me.MyGroupBox1.Controls.Add(Me.Button3)
        Me.MyGroupBox1.Controls.Add(Me.Button4)
        Me.MyGroupBox1.Location = New System.Drawing.Point(31, 6)
        Me.MyGroupBox1.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.MyGroupBox1.Name = "MyGroupBox1"
        Me.MyGroupBox1.Padding = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.MyGroupBox1.Size = New System.Drawing.Size(1039, 38)
        Me.MyGroupBox1.TabIndex = 262
        Me.MyGroupBox1.TabStop = False
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.LightGray
        Me.btnPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(193, 6)
        Me.btnPrint.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(62, 27)
        Me.btnPrint.TabIndex = 3
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.LightGray
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(132, 6)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 27)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel "
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.LightGray
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(71, 6)
        Me.Button3.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(56, 27)
        Me.Button3.TabIndex = 1
        Me.Button3.Text = "Edit"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.LightGray
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(5, 5)
        Me.Button4.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(62, 27)
        Me.Button4.TabIndex = 0
        Me.Button4.Text = "New"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'GroupBoxMenu
        '
        Me.GroupBoxMenu.BorderColor = System.Drawing.Color.Black
        Me.GroupBoxMenu.Controls.Add(Me.Button2)
        Me.GroupBoxMenu.Controls.Add(Me.btnDelete)
        Me.GroupBoxMenu.Controls.Add(Me.BtnEdit)
        Me.GroupBoxMenu.Controls.Add(Me.BtnAdd)
        Me.GroupBoxMenu.Location = New System.Drawing.Point(-127, -75)
        Me.GroupBoxMenu.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBoxMenu.Name = "GroupBoxMenu"
        Me.GroupBoxMenu.Padding = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBoxMenu.Size = New System.Drawing.Size(839, 49)
        Me.GroupBoxMenu.TabIndex = 252
        Me.GroupBoxMenu.TabStop = False
        Me.GroupBoxMenu.Text = "Menu"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(738, 23)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(101, 27)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Send for Approval"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(115, 17)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(56, 27)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Location = New System.Drawing.Point(62, 17)
        Me.BtnEdit.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(56, 27)
        Me.BtnEdit.TabIndex = 1
        Me.BtnEdit.Text = "Edit"
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnAdd
        '
        Me.BtnAdd.Location = New System.Drawing.Point(5, 17)
        Me.BtnAdd.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(62, 27)
        Me.BtnAdd.TabIndex = 0
        Me.BtnAdd.Text = "Add"
        Me.BtnAdd.UseVisualStyleBackColor = True
        '
        'WHMaterialIssue
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1101, 596)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.MyGroupBox1)
        Me.Controls.Add(Me.GroupBoxMenu)
        Me.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.Name = "WHMaterialIssue"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Material Issues"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.datagridDC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.datagridReqPending, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridViewMaterialReq, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MyGroupBox1.ResumeLayout(False)
        Me.GroupBoxMenu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnAdd As System.Windows.Forms.Button
    Friend WithEvents GroupBoxMenu As Focus.myGroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents RBNonFSItem As System.Windows.Forms.RadioButton
    Friend WithEvents RBFSItem As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtMatIssNo As System.Windows.Forms.TextBox
    Friend WithEvents DTPIssDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents ComboBoxdept As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtMO As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxCell As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxSD As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridViewMaterialReq As System.Windows.Forms.DataGridView
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents MyGroupBox1 As Focus.myGroupBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents txtHeaderNotes As System.Windows.Forms.TextBox
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents datagridReqPending As System.Windows.Forms.DataGridView
    Friend WithEvents txtMatReqNo As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblNote As System.Windows.Forms.Label
    Friend WithEvents datagridDC As System.Windows.Forms.DataGridView
End Class
