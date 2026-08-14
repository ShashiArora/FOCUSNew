<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WHMatReceiptMfgBldg
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.datagridDC = New System.Windows.Forms.DataGridView()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.dtpissdt = New System.Windows.Forms.DateTimePicker()
        Me.txtrem = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.dtpMatReqDt = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtFeedback = New System.Windows.Forms.TextBox()
        Me.txtMatReqNo = New System.Windows.Forms.TextBox()
        Me.lblScan = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtRecNo = New System.Windows.Forms.TextBox()
        Me.txtMatIssueNoImage = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtpReceiptdt = New System.Windows.Forms.DateTimePicker()
        Me.txtMatIssueNo = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnImageClear = New System.Windows.Forms.Button()
        Me.GroupBoxEdit = New System.Windows.Forms.GroupBox()
        Me.DataGridViewReceiptsMfg = New System.Windows.Forms.DataGridView()
        Me.btnRecAccept = New System.Windows.Forms.Button()
        Me.GroupBoxMenu = New Focus.myGroupBox()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.BtnEdit = New System.Windows.Forms.Button()
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.datagridDC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxEdit.SuspendLayout()
        CType(Me.DataGridViewReceiptsMfg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.datagridDC)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.dtpissdt)
        Me.GroupBox1.Controls.Add(Me.txtrem)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.dtpMatReqDt)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtFeedback)
        Me.GroupBox1.Controls.Add(Me.txtMatReqNo)
        Me.GroupBox1.Controls.Add(Me.lblScan)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtRecNo)
        Me.GroupBox1.Controls.Add(Me.txtMatIssueNoImage)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.dtpReceiptdt)
        Me.GroupBox1.Controls.Add(Me.txtMatIssueNo)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.btnImageClear)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 75)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1594, 217)
        Me.GroupBox1.TabIndex = 261
        Me.GroupBox1.TabStop = False
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
        Me.datagridDC.Location = New System.Drawing.Point(1300, 14)
        Me.datagridDC.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.datagridDC.Name = "datagridDC"
        Me.datagridDC.Size = New System.Drawing.Size(68, 71)
        Me.datagridDC.TabIndex = 267
        Me.datagridDC.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(360, 77)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(30, 20)
        Me.Label9.TabIndex = 263
        Me.Label9.Text = "Dt."
        '
        'dtpissdt
        '
        Me.dtpissdt.Location = New System.Drawing.Point(398, 74)
        Me.dtpissdt.Name = "dtpissdt"
        Me.dtpissdt.Size = New System.Drawing.Size(164, 26)
        Me.dtpissdt.TabIndex = 262
        '
        'txtrem
        '
        Me.txtrem.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtrem.Location = New System.Drawing.Point(744, 114)
        Me.txtrem.Name = "txtrem"
        Me.txtrem.Size = New System.Drawing.Size(817, 32)
        Me.txtrem.TabIndex = 261
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(612, 114)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(51, 20)
        Me.Label8.TabIndex = 260
        Me.Label8.Text = "Notes"
        '
        'dtpMatReqDt
        '
        Me.dtpMatReqDt.Location = New System.Drawing.Point(1008, 66)
        Me.dtpMatReqDt.Name = "dtpMatReqDt"
        Me.dtpMatReqDt.Size = New System.Drawing.Size(212, 26)
        Me.dtpMatReqDt.TabIndex = 259
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(909, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(91, 20)
        Me.Label3.TabIndex = 258
        Me.Label3.Text = "Mat Req Dt"
        '
        'txtFeedback
        '
        Me.txtFeedback.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFeedback.Location = New System.Drawing.Point(156, 157)
        Me.txtFeedback.MaxLength = 250
        Me.txtFeedback.Multiline = True
        Me.txtFeedback.Name = "txtFeedback"
        Me.txtFeedback.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtFeedback.Size = New System.Drawing.Size(1405, 38)
        Me.txtFeedback.TabIndex = 257
        '
        'txtMatReqNo
        '
        Me.txtMatReqNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMatReqNo.Location = New System.Drawing.Point(746, 68)
        Me.txtMatReqNo.Name = "txtMatReqNo"
        Me.txtMatReqNo.Size = New System.Drawing.Size(156, 32)
        Me.txtMatReqNo.TabIndex = 253
        '
        'lblScan
        '
        Me.lblScan.AutoSize = True
        Me.lblScan.Location = New System.Drawing.Point(6, 23)
        Me.lblScan.Name = "lblScan"
        Me.lblScan.Size = New System.Drawing.Size(122, 20)
        Me.lblScan.TabIndex = 0
        Me.lblScan.Text = "Scan the Image"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(612, 68)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(98, 20)
        Me.Label7.TabIndex = 252
        Me.Label7.Text = "Mat Req No."
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 157)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(84, 20)
        Me.Label5.TabIndex = 256
        Me.Label5.Text = "Feed back"
        '
        'txtRecNo
        '
        Me.txtRecNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRecNo.Location = New System.Drawing.Point(746, 17)
        Me.txtRecNo.Name = "txtRecNo"
        Me.txtRecNo.Size = New System.Drawing.Size(156, 32)
        Me.txtRecNo.TabIndex = 251
        '
        'txtMatIssueNoImage
        '
        Me.txtMatIssueNoImage.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMatIssueNoImage.Location = New System.Drawing.Point(156, 17)
        Me.txtMatIssueNoImage.Name = "txtMatIssueNoImage"
        Me.txtMatIssueNoImage.Size = New System.Drawing.Size(302, 39)
        Me.txtMatIssueNoImage.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(612, 17)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(127, 20)
        Me.Label6.TabIndex = 250
        Me.Label6.Text = "Mat. Receipt No."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(44, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(33, 20)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "OR"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 85)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(138, 20)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Mat. Issue/DC No."
        '
        'dtpReceiptdt
        '
        Me.dtpReceiptdt.Location = New System.Drawing.Point(1008, 14)
        Me.dtpReceiptdt.Name = "dtpReceiptdt"
        Me.dtpReceiptdt.Size = New System.Drawing.Size(212, 26)
        Me.dtpReceiptdt.TabIndex = 247
        '
        'txtMatIssueNo
        '
        Me.txtMatIssueNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMatIssueNo.Location = New System.Drawing.Point(156, 77)
        Me.txtMatIssueNo.Name = "txtMatIssueNo"
        Me.txtMatIssueNo.Size = New System.Drawing.Size(126, 32)
        Me.txtMatIssueNo.TabIndex = 248
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(909, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 20)
        Me.Label4.TabIndex = 246
        Me.Label4.Text = "Date"
        '
        'btnImageClear
        '
        Me.btnImageClear.Location = New System.Drawing.Point(466, 18)
        Me.btnImageClear.Name = "btnImageClear"
        Me.btnImageClear.Size = New System.Drawing.Size(98, 34)
        Me.btnImageClear.TabIndex = 2
        Me.btnImageClear.Text = "&OK"
        Me.btnImageClear.UseVisualStyleBackColor = True
        '
        'GroupBoxEdit
        '
        Me.GroupBoxEdit.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GroupBoxEdit.Controls.Add(Me.DataGridViewReceiptsMfg)
        Me.GroupBoxEdit.Controls.Add(Me.btnRecAccept)
        Me.GroupBoxEdit.Location = New System.Drawing.Point(10, 300)
        Me.GroupBoxEdit.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBoxEdit.Name = "GroupBoxEdit"
        Me.GroupBoxEdit.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBoxEdit.Size = New System.Drawing.Size(1594, 657)
        Me.GroupBoxEdit.TabIndex = 259
        Me.GroupBoxEdit.TabStop = False
        Me.GroupBoxEdit.Text = "Item Details"
        '
        'DataGridViewReceiptsMfg
        '
        Me.DataGridViewReceiptsMfg.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.DataGridViewReceiptsMfg.AllowUserToAddRows = False
        Me.DataGridViewReceiptsMfg.AllowUserToDeleteRows = False
        Me.DataGridViewReceiptsMfg.AllowUserToResizeRows = False
        Me.DataGridViewReceiptsMfg.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.DataGridViewReceiptsMfg.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewReceiptsMfg.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridViewReceiptsMfg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewReceiptsMfg.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewReceiptsMfg.GridColor = System.Drawing.Color.DimGray
        Me.DataGridViewReceiptsMfg.Location = New System.Drawing.Point(10, 29)
        Me.DataGridViewReceiptsMfg.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DataGridViewReceiptsMfg.Name = "DataGridViewReceiptsMfg"
        Me.DataGridViewReceiptsMfg.RowHeadersWidth = 56
        Me.DataGridViewReceiptsMfg.Size = New System.Drawing.Size(1575, 586)
        Me.DataGridViewReceiptsMfg.TabIndex = 245
        Me.DataGridViewReceiptsMfg.Visible = False
        '
        'btnRecAccept
        '
        Me.btnRecAccept.Location = New System.Drawing.Point(1464, 623)
        Me.btnRecAccept.Name = "btnRecAccept"
        Me.btnRecAccept.Size = New System.Drawing.Size(122, 32)
        Me.btnRecAccept.TabIndex = 243
        Me.btnRecAccept.Text = "Accept"
        Me.btnRecAccept.UseVisualStyleBackColor = True
        '
        'GroupBoxMenu
        '
        Me.GroupBoxMenu.BackColor = System.Drawing.Color.SteelBlue
        Me.GroupBoxMenu.BorderColor = System.Drawing.Color.Black
        Me.GroupBoxMenu.Controls.Add(Me.btnDelete)
        Me.GroupBoxMenu.Controls.Add(Me.BtnEdit)
        Me.GroupBoxMenu.Controls.Add(Me.BtnAdd)
        Me.GroupBoxMenu.Location = New System.Drawing.Point(12, 8)
        Me.GroupBoxMenu.Name = "GroupBoxMenu"
        Me.GroupBoxMenu.Size = New System.Drawing.Size(1593, 60)
        Me.GroupBoxMenu.TabIndex = 260
        Me.GroupBoxMenu.TabStop = False
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(198, 12)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(84, 42)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'BtnEdit
        '
        Me.BtnEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnEdit.Location = New System.Drawing.Point(106, 12)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(84, 42)
        Me.BtnEdit.TabIndex = 1
        Me.BtnEdit.Text = "Edit"
        Me.BtnEdit.UseVisualStyleBackColor = True
        '
        'BtnAdd
        '
        Me.BtnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAdd.Location = New System.Drawing.Point(6, 11)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(93, 43)
        Me.BtnAdd.TabIndex = 0
        Me.BtnAdd.Text = "New"
        Me.BtnAdd.UseVisualStyleBackColor = True
        '
        'WHMatReceiptMfgBldg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1622, 958)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBoxMenu)
        Me.Controls.Add(Me.GroupBoxEdit)
        Me.Name = "WHMatReceiptMfgBldg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Material Receipt  at Mfg Bldg"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.datagridDC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxEdit.ResumeLayout(False)
        CType(Me.DataGridViewReceiptsMfg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxMenu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtFeedback As System.Windows.Forms.TextBox
    Friend WithEvents txtMatReqNo As System.Windows.Forms.TextBox
    Friend WithEvents lblScan As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtRecNo As System.Windows.Forms.TextBox
    Friend WithEvents txtMatIssueNoImage As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpReceiptdt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtMatIssueNo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnImageClear As System.Windows.Forms.Button
    Friend WithEvents GroupBoxMenu As Focus.myGroupBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnAdd As System.Windows.Forms.Button
    Friend WithEvents GroupBoxEdit As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridViewReceiptsMfg As System.Windows.Forms.DataGridView
    Friend WithEvents btnRecAccept As System.Windows.Forms.Button
    Friend WithEvents dtpMatReqDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtrem As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents dtpissdt As System.Windows.Forms.DateTimePicker
    Friend WithEvents datagridDC As System.Windows.Forms.DataGridView
End Class
