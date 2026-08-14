<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RFQPriceViewQuick
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
        Me.DataGridViewPriceView = New System.Windows.Forms.DataGridView()
        Me.ButtonRefresh = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.lblCustName = New System.Windows.Forms.Label()
        Me.txtPartNumber = New System.Windows.Forms.TextBox()
        Me.RadioButtonPartNumber = New System.Windows.Forms.RadioButton()
        Me.txtCustID = New System.Windows.Forms.TextBox()
        Me.RadioButtonId = New System.Windows.Forms.RadioButton()
        Me.RadioButtonName = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.RadioButtonPurNeed = New System.Windows.Forms.RadioButton()
        Me.txtRegNo = New System.Windows.Forms.TextBox()
        Me.RadioButtonSingle = New System.Windows.Forms.RadioButton()
        Me.RadioButtonQuotePending = New System.Windows.Forms.RadioButton()
        Me.RadioButtonPendingPrice = New System.Windows.Forms.RadioButton()
        Me.RadioButtonAll = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.dtptodate = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpfrdate = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Note = New System.Windows.Forms.Label()
        Me.DataGridPartNumbers = New System.Windows.Forms.DataGrid()
        Me.DataGridCustomer1 = New System.Windows.Forms.DataGrid()
        CType(Me.DataGridViewPriceView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridPartNumbers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridCustomer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridViewPriceView
        '
        Me.DataGridViewPriceView.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.DataGridViewPriceView.AllowDrop = True
        Me.DataGridViewPriceView.AllowUserToOrderColumns = True
        Me.DataGridViewPriceView.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DataGridViewPriceView.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewPriceView.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridViewPriceView.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewPriceView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridViewPriceView.Enabled = False
        Me.DataGridViewPriceView.GridColor = System.Drawing.Color.Red
        Me.DataGridViewPriceView.Location = New System.Drawing.Point(15, 95)
        Me.DataGridViewPriceView.Name = "DataGridViewPriceView"
        Me.DataGridViewPriceView.RightToLeft = System.Windows.Forms.RightToLeft.No
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewPriceView.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridViewPriceView.Size = New System.Drawing.Size(1301, 663)
        Me.DataGridViewPriceView.TabIndex = 176
        '
        'ButtonRefresh
        '
        Me.ButtonRefresh.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ButtonRefresh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRefresh.ForeColor = System.Drawing.Color.White
        Me.ButtonRefresh.Location = New System.Drawing.Point(1241, 72)
        Me.ButtonRefresh.Name = "ButtonRefresh"
        Me.ButtonRefresh.Size = New System.Drawing.Size(75, 24)
        Me.ButtonRefresh.TabIndex = 177
        Me.ButtonRefresh.Text = "Refresh"
        Me.ButtonRefresh.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.RadioButtonPurNeed)
        Me.GroupBox1.Controls.Add(Me.txtRegNo)
        Me.GroupBox1.Controls.Add(Me.RadioButtonSingle)
        Me.GroupBox1.Controls.Add(Me.RadioButtonQuotePending)
        Me.GroupBox1.Controls.Add(Me.RadioButtonPendingPrice)
        Me.GroupBox1.Controls.Add(Me.RadioButtonAll)
        Me.GroupBox1.Location = New System.Drawing.Point(218, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1098, 70)
        Me.GroupBox1.TabIndex = 186
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select any one option"
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox5.Controls.Add(Me.lblCustName)
        Me.GroupBox5.Controls.Add(Me.txtPartNumber)
        Me.GroupBox5.Controls.Add(Me.RadioButtonPartNumber)
        Me.GroupBox5.Controls.Add(Me.txtCustID)
        Me.GroupBox5.Controls.Add(Me.RadioButtonId)
        Me.GroupBox5.Controls.Add(Me.RadioButtonName)
        Me.GroupBox5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox5.ForeColor = System.Drawing.Color.Black
        Me.GroupBox5.Location = New System.Drawing.Point(509, 11)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(583, 53)
        Me.GroupBox5.TabIndex = 205
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Select  Customer  OR Partnumber wise Search"
        '
        'lblCustName
        '
        Me.lblCustName.AutoSize = True
        Me.lblCustName.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCustName.Location = New System.Drawing.Point(143, 37)
        Me.lblCustName.Name = "lblCustName"
        Me.lblCustName.Size = New System.Drawing.Size(12, 16)
        Me.lblCustName.TabIndex = 206
        Me.lblCustName.Text = "-"
        '
        'txtPartNumber
        '
        Me.txtPartNumber.Location = New System.Drawing.Point(392, 14)
        Me.txtPartNumber.Name = "txtPartNumber"
        Me.txtPartNumber.Size = New System.Drawing.Size(182, 20)
        Me.txtPartNumber.TabIndex = 205
        '
        'RadioButtonPartNumber
        '
        Me.RadioButtonPartNumber.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonPartNumber.Location = New System.Drawing.Point(328, 8)
        Me.RadioButtonPartNumber.Name = "RadioButtonPartNumber"
        Me.RadioButtonPartNumber.Size = New System.Drawing.Size(69, 36)
        Me.RadioButtonPartNumber.TabIndex = 204
        Me.RadioButtonPartNumber.TabStop = True
        Me.RadioButtonPartNumber.Text = "Part Number"
        Me.RadioButtonPartNumber.UseVisualStyleBackColor = True
        '
        'txtCustID
        '
        Me.txtCustID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustID.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustID.Location = New System.Drawing.Point(146, 15)
        Me.txtCustID.MaxLength = 30
        Me.txtCustID.Name = "txtCustID"
        Me.txtCustID.Size = New System.Drawing.Size(165, 20)
        Me.txtCustID.TabIndex = 203
        '
        'RadioButtonId
        '
        Me.RadioButtonId.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold)
        Me.RadioButtonId.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonId.Location = New System.Drawing.Point(6, 15)
        Me.RadioButtonId.Name = "RadioButtonId"
        Me.RadioButtonId.Size = New System.Drawing.Size(71, 22)
        Me.RadioButtonId.TabIndex = 52
        Me.RadioButtonId.Text = "Cust Id"
        '
        'RadioButtonName
        '
        Me.RadioButtonName.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold)
        Me.RadioButtonName.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonName.Location = New System.Drawing.Point(83, 17)
        Me.RadioButtonName.Name = "RadioButtonName"
        Me.RadioButtonName.Size = New System.Drawing.Size(64, 17)
        Me.RadioButtonName.TabIndex = 53
        Me.RadioButtonName.Text = "Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label3.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(271, 44)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 16)
        Me.Label3.TabIndex = 192
        Me.Label3.Text = "(Price not given)"
        '
        'RadioButtonPurNeed
        '
        Me.RadioButtonPurNeed.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonPurNeed.Location = New System.Drawing.Point(258, 9)
        Me.RadioButtonPurNeed.Name = "RadioButtonPurNeed"
        Me.RadioButtonPurNeed.Size = New System.Drawing.Size(90, 39)
        Me.RadioButtonPurNeed.TabIndex = 191
        Me.RadioButtonPurNeed.TabStop = True
        Me.RadioButtonPurNeed.Tag = "(Purchase Feed-back)"
        Me.RadioButtonPurNeed.Text = "Pur && Apl Feed-back"
        Me.RadioButtonPurNeed.UseVisualStyleBackColor = True
        '
        'txtRegNo
        '
        Me.txtRegNo.Location = New System.Drawing.Point(426, 19)
        Me.txtRegNo.Name = "txtRegNo"
        Me.txtRegNo.Size = New System.Drawing.Size(77, 20)
        Me.txtRegNo.TabIndex = 190
        Me.txtRegNo.Visible = False
        '
        'RadioButtonSingle
        '
        Me.RadioButtonSingle.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonSingle.Location = New System.Drawing.Point(354, 13)
        Me.RadioButtonSingle.Name = "RadioButtonSingle"
        Me.RadioButtonSingle.Size = New System.Drawing.Size(73, 47)
        Me.RadioButtonSingle.TabIndex = 189
        Me.RadioButtonSingle.TabStop = True
        Me.RadioButtonSingle.Text = "Single Reg No. Status"
        Me.RadioButtonSingle.UseVisualStyleBackColor = True
        '
        'RadioButtonQuotePending
        '
        Me.RadioButtonQuotePending.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonQuotePending.Location = New System.Drawing.Point(129, 17)
        Me.RadioButtonQuotePending.Name = "RadioButtonQuotePending"
        Me.RadioButtonQuotePending.Size = New System.Drawing.Size(123, 36)
        Me.RadioButtonQuotePending.TabIndex = 188
        Me.RadioButtonQuotePending.TabStop = True
        Me.RadioButtonQuotePending.Text = "Price Recd Quote Pending"
        Me.RadioButtonQuotePending.UseVisualStyleBackColor = True
        '
        'RadioButtonPendingPrice
        '
        Me.RadioButtonPendingPrice.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonPendingPrice.Location = New System.Drawing.Point(51, 14)
        Me.RadioButtonPendingPrice.Name = "RadioButtonPendingPrice"
        Me.RadioButtonPendingPrice.Size = New System.Drawing.Size(82, 39)
        Me.RadioButtonPendingPrice.TabIndex = 187
        Me.RadioButtonPendingPrice.TabStop = True
        Me.RadioButtonPendingPrice.Text = "Pending for Price "
        Me.RadioButtonPendingPrice.UseVisualStyleBackColor = True
        '
        'RadioButtonAll
        '
        Me.RadioButtonAll.AutoSize = True
        Me.RadioButtonAll.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonAll.Location = New System.Drawing.Point(6, 23)
        Me.RadioButtonAll.Name = "RadioButtonAll"
        Me.RadioButtonAll.Size = New System.Drawing.Size(39, 20)
        Me.RadioButtonAll.TabIndex = 186
        Me.RadioButtonAll.TabStop = True
        Me.RadioButtonAll.Text = "All"
        Me.RadioButtonAll.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.dtptodate)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.dtpfrdate)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 70)
        Me.GroupBox2.TabIndex = 187
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Select Enq Reg Date"
        '
        'dtptodate
        '
        Me.dtptodate.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dtptodate.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtptodate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtptodate.Location = New System.Drawing.Point(50, 43)
        Me.dtptodate.Name = "dtptodate"
        Me.dtptodate.Size = New System.Drawing.Size(118, 22)
        Me.dtptodate.TabIndex = 185
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 16)
        Me.Label2.TabIndex = 184
        Me.Label2.Text = "To Dt."
        '
        'dtpfrdate
        '
        Me.dtpfrdate.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dtpfrdate.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpfrdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpfrdate.Location = New System.Drawing.Point(50, 16)
        Me.dtpfrdate.Name = "dtpfrdate"
        Me.dtpfrdate.Size = New System.Drawing.Size(118, 22)
        Me.dtpfrdate.TabIndex = 183
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 16)
        Me.Label1.TabIndex = 182
        Me.Label1.Text = "Fr Dt:"
        '
        'Note
        '
        Me.Note.AutoSize = True
        Me.Note.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Note.ForeColor = System.Drawing.Color.Red
        Me.Note.Location = New System.Drawing.Point(18, 76)
        Me.Note.Name = "Note"
        Me.Note.Size = New System.Drawing.Size(505, 16)
        Me.Note.TabIndex = 188
        Me.Note.Text = "NOTE : Certificate charges will  NOT BE displayed in this screen. pl refer  ""Enqu" & _
    "iry Status View"""
        '
        'DataGridPartNumbers
        '
        Me.DataGridPartNumbers.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridPartNumbers.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridPartNumbers.CaptionVisible = False
        Me.DataGridPartNumbers.DataMember = ""
        Me.DataGridPartNumbers.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridPartNumbers.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridPartNumbers.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridPartNumbers.Location = New System.Drawing.Point(705, 115)
        Me.DataGridPartNumbers.Name = "DataGridPartNumbers"
        Me.DataGridPartNumbers.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridPartNumbers.ParentRowsVisible = False
        Me.DataGridPartNumbers.PreferredColumnWidth = 85
        Me.DataGridPartNumbers.ReadOnly = True
        Me.DataGridPartNumbers.RowHeadersVisible = False
        Me.DataGridPartNumbers.Size = New System.Drawing.Size(585, 219)
        Me.DataGridPartNumbers.TabIndex = 205
        Me.DataGridPartNumbers.Visible = False
        '
        'DataGridCustomer1
        '
        Me.DataGridCustomer1.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridCustomer1.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCustomer1.CaptionVisible = False
        Me.DataGridCustomer1.DataMember = ""
        Me.DataGridCustomer1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCustomer1.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCustomer1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridCustomer1.Location = New System.Drawing.Point(727, 202)
        Me.DataGridCustomer1.Name = "DataGridCustomer1"
        Me.DataGridCustomer1.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridCustomer1.ParentRowsVisible = False
        Me.DataGridCustomer1.PreferredColumnWidth = 85
        Me.DataGridCustomer1.ReadOnly = True
        Me.DataGridCustomer1.RowHeadersVisible = False
        Me.DataGridCustomer1.Size = New System.Drawing.Size(548, 219)
        Me.DataGridCustomer1.TabIndex = 207
        Me.DataGridCustomer1.Visible = False
        '
        'RFQPriceViewQuick
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1507, 791)
        Me.Controls.Add(Me.DataGridCustomer1)
        Me.Controls.Add(Me.DataGridPartNumbers)
        Me.Controls.Add(Me.Note)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ButtonRefresh)
        Me.Controls.Add(Me.DataGridViewPriceView)
        Me.Name = "RFQPriceViewQuick"
        Me.Text = "RFQPriceViewQuick"
        CType(Me.DataGridViewPriceView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DataGridPartNumbers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridCustomer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButtonRefresh As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonQuotePending As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonPendingPrice As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonAll As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Protected WithEvents dtptodate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents dtpfrdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Note As System.Windows.Forms.Label
    Friend WithEvents DataGridViewPriceView As System.Windows.Forms.DataGridView
    Friend WithEvents RadioButtonSingle As System.Windows.Forms.RadioButton
    Friend WithEvents txtRegNo As System.Windows.Forms.TextBox
    Friend WithEvents RadioButtonPurNeed As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonId As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonName As System.Windows.Forms.RadioButton
    Friend WithEvents DataGridPartNumbers As System.Windows.Forms.DataGrid
    Friend WithEvents txtPartNumber As System.Windows.Forms.TextBox
    Friend WithEvents RadioButtonPartNumber As System.Windows.Forms.RadioButton
    Friend WithEvents txtCustID As System.Windows.Forms.TextBox
    Friend WithEvents lblCustName As System.Windows.Forms.Label
    Friend WithEvents DataGridCustomer1 As System.Windows.Forms.DataGrid
End Class
