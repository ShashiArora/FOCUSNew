<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EnquiryStatus
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
        Me.DataGridViewEnquiryStatus = New System.Windows.Forms.DataGridView
        Me.ComboBoxCSR = New System.Windows.Forms.ComboBox
        Me.ComboBoxISR = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtcustid = New System.Windows.Forms.TextBox
        Me.txtcustname = New System.Windows.Forms.TextBox
        Me.RadioButtonSummary = New System.Windows.Forms.RadioButton
        Me.RadioButtonDetail = New System.Windows.Forms.RadioButton
        Me.ButtonRefresh = New System.Windows.Forms.Button
        Me.DataGridCustomer = New System.Windows.Forms.DataGrid
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        CType(Me.DataGridViewEnquiryStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridViewEnquiryStatus
        '
        Me.DataGridViewEnquiryStatus.AllowUserToAddRows = False
        Me.DataGridViewEnquiryStatus.AllowUserToDeleteRows = False
        Me.DataGridViewEnquiryStatus.AllowUserToResizeRows = False
        Me.DataGridViewEnquiryStatus.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DataGridViewEnquiryStatus.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.DataGridViewEnquiryStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewEnquiryStatus.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridViewEnquiryStatus.GridColor = System.Drawing.Color.Red
        Me.DataGridViewEnquiryStatus.Location = New System.Drawing.Point(12, 68)
        Me.DataGridViewEnquiryStatus.Name = "DataGridViewEnquiryStatus"
        Me.DataGridViewEnquiryStatus.ReadOnly = True
        Me.DataGridViewEnquiryStatus.Size = New System.Drawing.Size(1159, 608)
        Me.DataGridViewEnquiryStatus.TabIndex = 0
        '
        'ComboBoxCSR
        '
        Me.ComboBoxCSR.FormattingEnabled = True
        Me.ComboBoxCSR.Location = New System.Drawing.Point(47, 12)
        Me.ComboBoxCSR.Name = "ComboBoxCSR"
        Me.ComboBoxCSR.Size = New System.Drawing.Size(71, 21)
        Me.ComboBoxCSR.TabIndex = 1
        '
        'ComboBoxISR
        '
        Me.ComboBoxISR.FormattingEnabled = True
        Me.ComboBoxISR.Location = New System.Drawing.Point(47, 39)
        Me.ComboBoxISR.Name = "ComboBoxISR"
        Me.ComboBoxISR.Size = New System.Drawing.Size(71, 21)
        Me.ComboBoxISR.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "CSR"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "ISR"
        '
        'dtpFromDate
        '
        Me.dtpFromDate.Location = New System.Drawing.Point(197, 13)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(154, 20)
        Me.dtpFromDate.TabIndex = 5
        '
        'dtpToDate
        '
        Me.dtpToDate.Location = New System.Drawing.Point(197, 39)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(154, 20)
        Me.dtpToDate.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(124, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "From Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(124, 37)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "To Date"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(362, 15)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Customer ID"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(362, 43)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(82, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Customer Name"
        '
        'txtcustid
        '
        Me.txtcustid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcustid.Location = New System.Drawing.Point(450, 12)
        Me.txtcustid.Name = "txtcustid"
        Me.txtcustid.Size = New System.Drawing.Size(100, 20)
        Me.txtcustid.TabIndex = 11
        '
        'txtcustname
        '
        Me.txtcustname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcustname.Location = New System.Drawing.Point(450, 42)
        Me.txtcustname.Name = "txtcustname"
        Me.txtcustname.Size = New System.Drawing.Size(262, 20)
        Me.txtcustname.TabIndex = 12
        '
        'RadioButtonSummary
        '
        Me.RadioButtonSummary.AutoSize = True
        Me.RadioButtonSummary.Location = New System.Drawing.Point(597, 10)
        Me.RadioButtonSummary.Name = "RadioButtonSummary"
        Me.RadioButtonSummary.Size = New System.Drawing.Size(68, 17)
        Me.RadioButtonSummary.TabIndex = 13
        Me.RadioButtonSummary.TabStop = True
        Me.RadioButtonSummary.Text = "Summary"
        Me.RadioButtonSummary.UseVisualStyleBackColor = True
        '
        'RadioButtonDetail
        '
        Me.RadioButtonDetail.AutoSize = True
        Me.RadioButtonDetail.Location = New System.Drawing.Point(402, 24)
        Me.RadioButtonDetail.Name = "RadioButtonDetail"
        Me.RadioButtonDetail.Size = New System.Drawing.Size(52, 17)
        Me.RadioButtonDetail.TabIndex = 14
        Me.RadioButtonDetail.TabStop = True
        Me.RadioButtonDetail.Text = "Detail"
        Me.RadioButtonDetail.UseVisualStyleBackColor = True
        Me.RadioButtonDetail.Visible = False
        '
        'ButtonRefresh
        '
        Me.ButtonRefresh.Location = New System.Drawing.Point(671, 5)
        Me.ButtonRefresh.Name = "ButtonRefresh"
        Me.ButtonRefresh.Size = New System.Drawing.Size(75, 23)
        Me.ButtonRefresh.TabIndex = 15
        Me.ButtonRefresh.Text = "Refresh"
        Me.ButtonRefresh.UseVisualStyleBackColor = True
        '
        'DataGridCustomer
        '
        Me.DataGridCustomer.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridCustomer.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCustomer.CaptionVisible = False
        Me.DataGridCustomer.DataMember = ""
        Me.DataGridCustomer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCustomer.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCustomer.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridCustomer.Location = New System.Drawing.Point(556, 9)
        Me.DataGridCustomer.Name = "DataGridCustomer"
        Me.DataGridCustomer.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridCustomer.ParentRowsVisible = False
        Me.DataGridCustomer.PreferredColumnWidth = 85
        Me.DataGridCustomer.ReadOnly = True
        Me.DataGridCustomer.RowHeadersVisible = False
        Me.DataGridCustomer.Size = New System.Drawing.Size(35, 24)
        Me.DataGridCustomer.TabIndex = 16
        Me.DataGridCustomer.Visible = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Magenta
        Me.Button1.Location = New System.Drawing.Point(1060, 15)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(111, 26)
        Me.Button1.TabIndex = 17
        Me.Button1.Text = "Pending"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Button2.Location = New System.Drawing.Point(1060, 39)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(111, 27)
        Me.Button2.TabIndex = 18
        Me.Button2.Text = "Rejected/Closed"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'EnquiryStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1233, 688)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridCustomer)
        Me.Controls.Add(Me.ButtonRefresh)
        Me.Controls.Add(Me.RadioButtonDetail)
        Me.Controls.Add(Me.RadioButtonSummary)
        Me.Controls.Add(Me.txtcustname)
        Me.Controls.Add(Me.txtcustid)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtpToDate)
        Me.Controls.Add(Me.dtpFromDate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxISR)
        Me.Controls.Add(Me.ComboBoxCSR)
        Me.Controls.Add(Me.DataGridViewEnquiryStatus)
        Me.Name = "EnquiryStatus"
        Me.Text = "Enquiry Status"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DataGridViewEnquiryStatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridViewEnquiryStatus As System.Windows.Forms.DataGridView
    Friend WithEvents ComboBoxCSR As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxISR As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtcustid As System.Windows.Forms.TextBox
    Friend WithEvents txtcustname As System.Windows.Forms.TextBox
    Friend WithEvents RadioButtonSummary As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonDetail As System.Windows.Forms.RadioButton
    Friend WithEvents ButtonRefresh As System.Windows.Forms.Button
    Friend WithEvents DataGridCustomer As System.Windows.Forms.DataGrid
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
