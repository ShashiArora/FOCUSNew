<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RFQPriceViewALL
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
        Me.datagridRFQView = New System.Windows.Forms.DataGrid
        Me.txtenqdetailcode = New System.Windows.Forms.TextBox
        Me.txtregno = New System.Windows.Forms.TextBox
        Me.GroupBoxSelect = New System.Windows.Forms.GroupBox
        Me.ButtonRefresh = New System.Windows.Forms.Button
        Me.RadioButtonall = New System.Windows.Forms.RadioButton
        Me.RadioButtonunread = New System.Windows.Forms.RadioButton
        Me.RadioButtonread = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ButtonPartsRefresh = New System.Windows.Forms.Button
        Me.ButtonCustomerRefresh = New System.Windows.Forms.Button
        Me.GroupBoxPrice = New System.Windows.Forms.GroupBox
        Me.lblSpecial = New System.Windows.Forms.Label
        Me.CheckBoxRead = New System.Windows.Forms.CheckBox
        Me.txtspecial = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.DataGridQtyview = New System.Windows.Forms.DataGrid
        Me.DataGridCertificateChargesview = New System.Windows.Forms.DataGrid
        Me.ToolDetails = New System.Windows.Forms.GroupBox
        Me.lblProd = New System.Windows.Forms.Label
        Me.lblProto = New System.Windows.Forms.Label
        Me.LabelToolClose = New System.Windows.Forms.Label
        Me.ProdLeadTime = New System.Windows.Forms.TextBox
        Me.ProtoLeadTime = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.ProdCustShare = New System.Windows.Forms.TextBox
        Me.ProtoCustShare = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.ProdTotalCost = New System.Windows.Forms.TextBox
        Me.ProtoTotal = New System.Windows.Forms.TextBox
        Me.lblTotalCost = New System.Windows.Forms.Label
        Me.GroupBoxPart = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtpartbyApl = New System.Windows.Forms.TextBox
        Me.lblpart = New System.Windows.Forms.Label
        Me.txtAplSpecial = New System.Windows.Forms.TextBox
        CType(Me.datagridRFQView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxSelect.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBoxPrice.SuspendLayout()
        CType(Me.DataGridQtyview, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridCertificateChargesview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolDetails.SuspendLayout()
        Me.GroupBoxPart.SuspendLayout()
        Me.SuspendLayout()
        '
        'datagridRFQView
        '
        Me.datagridRFQView.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.datagridRFQView.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridRFQView.CaptionVisible = False
        Me.datagridRFQView.DataMember = ""
        Me.datagridRFQView.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridRFQView.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridRFQView.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.datagridRFQView.Location = New System.Drawing.Point(12, 58)
        Me.datagridRFQView.Name = "datagridRFQView"
        Me.datagridRFQView.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.datagridRFQView.ParentRowsVisible = False
        Me.datagridRFQView.PreferredColumnWidth = 85
        Me.datagridRFQView.ReadOnly = True
        Me.datagridRFQView.RowHeadersVisible = False
        Me.datagridRFQView.Size = New System.Drawing.Size(1236, 402)
        Me.datagridRFQView.TabIndex = 18
        '
        'txtenqdetailcode
        '
        Me.txtenqdetailcode.Location = New System.Drawing.Point(-7, 482)
        Me.txtenqdetailcode.Name = "txtenqdetailcode"
        Me.txtenqdetailcode.Size = New System.Drawing.Size(20, 20)
        Me.txtenqdetailcode.TabIndex = 224
        Me.txtenqdetailcode.Visible = False
        '
        'txtregno
        '
        Me.txtregno.Location = New System.Drawing.Point(-7, 456)
        Me.txtregno.Name = "txtregno"
        Me.txtregno.Size = New System.Drawing.Size(20, 20)
        Me.txtregno.TabIndex = 225
        Me.txtregno.Visible = False
        '
        'GroupBoxSelect
        '
        Me.GroupBoxSelect.Controls.Add(Me.ButtonRefresh)
        Me.GroupBoxSelect.Controls.Add(Me.RadioButtonall)
        Me.GroupBoxSelect.Controls.Add(Me.RadioButtonunread)
        Me.GroupBoxSelect.Controls.Add(Me.RadioButtonread)
        Me.GroupBoxSelect.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBoxSelect.ForeColor = System.Drawing.Color.Black
        Me.GroupBoxSelect.Location = New System.Drawing.Point(21, 3)
        Me.GroupBoxSelect.Name = "GroupBoxSelect"
        Me.GroupBoxSelect.Size = New System.Drawing.Size(339, 37)
        Me.GroupBoxSelect.TabIndex = 226
        Me.GroupBoxSelect.TabStop = False
        Me.GroupBoxSelect.Text = "Price data selection"
        '
        'ButtonRefresh
        '
        Me.ButtonRefresh.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ButtonRefresh.ForeColor = System.Drawing.Color.White
        Me.ButtonRefresh.Location = New System.Drawing.Point(245, 8)
        Me.ButtonRefresh.Name = "ButtonRefresh"
        Me.ButtonRefresh.Size = New System.Drawing.Size(75, 23)
        Me.ButtonRefresh.TabIndex = 116
        Me.ButtonRefresh.Text = "Refresh"
        Me.ButtonRefresh.UseVisualStyleBackColor = False
        '
        'RadioButtonall
        '
        Me.RadioButtonall.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonall.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonall.Location = New System.Drawing.Point(192, 10)
        Me.RadioButtonall.Name = "RadioButtonall"
        Me.RadioButtonall.Size = New System.Drawing.Size(64, 18)
        Me.RadioButtonall.TabIndex = 115
        Me.RadioButtonall.Text = "All"
        '
        'RadioButtonunread
        '
        Me.RadioButtonunread.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonunread.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonunread.Location = New System.Drawing.Point(34, 12)
        Me.RadioButtonunread.Name = "RadioButtonunread"
        Me.RadioButtonunread.Size = New System.Drawing.Size(72, 18)
        Me.RadioButtonunread.TabIndex = 113
        Me.RadioButtonunread.Text = "Unread"
        '
        'RadioButtonread
        '
        Me.RadioButtonread.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonread.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonread.Location = New System.Drawing.Point(110, 10)
        Me.RadioButtonread.Name = "RadioButtonread"
        Me.RadioButtonread.Size = New System.Drawing.Size(80, 18)
        Me.RadioButtonread.TabIndex = 114
        Me.RadioButtonread.Text = "Read"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ButtonPartsRefresh)
        Me.GroupBox1.Controls.Add(Me.ButtonCustomerRefresh)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(461, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(339, 37)
        Me.GroupBox1.TabIndex = 228
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Customer  And Parts"
        '
        'ButtonPartsRefresh
        '
        Me.ButtonPartsRefresh.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ButtonPartsRefresh.ForeColor = System.Drawing.Color.White
        Me.ButtonPartsRefresh.Location = New System.Drawing.Point(223, 11)
        Me.ButtonPartsRefresh.Name = "ButtonPartsRefresh"
        Me.ButtonPartsRefresh.Size = New System.Drawing.Size(110, 23)
        Me.ButtonPartsRefresh.TabIndex = 117
        Me.ButtonPartsRefresh.Text = "Parts Refresh"
        Me.ButtonPartsRefresh.UseVisualStyleBackColor = False
        '
        'ButtonCustomerRefresh
        '
        Me.ButtonCustomerRefresh.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ButtonCustomerRefresh.ForeColor = System.Drawing.Color.White
        Me.ButtonCustomerRefresh.Location = New System.Drawing.Point(109, 11)
        Me.ButtonCustomerRefresh.Name = "ButtonCustomerRefresh"
        Me.ButtonCustomerRefresh.Size = New System.Drawing.Size(110, 23)
        Me.ButtonCustomerRefresh.TabIndex = 116
        Me.ButtonCustomerRefresh.Text = "Customer Refresh"
        Me.ButtonCustomerRefresh.UseVisualStyleBackColor = False
        '
        'GroupBoxPrice
        '
        Me.GroupBoxPrice.Controls.Add(Me.lblSpecial)
        Me.GroupBoxPrice.Controls.Add(Me.CheckBoxRead)
        Me.GroupBoxPrice.Controls.Add(Me.txtspecial)
        Me.GroupBoxPrice.Controls.Add(Me.TextBox1)
        Me.GroupBoxPrice.Controls.Add(Me.DataGridQtyview)
        Me.GroupBoxPrice.Controls.Add(Me.DataGridCertificateChargesview)
        Me.GroupBoxPrice.Controls.Add(Me.ToolDetails)
        Me.GroupBoxPrice.Enabled = False
        Me.GroupBoxPrice.Location = New System.Drawing.Point(12, 482)
        Me.GroupBoxPrice.Name = "GroupBoxPrice"
        Me.GroupBoxPrice.Size = New System.Drawing.Size(1237, 192)
        Me.GroupBoxPrice.TabIndex = 230
        Me.GroupBoxPrice.TabStop = False
        Me.GroupBoxPrice.Text = "Price Details"
        '
        'lblSpecial
        '
        Me.lblSpecial.AutoSize = True
        Me.lblSpecial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSpecial.ForeColor = System.Drawing.Color.Black
        Me.lblSpecial.Location = New System.Drawing.Point(6, 20)
        Me.lblSpecial.Name = "lblSpecial"
        Me.lblSpecial.Size = New System.Drawing.Size(109, 14)
        Me.lblSpecial.TabIndex = 230
        Me.lblSpecial.Text = "Special Instruction"
        '
        'CheckBoxRead
        '
        Me.CheckBoxRead.AutoSize = True
        Me.CheckBoxRead.Location = New System.Drawing.Point(1128, 19)
        Me.CheckBoxRead.Name = "CheckBoxRead"
        Me.CheckBoxRead.Size = New System.Drawing.Size(94, 17)
        Me.CheckBoxRead.TabIndex = 229
        Me.CheckBoxRead.Text = "Mark As Read"
        Me.CheckBoxRead.UseVisualStyleBackColor = True
        '
        'txtspecial
        '
        Me.txtspecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtspecial.Location = New System.Drawing.Point(115, 11)
        Me.txtspecial.Multiline = True
        Me.txtspecial.Name = "txtspecial"
        Me.txtspecial.Size = New System.Drawing.Size(1011, 35)
        Me.txtspecial.TabIndex = 228
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.TextBox1.Location = New System.Drawing.Point(9, 159)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(353, 20)
        Me.TextBox1.TabIndex = 227
        Me.TextBox1.Text = "Note : Lead Time = Product Lead Time + Tool Lead Time"
        '
        'DataGridQtyview
        '
        Me.DataGridQtyview.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridQtyview.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQtyview.CaptionForeColor = System.Drawing.Color.Black
        Me.DataGridQtyview.CaptionVisible = False
        Me.DataGridQtyview.DataMember = ""
        Me.DataGridQtyview.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQtyview.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQtyview.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridQtyview.Location = New System.Drawing.Point(833, 51)
        Me.DataGridQtyview.Name = "DataGridQtyview"
        Me.DataGridQtyview.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridQtyview.ParentRowsVisible = False
        Me.DataGridQtyview.RowHeadersVisible = False
        Me.DataGridQtyview.RowHeaderWidth = 20
        Me.DataGridQtyview.Size = New System.Drawing.Size(389, 132)
        Me.DataGridQtyview.TabIndex = 226
        '
        'DataGridCertificateChargesview
        '
        Me.DataGridCertificateChargesview.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridCertificateChargesview.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCertificateChargesview.CaptionForeColor = System.Drawing.Color.Black
        Me.DataGridCertificateChargesview.CaptionVisible = False
        Me.DataGridCertificateChargesview.DataMember = ""
        Me.DataGridCertificateChargesview.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCertificateChargesview.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCertificateChargesview.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridCertificateChargesview.Location = New System.Drawing.Point(434, 51)
        Me.DataGridCertificateChargesview.Name = "DataGridCertificateChargesview"
        Me.DataGridCertificateChargesview.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridCertificateChargesview.ParentRowsVisible = False
        Me.DataGridCertificateChargesview.PreferredColumnWidth = 100
        Me.DataGridCertificateChargesview.RowHeadersVisible = False
        Me.DataGridCertificateChargesview.RowHeaderWidth = 20
        Me.DataGridCertificateChargesview.Size = New System.Drawing.Size(393, 132)
        Me.DataGridCertificateChargesview.TabIndex = 225
        '
        'ToolDetails
        '
        Me.ToolDetails.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ToolDetails.Controls.Add(Me.lblProd)
        Me.ToolDetails.Controls.Add(Me.lblProto)
        Me.ToolDetails.Controls.Add(Me.LabelToolClose)
        Me.ToolDetails.Controls.Add(Me.ProdLeadTime)
        Me.ToolDetails.Controls.Add(Me.ProtoLeadTime)
        Me.ToolDetails.Controls.Add(Me.Label20)
        Me.ToolDetails.Controls.Add(Me.ProdCustShare)
        Me.ToolDetails.Controls.Add(Me.ProtoCustShare)
        Me.ToolDetails.Controls.Add(Me.Label13)
        Me.ToolDetails.Controls.Add(Me.ProdTotalCost)
        Me.ToolDetails.Controls.Add(Me.ProtoTotal)
        Me.ToolDetails.Controls.Add(Me.lblTotalCost)
        Me.ToolDetails.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ToolDetails.Location = New System.Drawing.Point(9, 51)
        Me.ToolDetails.Name = "ToolDetails"
        Me.ToolDetails.Size = New System.Drawing.Size(414, 94)
        Me.ToolDetails.TabIndex = 224
        Me.ToolDetails.TabStop = False
        Me.ToolDetails.Text = "Tool Details"
        '
        'lblProd
        '
        Me.lblProd.AutoSize = True
        Me.lblProd.Font = New System.Drawing.Font("Arial", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProd.ForeColor = System.Drawing.Color.Blue
        Me.lblProd.Location = New System.Drawing.Point(11, 62)
        Me.lblProd.Name = "lblProd"
        Me.lblProd.Size = New System.Drawing.Size(94, 15)
        Me.lblProd.TabIndex = 30
        Me.lblProd.Text = "Prod. Tool Cost"
        '
        'lblProto
        '
        Me.lblProto.AutoSize = True
        Me.lblProto.Font = New System.Drawing.Font("Arial", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProto.ForeColor = System.Drawing.Color.Blue
        Me.lblProto.Location = New System.Drawing.Point(11, 27)
        Me.lblProto.Name = "lblProto"
        Me.lblProto.Size = New System.Drawing.Size(95, 15)
        Me.lblProto.TabIndex = 29
        Me.lblProto.Text = "Proto Tool Cost"
        '
        'LabelToolClose
        '
        Me.LabelToolClose.AutoSize = True
        Me.LabelToolClose.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelToolClose.Location = New System.Drawing.Point(612, 11)
        Me.LabelToolClose.Name = "LabelToolClose"
        Me.LabelToolClose.Size = New System.Drawing.Size(21, 22)
        Me.LabelToolClose.TabIndex = 17
        Me.LabelToolClose.Text = "+"
        '
        'ProdLeadTime
        '
        Me.ProdLeadTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdLeadTime.Location = New System.Drawing.Point(294, 61)
        Me.ProdLeadTime.Name = "ProdLeadTime"
        Me.ProdLeadTime.Size = New System.Drawing.Size(83, 20)
        Me.ProdLeadTime.TabIndex = 26
        '
        'ProtoLeadTime
        '
        Me.ProtoLeadTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProtoLeadTime.Location = New System.Drawing.Point(294, 28)
        Me.ProtoLeadTime.Name = "ProtoLeadTime"
        Me.ProtoLeadTime.Size = New System.Drawing.Size(83, 20)
        Me.ProtoLeadTime.TabIndex = 21
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Blue
        Me.Label20.Location = New System.Drawing.Point(291, 11)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(102, 14)
        Me.Label20.TabIndex = 8
        Me.Label20.Text = "Lead Time (Days)"
        '
        'ProdCustShare
        '
        Me.ProdCustShare.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdCustShare.Location = New System.Drawing.Point(205, 61)
        Me.ProdCustShare.Name = "ProdCustShare"
        Me.ProdCustShare.Size = New System.Drawing.Size(83, 20)
        Me.ProdCustShare.TabIndex = 25
        '
        'ProtoCustShare
        '
        Me.ProtoCustShare.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProtoCustShare.Location = New System.Drawing.Point(205, 28)
        Me.ProtoCustShare.Name = "ProtoCustShare"
        Me.ProtoCustShare.Size = New System.Drawing.Size(83, 20)
        Me.ProtoCustShare.TabIndex = 20
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Blue
        Me.Label13.Location = New System.Drawing.Point(205, 11)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(87, 14)
        Me.Label13.TabIndex = 5
        Me.Label13.Text = "Cust Cost(INR)"
        '
        'ProdTotalCost
        '
        Me.ProdTotalCost.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdTotalCost.Location = New System.Drawing.Point(124, 61)
        Me.ProdTotalCost.Name = "ProdTotalCost"
        Me.ProdTotalCost.Size = New System.Drawing.Size(75, 20)
        Me.ProdTotalCost.TabIndex = 24
        '
        'ProtoTotal
        '
        Me.ProtoTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProtoTotal.Location = New System.Drawing.Point(124, 28)
        Me.ProtoTotal.Name = "ProtoTotal"
        Me.ProtoTotal.Size = New System.Drawing.Size(75, 20)
        Me.ProtoTotal.TabIndex = 19
        '
        'lblTotalCost
        '
        Me.lblTotalCost.AutoSize = True
        Me.lblTotalCost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalCost.ForeColor = System.Drawing.Color.Blue
        Me.lblTotalCost.Location = New System.Drawing.Point(109, 11)
        Me.lblTotalCost.Name = "lblTotalCost"
        Me.lblTotalCost.Size = New System.Drawing.Size(90, 14)
        Me.lblTotalCost.TabIndex = 2
        Me.lblTotalCost.Text = "Total Cost (INR)"
        '
        'GroupBoxPart
        '
        Me.GroupBoxPart.Controls.Add(Me.Label1)
        Me.GroupBoxPart.Controls.Add(Me.txtpartbyApl)
        Me.GroupBoxPart.Controls.Add(Me.lblpart)
        Me.GroupBoxPart.Controls.Add(Me.txtAplSpecial)
        Me.GroupBoxPart.Location = New System.Drawing.Point(12, 466)
        Me.GroupBoxPart.Name = "GroupBoxPart"
        Me.GroupBoxPart.Size = New System.Drawing.Size(1122, 100)
        Me.GroupBoxPart.TabIndex = 234
        Me.GroupBoxPart.TabStop = False
        Me.GroupBoxPart.Text = "Part Details"
        Me.GroupBoxPart.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(321, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(194, 14)
        Me.Label1.TabIndex = 237
        Me.Label1.Text = "Special Instruction from Apl. Dept."
        Me.Label1.Visible = False
        '
        'txtpartbyApl
        '
        Me.txtpartbyApl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpartbyApl.Location = New System.Drawing.Point(18, 41)
        Me.txtpartbyApl.Name = "txtpartbyApl"
        Me.txtpartbyApl.Size = New System.Drawing.Size(279, 20)
        Me.txtpartbyApl.TabIndex = 236
        Me.txtpartbyApl.Visible = False
        '
        'lblpart
        '
        Me.lblpart.AutoSize = True
        Me.lblpart.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblpart.ForeColor = System.Drawing.Color.Black
        Me.lblpart.Location = New System.Drawing.Point(20, 19)
        Me.lblpart.Name = "lblpart"
        Me.lblpart.Size = New System.Drawing.Size(144, 14)
        Me.lblpart.TabIndex = 235
        Me.lblpart.Text = "Part Number by Apl.Dept:"
        Me.lblpart.Visible = False
        '
        'txtAplSpecial
        '
        Me.txtAplSpecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAplSpecial.Location = New System.Drawing.Point(324, 39)
        Me.txtAplSpecial.Multiline = True
        Me.txtAplSpecial.Name = "txtAplSpecial"
        Me.txtAplSpecial.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtAplSpecial.Size = New System.Drawing.Size(742, 55)
        Me.txtAplSpecial.TabIndex = 234
        Me.txtAplSpecial.Visible = False
        '
        'RFQPriceViewALL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1274, 678)
        Me.Controls.Add(Me.GroupBoxPart)
        Me.Controls.Add(Me.datagridRFQView)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBoxSelect)
        Me.Controls.Add(Me.txtregno)
        Me.Controls.Add(Me.txtenqdetailcode)
        Me.Controls.Add(Me.GroupBoxPrice)
        Me.Name = "RFQPriceViewALL"
        Me.Text = "RFQPriceView"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.datagridRFQView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxSelect.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBoxPrice.ResumeLayout(False)
        Me.GroupBoxPrice.PerformLayout()
        CType(Me.DataGridQtyview, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridCertificateChargesview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolDetails.ResumeLayout(False)
        Me.ToolDetails.PerformLayout()
        Me.GroupBoxPart.ResumeLayout(False)
        Me.GroupBoxPart.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents datagridRFQView As System.Windows.Forms.DataGrid
    Friend WithEvents txtenqdetailcode As System.Windows.Forms.TextBox
    Friend WithEvents txtregno As System.Windows.Forms.TextBox
    Friend WithEvents GroupBoxSelect As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonRefresh As System.Windows.Forms.Button
    Friend WithEvents RadioButtonall As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonunread As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonread As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonCustomerRefresh As System.Windows.Forms.Button
    Friend WithEvents ButtonPartsRefresh As System.Windows.Forms.Button
    Friend WithEvents GroupBoxPrice As System.Windows.Forms.GroupBox
    Protected WithEvents DataGridCertificateChargesview As System.Windows.Forms.DataGrid
    Friend WithEvents ToolDetails As System.Windows.Forms.GroupBox
    Friend WithEvents lblProd As System.Windows.Forms.Label
    Friend WithEvents lblProto As System.Windows.Forms.Label
    Friend WithEvents LabelToolClose As System.Windows.Forms.Label
    Friend WithEvents ProdLeadTime As System.Windows.Forms.TextBox
    Friend WithEvents ProtoLeadTime As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents ProdCustShare As System.Windows.Forms.TextBox
    Friend WithEvents ProtoCustShare As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ProdTotalCost As System.Windows.Forms.TextBox
    Friend WithEvents ProtoTotal As System.Windows.Forms.TextBox
    Friend WithEvents lblTotalCost As System.Windows.Forms.Label
    Friend WithEvents lblSpecial As System.Windows.Forms.Label
    Friend WithEvents CheckBoxRead As System.Windows.Forms.CheckBox
    Friend WithEvents txtspecial As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Protected WithEvents DataGridQtyview As System.Windows.Forms.DataGrid
    Friend WithEvents GroupBoxPart As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtpartbyApl As System.Windows.Forms.TextBox
    Friend WithEvents lblpart As System.Windows.Forms.Label
    Friend WithEvents txtAplSpecial As System.Windows.Forms.TextBox
End Class
