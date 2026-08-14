<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ShipLabelDomestic
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
        Me.DataGridViewShipLabel = New System.Windows.Forms.DataGridView()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.LblInvoice = New System.Windows.Forms.Label()
        Me.txtInvoiceNo = New System.Windows.Forms.TextBox()
        Me.LblLineNo = New System.Windows.Forms.Label()
        Me.TxtLineNo = New System.Windows.Forms.TextBox()
        Me.btnclear = New System.Windows.Forms.Button()
        Me.BtnShipOK = New System.Windows.Forms.Button()
        Me.BtnPrint = New System.Windows.Forms.Button()
        Me.txtModeofDespatch = New System.Windows.Forms.TextBox()
        Me.lblModeofdespatch = New System.Windows.Forms.Label()
        Me.lblCustomer = New System.Windows.Forms.Label()
        Me.txtCustomer = New System.Windows.Forms.TextBox()
        CType(Me.DataGridViewShipLabel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridViewShipLabel
        '
        Me.DataGridViewShipLabel.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.DataGridViewShipLabel.AllowUserToAddRows = False
        Me.DataGridViewShipLabel.AllowUserToDeleteRows = False
        Me.DataGridViewShipLabel.AllowUserToResizeRows = False
        Me.DataGridViewShipLabel.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.DataGridViewShipLabel.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewShipLabel.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewShipLabel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewShipLabel.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewShipLabel.GridColor = System.Drawing.Color.Red
        Me.DataGridViewShipLabel.Location = New System.Drawing.Point(15, 119)
        Me.DataGridViewShipLabel.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridViewShipLabel.Name = "DataGridViewShipLabel"
        Me.DataGridViewShipLabel.RowHeadersWidth = 56
        Me.DataGridViewShipLabel.Size = New System.Drawing.Size(1436, 439)
        Me.DataGridViewShipLabel.TabIndex = 243
        '
        'ComboBox1
        '
        Me.ComboBox1.BackColor = System.Drawing.Color.White
        Me.ComboBox1.Font = New System.Drawing.Font("Arial Narrow", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.Items.AddRange(New Object() {"A)Item Labels Pre-Invoice", "B)Item Labels Post-Invoice", "C)Box Label Pre-Invoice", "D)Box Label Post-Invoice"})
        Me.ComboBox1.Location = New System.Drawing.Point(121, 16)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(366, 30)
        Me.ComboBox1.TabIndex = 249
        '
        'lblType
        '
        Me.lblType.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblType.Location = New System.Drawing.Point(12, 14)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(103, 19)
        Me.lblType.TabIndex = 244
        Me.lblType.Text = "Label Type"
        '
        'LblInvoice
        '
        Me.LblInvoice.Location = New System.Drawing.Point(493, 17)
        Me.LblInvoice.Name = "LblInvoice"
        Me.LblInvoice.Size = New System.Drawing.Size(143, 27)
        Me.LblInvoice.TabIndex = 245
        Me.LblInvoice.Text = "1"
        Me.LblInvoice.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtInvoiceNo
        '
        Me.txtInvoiceNo.BackColor = System.Drawing.Color.White
        Me.txtInvoiceNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtInvoiceNo.Location = New System.Drawing.Point(642, 14)
        Me.txtInvoiceNo.MaxLength = 20
        Me.txtInvoiceNo.Name = "txtInvoiceNo"
        Me.txtInvoiceNo.Size = New System.Drawing.Size(124, 22)
        Me.txtInvoiceNo.TabIndex = 246
        '
        'LblLineNo
        '
        Me.LblLineNo.Location = New System.Drawing.Point(788, 11)
        Me.LblLineNo.Name = "LblLineNo"
        Me.LblLineNo.Size = New System.Drawing.Size(86, 28)
        Me.LblLineNo.TabIndex = 248
        Me.LblLineNo.Text = "Line No."
        Me.LblLineNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TxtLineNo
        '
        Me.TxtLineNo.BackColor = System.Drawing.Color.White
        Me.TxtLineNo.ForeColor = System.Drawing.Color.Black
        Me.TxtLineNo.Location = New System.Drawing.Point(902, 14)
        Me.TxtLineNo.Name = "TxtLineNo"
        Me.TxtLineNo.Size = New System.Drawing.Size(103, 22)
        Me.TxtLineNo.TabIndex = 247
        Me.TxtLineNo.Text = "%"
        '
        'btnclear
        '
        Me.btnclear.BackColor = System.Drawing.Color.Silver
        Me.btnclear.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold)
        Me.btnclear.ForeColor = System.Drawing.Color.Black
        Me.btnclear.Location = New System.Drawing.Point(1347, 8)
        Me.btnclear.Name = "btnclear"
        Me.btnclear.Size = New System.Drawing.Size(104, 29)
        Me.btnclear.TabIndex = 251
        Me.btnclear.Text = "Clear"
        Me.btnclear.UseVisualStyleBackColor = False
        '
        'BtnShipOK
        '
        Me.BtnShipOK.BackColor = System.Drawing.Color.Silver
        Me.BtnShipOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold)
        Me.BtnShipOK.ForeColor = System.Drawing.Color.Black
        Me.BtnShipOK.Location = New System.Drawing.Point(1120, 9)
        Me.BtnShipOK.Name = "BtnShipOK"
        Me.BtnShipOK.Size = New System.Drawing.Size(97, 27)
        Me.BtnShipOK.TabIndex = 250
        Me.BtnShipOK.Text = "OK"
        Me.BtnShipOK.UseVisualStyleBackColor = False
        '
        'BtnPrint
        '
        Me.BtnPrint.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.BtnPrint.Enabled = False
        Me.BtnPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnPrint.Location = New System.Drawing.Point(1223, 7)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(118, 30)
        Me.BtnPrint.TabIndex = 252
        Me.BtnPrint.Text = "Print Label"
        Me.BtnPrint.UseVisualStyleBackColor = False
        '
        'txtModeofDespatch
        '
        Me.txtModeofDespatch.Location = New System.Drawing.Point(121, 53)
        Me.txtModeofDespatch.MaxLength = 60
        Me.txtModeofDespatch.Name = "txtModeofDespatch"
        Me.txtModeofDespatch.Size = New System.Drawing.Size(645, 22)
        Me.txtModeofDespatch.TabIndex = 253
        Me.txtModeofDespatch.Visible = False
        '
        'lblModeofdespatch
        '
        Me.lblModeofdespatch.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModeofdespatch.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblModeofdespatch.Location = New System.Drawing.Point(12, 53)
        Me.lblModeofdespatch.Name = "lblModeofdespatch"
        Me.lblModeofdespatch.Size = New System.Drawing.Size(103, 41)
        Me.lblModeofdespatch.TabIndex = 254
        Me.lblModeofdespatch.Text = "Mode of Despatch"
        Me.lblModeofdespatch.Visible = False
        '
        'lblCustomer
        '
        Me.lblCustomer.AutoSize = True
        Me.lblCustomer.Location = New System.Drawing.Point(788, 56)
        Me.lblCustomer.Name = "lblCustomer"
        Me.lblCustomer.Size = New System.Drawing.Size(109, 17)
        Me.lblCustomer.TabIndex = 255
        Me.lblCustomer.Text = "Customer Name"
        Me.lblCustomer.Visible = False
        '
        'txtCustomer
        '
        Me.txtCustomer.Location = New System.Drawing.Point(902, 53)
        Me.txtCustomer.Name = "txtCustomer"
        Me.txtCustomer.Size = New System.Drawing.Size(315, 22)
        Me.txtCustomer.TabIndex = 256
        Me.txtCustomer.Visible = False
        '
        'ShipLabelDomestic
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1511, 663)
        Me.Controls.Add(Me.txtCustomer)
        Me.Controls.Add(Me.lblCustomer)
        Me.Controls.Add(Me.lblModeofdespatch)
        Me.Controls.Add(Me.txtModeofDespatch)
        Me.Controls.Add(Me.BtnPrint)
        Me.Controls.Add(Me.btnclear)
        Me.Controls.Add(Me.BtnShipOK)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.LblInvoice)
        Me.Controls.Add(Me.txtInvoiceNo)
        Me.Controls.Add(Me.LblLineNo)
        Me.Controls.Add(Me.TxtLineNo)
        Me.Controls.Add(Me.DataGridViewShipLabel)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "ShipLabelDomestic"
        Me.Text = "Domestic Shipping Labels"
        CType(Me.DataGridViewShipLabel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridViewShipLabel As System.Windows.Forms.DataGridView
    Protected WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents LblInvoice As System.Windows.Forms.Label
    Friend WithEvents txtInvoiceNo As System.Windows.Forms.TextBox
    Friend WithEvents LblLineNo As System.Windows.Forms.Label
    Friend WithEvents TxtLineNo As System.Windows.Forms.TextBox
    Friend WithEvents btnclear As System.Windows.Forms.Button
    Friend WithEvents BtnShipOK As System.Windows.Forms.Button
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents txtModeofDespatch As System.Windows.Forms.TextBox
    Friend WithEvents lblModeofdespatch As System.Windows.Forms.Label
    Friend WithEvents lblCustomer As System.Windows.Forms.Label
    Friend WithEvents txtCustomer As System.Windows.Forms.TextBox
End Class
