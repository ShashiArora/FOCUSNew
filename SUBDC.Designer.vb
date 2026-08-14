<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SUBDC
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
        Me.DataGridViewSCDC = New System.Windows.Forms.DataGridView()
        Me.GroupBoxPODetails = New System.Windows.Forms.GroupBox()
        Me.BTNDCSAVE = New System.Windows.Forms.Button()
        Me.txtRemark2 = New System.Windows.Forms.TextBox()
        Me.txtRemarks1 = New System.Windows.Forms.TextBox()
        Me.LblRemarks2 = New System.Windows.Forms.Label()
        Me.LblRemarks1 = New System.Windows.Forms.Label()
        Me.RBNewDC = New System.Windows.Forms.RadioButton()
        Me.RBDelDC = New System.Windows.Forms.RadioButton()
        Me.RBCancel = New System.Windows.Forms.RadioButton()
        Me.RBView = New System.Windows.Forms.RadioButton()
        Me.lblSelectVendor = New System.Windows.Forms.Label()
        Me.ComboBoxVendors = New System.Windows.Forms.ComboBox()
        Me.BtnOK = New System.Windows.Forms.Button()
        Me.txtdc = New System.Windows.Forms.TextBox()
        Me.txtHDKEY = New System.Windows.Forms.TextBox()
        Me.txtPONumber = New System.Windows.Forms.TextBox()
        Me.lblDC = New System.Windows.Forms.Label()
        Me.lblDCDt = New System.Windows.Forms.Label()
        Me.GBSC = New System.Windows.Forms.GroupBox()
        Me.lblLineSel = New System.Windows.Forms.Label()
        Me.txtchkcount = New System.Windows.Forms.TextBox()
        Me.DateTimeDCDT = New System.Windows.Forms.DateTimePicker()
        CType(Me.DataGridViewSCDC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxPODetails.SuspendLayout()
        Me.GBSC.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridViewSCDC
        '
        Me.DataGridViewSCDC.AllowUserToAddRows = False
        Me.DataGridViewSCDC.AllowUserToDeleteRows = False
        Me.DataGridViewSCDC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewSCDC.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewSCDC.Location = New System.Drawing.Point(15, 19)
        Me.DataGridViewSCDC.Name = "DataGridViewSCDC"
        Me.DataGridViewSCDC.Size = New System.Drawing.Size(954, 433)
        Me.DataGridViewSCDC.TabIndex = 1
        '
        'GroupBoxPODetails
        '
        Me.GroupBoxPODetails.Controls.Add(Me.BTNDCSAVE)
        Me.GroupBoxPODetails.Controls.Add(Me.txtRemark2)
        Me.GroupBoxPODetails.Controls.Add(Me.txtRemarks1)
        Me.GroupBoxPODetails.Controls.Add(Me.LblRemarks2)
        Me.GroupBoxPODetails.Controls.Add(Me.LblRemarks1)
        Me.GroupBoxPODetails.Controls.Add(Me.DataGridViewSCDC)
        Me.GroupBoxPODetails.Location = New System.Drawing.Point(28, 109)
        Me.GroupBoxPODetails.Name = "GroupBoxPODetails"
        Me.GroupBoxPODetails.Size = New System.Drawing.Size(989, 600)
        Me.GroupBoxPODetails.TabIndex = 2
        Me.GroupBoxPODetails.TabStop = False
        Me.GroupBoxPODetails.Text = "Purchase Order Details"
        '
        'BTNDCSAVE
        '
        Me.BTNDCSAVE.Location = New System.Drawing.Point(861, 565)
        Me.BTNDCSAVE.Name = "BTNDCSAVE"
        Me.BTNDCSAVE.Size = New System.Drawing.Size(108, 23)
        Me.BTNDCSAVE.TabIndex = 6
        Me.BTNDCSAVE.Text = "Generate Dc"
        Me.BTNDCSAVE.UseVisualStyleBackColor = True
        '
        'txtRemark2
        '
        Me.txtRemark2.Location = New System.Drawing.Point(73, 516)
        Me.txtRemark2.MaxLength = 250
        Me.txtRemark2.Multiline = True
        Me.txtRemark2.Name = "txtRemark2"
        Me.txtRemark2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRemark2.Size = New System.Drawing.Size(896, 43)
        Me.txtRemark2.TabIndex = 5
        '
        'txtRemarks1
        '
        Me.txtRemarks1.Location = New System.Drawing.Point(73, 467)
        Me.txtRemarks1.MaxLength = 250
        Me.txtRemarks1.Multiline = True
        Me.txtRemarks1.Name = "txtRemarks1"
        Me.txtRemarks1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRemarks1.Size = New System.Drawing.Size(896, 43)
        Me.txtRemarks1.TabIndex = 4
        '
        'LblRemarks2
        '
        Me.LblRemarks2.AutoSize = True
        Me.LblRemarks2.Location = New System.Drawing.Point(12, 513)
        Me.LblRemarks2.Name = "LblRemarks2"
        Me.LblRemarks2.Size = New System.Drawing.Size(55, 13)
        Me.LblRemarks2.TabIndex = 3
        Me.LblRemarks2.Text = "Remarks2"
        '
        'LblRemarks1
        '
        Me.LblRemarks1.AutoSize = True
        Me.LblRemarks1.Location = New System.Drawing.Point(12, 467)
        Me.LblRemarks1.Name = "LblRemarks1"
        Me.LblRemarks1.Size = New System.Drawing.Size(55, 13)
        Me.LblRemarks1.TabIndex = 2
        Me.LblRemarks1.Text = "Remarks1"
        '
        'RBNewDC
        '
        Me.RBNewDC.AutoSize = True
        Me.RBNewDC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBNewDC.Location = New System.Drawing.Point(52, 31)
        Me.RBNewDC.Name = "RBNewDC"
        Me.RBNewDC.Size = New System.Drawing.Size(50, 17)
        Me.RBNewDC.TabIndex = 0
        Me.RBNewDC.TabStop = True
        Me.RBNewDC.Text = "New"
        Me.RBNewDC.UseVisualStyleBackColor = True
        '
        'RBDelDC
        '
        Me.RBDelDC.AutoSize = True
        Me.RBDelDC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBDelDC.Location = New System.Drawing.Point(196, 31)
        Me.RBDelDC.Name = "RBDelDC"
        Me.RBDelDC.Size = New System.Drawing.Size(66, 17)
        Me.RBDelDC.TabIndex = 1
        Me.RBDelDC.TabStop = True
        Me.RBDelDC.Text = "Delete "
        Me.RBDelDC.UseVisualStyleBackColor = True
        '
        'RBCancel
        '
        Me.RBCancel.AutoSize = True
        Me.RBCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBCancel.Location = New System.Drawing.Point(374, 31)
        Me.RBCancel.Name = "RBCancel"
        Me.RBCancel.Size = New System.Drawing.Size(64, 17)
        Me.RBCancel.TabIndex = 2
        Me.RBCancel.TabStop = True
        Me.RBCancel.Text = "Cancel"
        Me.RBCancel.UseVisualStyleBackColor = True
        '
        'RBView
        '
        Me.RBView.AutoSize = True
        Me.RBView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBView.Location = New System.Drawing.Point(581, 31)
        Me.RBView.Name = "RBView"
        Me.RBView.Size = New System.Drawing.Size(52, 17)
        Me.RBView.TabIndex = 3
        Me.RBView.TabStop = True
        Me.RBView.Text = "View"
        Me.RBView.UseVisualStyleBackColor = True
        '
        'lblSelectVendor
        '
        Me.lblSelectVendor.AutoSize = True
        Me.lblSelectVendor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectVendor.Location = New System.Drawing.Point(49, 62)
        Me.lblSelectVendor.Name = "lblSelectVendor"
        Me.lblSelectVendor.Size = New System.Drawing.Size(87, 13)
        Me.lblSelectVendor.TabIndex = 4
        Me.lblSelectVendor.Text = "Select Vendor"
        '
        'ComboBoxVendors
        '
        Me.ComboBoxVendors.FormattingEnabled = True
        Me.ComboBoxVendors.Location = New System.Drawing.Point(142, 62)
        Me.ComboBoxVendors.Name = "ComboBoxVendors"
        Me.ComboBoxVendors.Size = New System.Drawing.Size(322, 21)
        Me.ComboBoxVendors.TabIndex = 5
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(478, 62)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(76, 23)
        Me.BtnOK.TabIndex = 7
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'txtdc
        '
        Me.txtdc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdc.Location = New System.Drawing.Point(728, 13)
        Me.txtdc.Name = "txtdc"
        Me.txtdc.Size = New System.Drawing.Size(100, 21)
        Me.txtdc.TabIndex = 8
        '
        'txtHDKEY
        '
        Me.txtHDKEY.Location = New System.Drawing.Point(951, 63)
        Me.txtHDKEY.Name = "txtHDKEY"
        Me.txtHDKEY.Size = New System.Drawing.Size(18, 20)
        Me.txtHDKEY.TabIndex = 9
        Me.txtHDKEY.Visible = False
        '
        'txtPONumber
        '
        Me.txtPONumber.Location = New System.Drawing.Point(854, 62)
        Me.txtPONumber.Name = "txtPONumber"
        Me.txtPONumber.Size = New System.Drawing.Size(91, 20)
        Me.txtPONumber.TabIndex = 10
        Me.txtPONumber.Visible = False
        '
        'lblDC
        '
        Me.lblDC.AutoSize = True
        Me.lblDC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDC.Location = New System.Drawing.Point(635, 13)
        Me.lblDC.Name = "lblDC"
        Me.lblDC.Size = New System.Drawing.Size(87, 16)
        Me.lblDC.TabIndex = 11
        Me.lblDC.Text = "DC Number"
        '
        'lblDCDt
        '
        Me.lblDCDt.AutoSize = True
        Me.lblDCDt.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDCDt.Location = New System.Drawing.Point(834, 13)
        Me.lblDCDt.Name = "lblDCDt"
        Me.lblDCDt.Size = New System.Drawing.Size(41, 16)
        Me.lblDCDt.TabIndex = 12
        Me.lblDCDt.Text = "Date"
        '
        'GBSC
        '
        Me.GBSC.Controls.Add(Me.lblLineSel)
        Me.GBSC.Controls.Add(Me.txtchkcount)
        Me.GBSC.Controls.Add(Me.DateTimeDCDT)
        Me.GBSC.Controls.Add(Me.lblDCDt)
        Me.GBSC.Controls.Add(Me.lblDC)
        Me.GBSC.Controls.Add(Me.txtPONumber)
        Me.GBSC.Controls.Add(Me.txtHDKEY)
        Me.GBSC.Controls.Add(Me.txtdc)
        Me.GBSC.Controls.Add(Me.BtnOK)
        Me.GBSC.Controls.Add(Me.ComboBoxVendors)
        Me.GBSC.Controls.Add(Me.lblSelectVendor)
        Me.GBSC.Controls.Add(Me.RBView)
        Me.GBSC.Controls.Add(Me.RBCancel)
        Me.GBSC.Controls.Add(Me.RBDelDC)
        Me.GBSC.Controls.Add(Me.RBNewDC)
        Me.GBSC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GBSC.Location = New System.Drawing.Point(28, 15)
        Me.GBSC.Name = "GBSC"
        Me.GBSC.Size = New System.Drawing.Size(989, 88)
        Me.GBSC.TabIndex = 0
        Me.GBSC.TabStop = False
        Me.GBSC.Text = "SubContractor DC Maintenance"
        '
        'lblLineSel
        '
        Me.lblLineSel.AutoSize = True
        Me.lblLineSel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLineSel.Location = New System.Drawing.Point(579, 65)
        Me.lblLineSel.Name = "lblLineSel"
        Me.lblLineSel.Size = New System.Drawing.Size(91, 13)
        Me.lblLineSel.TabIndex = 15
        Me.lblLineSel.Text = "Lines Selected"
        '
        'txtchkcount
        '
        Me.txtchkcount.Location = New System.Drawing.Point(676, 63)
        Me.txtchkcount.Name = "txtchkcount"
        Me.txtchkcount.Size = New System.Drawing.Size(85, 20)
        Me.txtchkcount.TabIndex = 14
        '
        'DateTimeDCDT
        '
        Me.DateTimeDCDT.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimeDCDT.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimeDCDT.Location = New System.Drawing.Point(881, 13)
        Me.DateTimeDCDT.Name = "DateTimeDCDT"
        Me.DateTimeDCDT.Size = New System.Drawing.Size(98, 21)
        Me.DateTimeDCDT.TabIndex = 13
        '
        'SUBDC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1436, 751)
        Me.Controls.Add(Me.GroupBoxPODetails)
        Me.Controls.Add(Me.GBSC)
        Me.Name = "SUBDC"
        Me.Text = "Sub Contractor DC"
        CType(Me.DataGridViewSCDC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxPODetails.ResumeLayout(False)
        Me.GroupBoxPODetails.PerformLayout()
        Me.GBSC.ResumeLayout(False)
        Me.GBSC.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridViewSCDC As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBoxPODetails As System.Windows.Forms.GroupBox
    Friend WithEvents BTNDCSAVE As System.Windows.Forms.Button
    Friend WithEvents txtRemark2 As System.Windows.Forms.TextBox
    Friend WithEvents txtRemarks1 As System.Windows.Forms.TextBox
    Friend WithEvents LblRemarks2 As System.Windows.Forms.Label
    Friend WithEvents LblRemarks1 As System.Windows.Forms.Label
    Friend WithEvents RBNewDC As System.Windows.Forms.RadioButton
    Friend WithEvents RBDelDC As System.Windows.Forms.RadioButton
    Friend WithEvents RBCancel As System.Windows.Forms.RadioButton
    Friend WithEvents RBView As System.Windows.Forms.RadioButton
    Friend WithEvents lblSelectVendor As System.Windows.Forms.Label
    Friend WithEvents ComboBoxVendors As System.Windows.Forms.ComboBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents txtdc As System.Windows.Forms.TextBox
    Friend WithEvents txtHDKEY As System.Windows.Forms.TextBox
    Friend WithEvents txtPONumber As System.Windows.Forms.TextBox
    Friend WithEvents lblDC As System.Windows.Forms.Label
    Friend WithEvents lblDCDt As System.Windows.Forms.Label
    Friend WithEvents GBSC As System.Windows.Forms.GroupBox
    Friend WithEvents DateTimeDCDT As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtchkcount As System.Windows.Forms.TextBox
    Friend WithEvents lblLineSel As System.Windows.Forms.Label
End Class
