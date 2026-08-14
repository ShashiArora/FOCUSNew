<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Receipts
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
        Me.lblScan = New System.Windows.Forms.Label()
        Me.txtGRNscan = New System.Windows.Forms.TextBox()
        Me.btnImageClear = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'lblScan
        '
        Me.lblScan.AutoSize = True
        Me.lblScan.Location = New System.Drawing.Point(23, 30)
        Me.lblScan.Name = "lblScan"
        Me.lblScan.Size = New System.Drawing.Size(106, 17)
        Me.lblScan.TabIndex = 0
        Me.lblScan.Text = "Scan the Image"
        '
        'txtGRNscan
        '
        Me.txtGRNscan.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGRNscan.Location = New System.Drawing.Point(147, 30)
        Me.txtGRNscan.Name = "txtGRNscan"
        Me.txtGRNscan.Size = New System.Drawing.Size(583, 47)
        Me.txtGRNscan.TabIndex = 1
        '
        'btnImageClear
        '
        Me.btnImageClear.Location = New System.Drawing.Point(737, 53)
        Me.btnImageClear.Name = "btnImageClear"
        Me.btnImageClear.Size = New System.Drawing.Size(75, 23)
        Me.btnImageClear.TabIndex = 2
        Me.btnImageClear.Text = "Clear"
        Me.btnImageClear.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 101)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "GRN No."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(23, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "OR"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(147, 83)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(161, 28)
        Me.TextBox1.TabIndex = 5
        '
        'Receipts
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1085, 703)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnImageClear)
        Me.Controls.Add(Me.txtGRNscan)
        Me.Controls.Add(Me.lblScan)
        Me.Name = "Receipts"
        Me.Text = "Receipts"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblScan As System.Windows.Forms.Label
    Friend WithEvents txtGRNscan As System.Windows.Forms.TextBox
    Friend WithEvents btnImageClear As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
