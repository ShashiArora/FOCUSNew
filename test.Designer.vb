<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class test
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
        Me.CheckBoxMale = New System.Windows.Forms.CheckBox()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'CheckBoxMale
        '
        Me.CheckBoxMale.AutoSize = True
        Me.CheckBoxMale.Location = New System.Drawing.Point(35, 27)
        Me.CheckBoxMale.Name = "CheckBoxMale"
        Me.CheckBoxMale.Size = New System.Drawing.Size(69, 24)
        Me.CheckBoxMale.TabIndex = 0
        Me.CheckBoxMale.Text = "Male"
        Me.CheckBoxMale.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(508, 46)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(133, 24)
        Me.RadioButton1.TabIndex = 1
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "RadioButton1"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'test
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1347, 667)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.CheckBoxMale)
        Me.Name = "test"
        Me.Text = "test"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CheckBoxMale As System.Windows.Forms.CheckBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
End Class
