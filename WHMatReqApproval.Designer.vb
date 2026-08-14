<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WHMatRequestApproval
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
        Me.DataGridViewAprSummary = New System.Windows.Forms.DataGridView()
        Me.DataGridViewApprovalDetail = New System.Windows.Forms.DataGridView()
        Me.GroupBoxPendApr = New System.Windows.Forms.GroupBox()
        Me.CheckBoxItems = New System.Windows.Forms.CheckBox()
        Me.btnReject = New System.Windows.Forms.Button()
        Me.btnApprove = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.BtnItemReject = New System.Windows.Forms.Button()
        Me.GroupBoxMenu = New Focus.myGroupBox()
        Me.BtnApproved = New System.Windows.Forms.Button()
        Me.BtnEdit = New System.Windows.Forms.Button()
        Me.BtnMyApproval = New System.Windows.Forms.Button()
        CType(Me.DataGridViewAprSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridViewApprovalDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxPendApr.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBoxMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridViewAprSummary
        '
        Me.DataGridViewAprSummary.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.DataGridViewAprSummary.AllowUserToAddRows = False
        Me.DataGridViewAprSummary.AllowUserToDeleteRows = False
        Me.DataGridViewAprSummary.AllowUserToResizeRows = False
        Me.DataGridViewAprSummary.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.DataGridViewAprSummary.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewAprSummary.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridViewAprSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewAprSummary.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewAprSummary.GridColor = System.Drawing.Color.Gray
        Me.DataGridViewAprSummary.Location = New System.Drawing.Point(5, 18)
        Me.DataGridViewAprSummary.Name = "DataGridViewAprSummary"
        Me.DataGridViewAprSummary.RowHeadersWidth = 56
        Me.DataGridViewAprSummary.Size = New System.Drawing.Size(1036, 199)
        Me.DataGridViewAprSummary.TabIndex = 243
        '
        'DataGridViewApprovalDetail
        '
        Me.DataGridViewApprovalDetail.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.DataGridViewApprovalDetail.AllowUserToAddRows = False
        Me.DataGridViewApprovalDetail.AllowUserToDeleteRows = False
        Me.DataGridViewApprovalDetail.AllowUserToResizeRows = False
        Me.DataGridViewApprovalDetail.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.DataGridViewApprovalDetail.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewApprovalDetail.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewApprovalDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewApprovalDetail.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewApprovalDetail.GridColor = System.Drawing.Color.Gray
        Me.DataGridViewApprovalDetail.Location = New System.Drawing.Point(9, 18)
        Me.DataGridViewApprovalDetail.Name = "DataGridViewApprovalDetail"
        Me.DataGridViewApprovalDetail.RowHeadersWidth = 56
        Me.DataGridViewApprovalDetail.Size = New System.Drawing.Size(1026, 277)
        Me.DataGridViewApprovalDetail.TabIndex = 244
        '
        'GroupBoxPendApr
        '
        Me.GroupBoxPendApr.Controls.Add(Me.CheckBoxItems)
        Me.GroupBoxPendApr.Controls.Add(Me.btnReject)
        Me.GroupBoxPendApr.Controls.Add(Me.btnApprove)
        Me.GroupBoxPendApr.Controls.Add(Me.DataGridViewAprSummary)
        Me.GroupBoxPendApr.Location = New System.Drawing.Point(8, 47)
        Me.GroupBoxPendApr.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPendApr.Name = "GroupBoxPendApr"
        Me.GroupBoxPendApr.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPendApr.Size = New System.Drawing.Size(1058, 244)
        Me.GroupBoxPendApr.TabIndex = 251
        Me.GroupBoxPendApr.TabStop = False
        Me.GroupBoxPendApr.Text = "Pending Approvals"
        '
        'CheckBoxItems
        '
        Me.CheckBoxItems.AutoSize = True
        Me.CheckBoxItems.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxItems.Location = New System.Drawing.Point(4, 223)
        Me.CheckBoxItems.Margin = New System.Windows.Forms.Padding(2)
        Me.CheckBoxItems.Name = "CheckBoxItems"
        Me.CheckBoxItems.Size = New System.Drawing.Size(93, 17)
        Me.CheckBoxItems.TabIndex = 246
        Me.CheckBoxItems.Text = "Item Details"
        Me.CheckBoxItems.UseVisualStyleBackColor = True
        '
        'btnReject
        '
        Me.btnReject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReject.Location = New System.Drawing.Point(961, 221)
        Me.btnReject.Margin = New System.Windows.Forms.Padding(2)
        Me.btnReject.Name = "btnReject"
        Me.btnReject.Size = New System.Drawing.Size(78, 23)
        Me.btnReject.TabIndex = 245
        Me.btnReject.Text = "Reject"
        Me.btnReject.UseVisualStyleBackColor = True
        '
        'btnApprove
        '
        Me.btnApprove.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnApprove.Location = New System.Drawing.Point(868, 221)
        Me.btnApprove.Margin = New System.Windows.Forms.Padding(2)
        Me.btnApprove.Name = "btnApprove"
        Me.btnApprove.Size = New System.Drawing.Size(89, 23)
        Me.btnApprove.TabIndex = 244
        Me.btnApprove.Text = "Approve"
        Me.btnApprove.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataGridViewApprovalDetail)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 295)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(1057, 317)
        Me.GroupBox1.TabIndex = 252
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Items"
        '
        'BtnItemReject
        '
        Me.BtnItemReject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnItemReject.Location = New System.Drawing.Point(1053, 550)
        Me.BtnItemReject.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnItemReject.Name = "BtnItemReject"
        Me.BtnItemReject.Size = New System.Drawing.Size(28, 23)
        Me.BtnItemReject.TabIndex = 246
        Me.BtnItemReject.Text = "Item Reject"
        Me.BtnItemReject.UseVisualStyleBackColor = True
        Me.BtnItemReject.Visible = False
        '
        'GroupBoxMenu
        '
        Me.GroupBoxMenu.BackColor = System.Drawing.Color.SteelBlue
        Me.GroupBoxMenu.BorderColor = System.Drawing.Color.Black
        Me.GroupBoxMenu.Controls.Add(Me.BtnApproved)
        Me.GroupBoxMenu.Controls.Add(Me.BtnEdit)
        Me.GroupBoxMenu.Controls.Add(Me.BtnMyApproval)
        Me.GroupBoxMenu.Location = New System.Drawing.Point(9, 6)
        Me.GroupBoxMenu.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxMenu.Name = "GroupBoxMenu"
        Me.GroupBoxMenu.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxMenu.Size = New System.Drawing.Size(1061, 35)
        Me.GroupBoxMenu.TabIndex = 250
        Me.GroupBoxMenu.TabStop = False
        '
        'BtnApproved
        '
        Me.BtnApproved.BackColor = System.Drawing.Color.LightGray
        Me.BtnApproved.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnApproved.Location = New System.Drawing.Point(4, 4)
        Me.BtnApproved.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnApproved.Name = "BtnApproved"
        Me.BtnApproved.Size = New System.Drawing.Size(134, 27)
        Me.BtnApproved.TabIndex = 2
        Me.BtnApproved.Text = "Approvals Completed"
        Me.BtnApproved.UseVisualStyleBackColor = False
        '
        'BtnEdit
        '
        Me.BtnEdit.BackColor = System.Drawing.Color.LightGray
        Me.BtnEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnEdit.Location = New System.Drawing.Point(898, 3)
        Me.BtnEdit.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnEdit.Name = "BtnEdit"
        Me.BtnEdit.Size = New System.Drawing.Size(159, 27)
        Me.BtnEdit.TabIndex = 1
        Me.BtnEdit.Text = "Substitute Approval"
        Me.BtnEdit.UseVisualStyleBackColor = False
        '
        'BtnMyApproval
        '
        Me.BtnMyApproval.BackColor = System.Drawing.Color.LightGray
        Me.BtnMyApproval.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnMyApproval.Location = New System.Drawing.Point(142, 5)
        Me.BtnMyApproval.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnMyApproval.Name = "BtnMyApproval"
        Me.BtnMyApproval.Size = New System.Drawing.Size(172, 28)
        Me.BtnMyApproval.TabIndex = 0
        Me.BtnMyApproval.Text = "Pending- My Approvals"
        Me.BtnMyApproval.UseVisualStyleBackColor = False
        '
        'WHMatRequestApproval
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1081, 623)
        Me.Controls.Add(Me.BtnItemReject)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBoxPendApr)
        Me.Controls.Add(Me.GroupBoxMenu)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "WHMatRequestApproval"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Material Request Approval"
        CType(Me.DataGridViewAprSummary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridViewApprovalDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxPendApr.ResumeLayout(False)
        Me.GroupBoxPendApr.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBoxMenu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridViewAprSummary As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewApprovalDetail As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBoxMenu As Focus.myGroupBox
    Friend WithEvents BtnEdit As System.Windows.Forms.Button
    Friend WithEvents BtnMyApproval As System.Windows.Forms.Button
    Friend WithEvents BtnApproved As System.Windows.Forms.Button
    Friend WithEvents GroupBoxPendApr As System.Windows.Forms.GroupBox
    Friend WithEvents btnApprove As System.Windows.Forms.Button
    Friend WithEvents btnReject As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBoxItems As System.Windows.Forms.CheckBox
    Friend WithEvents BtnItemReject As System.Windows.Forms.Button
End Class
