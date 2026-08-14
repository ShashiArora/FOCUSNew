Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms

Public Class RFQCompleted

    Inherits System.Windows.Forms.Form

    Private ConnectionString As String
    Public stockDA As SqlDataAdapter = New SqlDataAdapter

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents datagridRFQCompleted As System.Windows.Forms.DataGrid
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonNew As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonExisting As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBoxCustomerClass As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpenqduedt As System.Windows.Forms.DateTimePicker
    Friend WithEvents CheckedListBoxCertificate As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSpecialInst As System.Windows.Forms.TextBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents txtdocdetails As System.Windows.Forms.TextBox
    Friend WithEvents RadioButtondocno As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtondocyes As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBoxcsr As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxisr As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxtssiseg As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxtssseg As System.Windows.Forms.ComboBox
    Friend WithEvents DataUpdation As System.Windows.Forms.GroupBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtvendquote As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtvendorref As System.Windows.Forms.TextBox
    Friend WithEvents txtvendorcontact As System.Windows.Forms.TextBox
    Friend WithEvents txtspecial As System.Windows.Forms.TextBox
    Friend WithEvents l As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBoxPriceStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents RadioButton3P As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonGroup As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBoxuom As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxItemSource As System.Windows.Forms.ComboBox
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents txtRecVend As System.Windows.Forms.TextBox
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents txtDimension As System.Windows.Forms.TextBox
    Friend WithEvents txtCustDesc As System.Windows.Forms.TextBox
    Friend WithEvents ComboBoxFSYesNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtCustPart As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterial As System.Windows.Forms.TextBox
    Friend WithEvents txtDetailSpecial As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtstockavble As System.Windows.Forms.TextBox
    Friend WithEvents txttooling As System.Windows.Forms.TextBox
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents txtdetailremarks As System.Windows.Forms.TextBox
    Friend WithEvents txtrecno As System.Windows.Forms.TextBox
    Friend WithEvents txtregno As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.datagridRFQCompleted = New System.Windows.Forms.DataGrid
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtregno = New System.Windows.Forms.TextBox
        Me.txtrecno = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RadioButtonExisting = New System.Windows.Forms.RadioButton
        Me.RadioButtonNew = New System.Windows.Forms.RadioButton
        Me.ComboBoxCustomerClass = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpenqduedt = New System.Windows.Forms.DateTimePicker
        Me.CheckedListBoxCertificate = New System.Windows.Forms.CheckedListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSpecialInst = New System.Windows.Forms.TextBox
        Me.Label35 = New System.Windows.Forms.Label
        Me.txtdocdetails = New System.Windows.Forms.TextBox
        Me.RadioButtondocno = New System.Windows.Forms.RadioButton
        Me.RadioButtondocyes = New System.Windows.Forms.RadioButton
        Me.ComboBoxcsr = New System.Windows.Forms.ComboBox
        Me.ComboBoxisr = New System.Windows.Forms.ComboBox
        Me.ComboBoxtssiseg = New System.Windows.Forms.ComboBox
        Me.ComboBoxtssseg = New System.Windows.Forms.ComboBox
        Me.DataUpdation = New System.Windows.Forms.GroupBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtvendquote = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtvendorref = New System.Windows.Forms.TextBox
        Me.txtvendorcontact = New System.Windows.Forms.TextBox
        Me.txtspecial = New System.Windows.Forms.TextBox
        Me.l = New System.Windows.Forms.GroupBox
        Me.ComboBoxPriceStatus = New System.Windows.Forms.ComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.RadioButton3P = New System.Windows.Forms.RadioButton
        Me.RadioButtonGroup = New System.Windows.Forms.RadioButton
        Me.ComboBoxuom = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label54 = New System.Windows.Forms.Label
        Me.ComboBoxItemSource = New System.Windows.Forms.ComboBox
        Me.Label46 = New System.Windows.Forms.Label
        Me.txtRecVend = New System.Windows.Forms.TextBox
        Me.Label45 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label42 = New System.Windows.Forms.Label
        Me.txtDimension = New System.Windows.Forms.TextBox
        Me.txtCustDesc = New System.Windows.Forms.TextBox
        Me.ComboBoxFSYesNo = New System.Windows.Forms.ComboBox
        Me.Label40 = New System.Windows.Forms.Label
        Me.Label39 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtCustPart = New System.Windows.Forms.TextBox
        Me.txtMaterial = New System.Windows.Forms.TextBox
        Me.txtDetailSpecial = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtstockavble = New System.Windows.Forms.TextBox
        Me.txttooling = New System.Windows.Forms.TextBox
        Me.Label43 = New System.Windows.Forms.Label
        Me.txtdetailremarks = New System.Windows.Forms.TextBox
        CType(Me.datagridRFQCompleted, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.DataUpdation.SuspendLayout()
        Me.SuspendLayout()
        '
        'datagridRFQCompleted
        '
        Me.datagridRFQCompleted.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(192, Byte))
        Me.datagridRFQCompleted.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridRFQCompleted.CaptionVisible = False
        Me.datagridRFQCompleted.DataMember = ""
        Me.datagridRFQCompleted.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridRFQCompleted.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridRFQCompleted.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.datagridRFQCompleted.Location = New System.Drawing.Point(8, 16)
        Me.datagridRFQCompleted.Name = "datagridRFQCompleted"
        Me.datagridRFQCompleted.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.datagridRFQCompleted.ParentRowsVisible = False
        Me.datagridRFQCompleted.PreferredColumnWidth = 85
        Me.datagridRFQCompleted.ReadOnly = True
        Me.datagridRFQCompleted.RowHeadersVisible = False
        Me.datagridRFQCompleted.Size = New System.Drawing.Size(1272, 384)
        Me.datagridRFQCompleted.TabIndex = 18
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.GroupBox3.Controls.Add(Me.txtregno)
        Me.GroupBox3.Controls.Add(Me.txtrecno)
        Me.GroupBox3.Controls.Add(Me.GroupBox1)
        Me.GroupBox3.Controls.Add(Me.ComboBoxCustomerClass)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.dtpenqduedt)
        Me.GroupBox3.Controls.Add(Me.CheckedListBoxCertificate)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.txtSpecialInst)
        Me.GroupBox3.Controls.Add(Me.Label35)
        Me.GroupBox3.Controls.Add(Me.txtdocdetails)
        Me.GroupBox3.Controls.Add(Me.RadioButtondocno)
        Me.GroupBox3.Controls.Add(Me.RadioButtondocyes)
        Me.GroupBox3.Controls.Add(Me.ComboBoxcsr)
        Me.GroupBox3.Controls.Add(Me.ComboBoxisr)
        Me.GroupBox3.Controls.Add(Me.ComboBoxtssiseg)
        Me.GroupBox3.Controls.Add(Me.ComboBoxtssseg)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.Black
        Me.GroupBox3.Location = New System.Drawing.Point(8, 408)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1272, 104)
        Me.GroupBox3.TabIndex = 117
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Customer Details"
        '
        'txtregno
        '
        Me.txtregno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtregno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtregno.Enabled = False
        Me.txtregno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtregno.Location = New System.Drawing.Point(688, 0)
        Me.txtregno.MaxLength = 50
        Me.txtregno.Name = "txtregno"
        Me.txtregno.Size = New System.Drawing.Size(20, 20)
        Me.txtregno.TabIndex = 167
        Me.txtregno.Text = ""
        Me.txtregno.Visible = False
        '
        'txtrecno
        '
        Me.txtrecno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtrecno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtrecno.Enabled = False
        Me.txtrecno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtrecno.Location = New System.Drawing.Point(640, 0)
        Me.txtrecno.MaxLength = 50
        Me.txtrecno.Name = "txtrecno"
        Me.txtrecno.Size = New System.Drawing.Size(20, 20)
        Me.txtrecno.TabIndex = 166
        Me.txtrecno.Text = ""
        Me.txtrecno.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButtonExisting)
        Me.GroupBox1.Controls.Add(Me.RadioButtonNew)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(144, 32)
        Me.GroupBox1.TabIndex = 165
        Me.GroupBox1.TabStop = False
        '
        'RadioButtonExisting
        '
        Me.RadioButtonExisting.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonExisting.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonExisting.Location = New System.Drawing.Point(16, 8)
        Me.RadioButtonExisting.Name = "RadioButtonExisting"
        Me.RadioButtonExisting.Size = New System.Drawing.Size(64, 16)
        Me.RadioButtonExisting.TabIndex = 139
        Me.RadioButtonExisting.Text = "Existing"
        '
        'RadioButtonNew
        '
        Me.RadioButtonNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonNew.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonNew.Location = New System.Drawing.Point(88, 8)
        Me.RadioButtonNew.Name = "RadioButtonNew"
        Me.RadioButtonNew.Size = New System.Drawing.Size(48, 16)
        Me.RadioButtonNew.TabIndex = 140
        Me.RadioButtonNew.Text = "New"
        '
        'ComboBoxCustomerClass
        '
        Me.ComboBoxCustomerClass.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxCustomerClass.Location = New System.Drawing.Point(1152, 48)
        Me.ComboBoxCustomerClass.Name = "ComboBoxCustomerClass"
        Me.ComboBoxCustomerClass.Size = New System.Drawing.Size(64, 22)
        Me.ComboBoxCustomerClass.TabIndex = 164
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(8, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 163
        Me.Label3.Text = "Enq. Due Dt."
        '
        'dtpenqduedt
        '
        Me.dtpenqduedt.Checked = False
        Me.dtpenqduedt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpenqduedt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpenqduedt.Location = New System.Drawing.Point(8, 72)
        Me.dtpenqduedt.Name = "dtpenqduedt"
        Me.dtpenqduedt.ShowCheckBox = True
        Me.dtpenqduedt.Size = New System.Drawing.Size(128, 22)
        Me.dtpenqduedt.TabIndex = 162
        '
        'CheckedListBoxCertificate
        '
        Me.CheckedListBoxCertificate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.CheckedListBoxCertificate.Location = New System.Drawing.Point(896, 24)
        Me.CheckedListBoxCertificate.Name = "CheckedListBoxCertificate"
        Me.CheckedListBoxCertificate.Size = New System.Drawing.Size(248, 64)
        Me.CheckedListBoxCertificate.TabIndex = 160
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(504, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 16)
        Me.Label1.TabIndex = 147
        Me.Label1.Text = "Special Instructions"
        '
        'txtSpecialInst
        '
        Me.txtSpecialInst.AutoSize = False
        Me.txtSpecialInst.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSpecialInst.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpecialInst.Location = New System.Drawing.Point(504, 24)
        Me.txtSpecialInst.Multiline = True
        Me.txtSpecialInst.Name = "txtSpecialInst"
        Me.txtSpecialInst.Size = New System.Drawing.Size(384, 64)
        Me.txtSpecialInst.TabIndex = 146
        Me.txtSpecialInst.Text = ""
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label35.Enabled = False
        Me.Label35.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.ForeColor = System.Drawing.Color.Black
        Me.Label35.Location = New System.Drawing.Point(168, 8)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(120, 16)
        Me.Label35.TabIndex = 135
        Me.Label35.Text = "Document uploaded:"
        '
        'txtdocdetails
        '
        Me.txtdocdetails.AutoSize = False
        Me.txtdocdetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtdocdetails.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdocdetails.ForeColor = System.Drawing.Color.Red
        Me.txtdocdetails.Location = New System.Drawing.Point(168, 24)
        Me.txtdocdetails.Multiline = True
        Me.txtdocdetails.Name = "txtdocdetails"
        Me.txtdocdetails.Size = New System.Drawing.Size(312, 64)
        Me.txtdocdetails.TabIndex = 134
        Me.txtdocdetails.Text = ""
        '
        'RadioButtondocno
        '
        Me.RadioButtondocno.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.RadioButtondocno.Enabled = False
        Me.RadioButtondocno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtondocno.ForeColor = System.Drawing.Color.Black
        Me.RadioButtondocno.Location = New System.Drawing.Point(352, 8)
        Me.RadioButtondocno.Name = "RadioButtondocno"
        Me.RadioButtondocno.Size = New System.Drawing.Size(64, 16)
        Me.RadioButtondocno.TabIndex = 114
        Me.RadioButtondocno.Text = "No"
        '
        'RadioButtondocyes
        '
        Me.RadioButtondocyes.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.RadioButtondocyes.Enabled = False
        Me.RadioButtondocyes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtondocyes.ForeColor = System.Drawing.Color.Black
        Me.RadioButtondocyes.Location = New System.Drawing.Point(288, 8)
        Me.RadioButtondocyes.Name = "RadioButtondocyes"
        Me.RadioButtondocyes.Size = New System.Drawing.Size(48, 16)
        Me.RadioButtondocyes.TabIndex = 113
        Me.RadioButtondocyes.Text = "Yes"
        '
        'ComboBoxcsr
        '
        Me.ComboBoxcsr.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxcsr.Location = New System.Drawing.Point(1152, 24)
        Me.ComboBoxcsr.Name = "ComboBoxcsr"
        Me.ComboBoxcsr.Size = New System.Drawing.Size(56, 22)
        Me.ComboBoxcsr.TabIndex = 107
        Me.ComboBoxcsr.Text = "CSR"
        '
        'ComboBoxisr
        '
        Me.ComboBoxisr.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxisr.Location = New System.Drawing.Point(1208, 24)
        Me.ComboBoxisr.Name = "ComboBoxisr"
        Me.ComboBoxisr.Size = New System.Drawing.Size(56, 22)
        Me.ComboBoxisr.TabIndex = 110
        Me.ComboBoxisr.Text = "ISR"
        '
        'ComboBoxtssiseg
        '
        Me.ComboBoxtssiseg.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxtssiseg.Location = New System.Drawing.Point(1152, 72)
        Me.ComboBoxtssiseg.Name = "ComboBoxtssiseg"
        Me.ComboBoxtssiseg.Size = New System.Drawing.Size(64, 22)
        Me.ComboBoxtssiseg.TabIndex = 109
        Me.ComboBoxtssiseg.Text = "TSSI SEG"
        '
        'ComboBoxtssseg
        '
        Me.ComboBoxtssseg.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxtssseg.Location = New System.Drawing.Point(1216, 72)
        Me.ComboBoxtssseg.Name = "ComboBoxtssseg"
        Me.ComboBoxtssseg.Size = New System.Drawing.Size(48, 22)
        Me.ComboBoxtssseg.TabIndex = 108
        Me.ComboBoxtssseg.Text = "SEGMENT"
        '
        'DataUpdation
        '
        Me.DataUpdation.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.DataUpdation.Controls.Add(Me.Label18)
        Me.DataUpdation.Controls.Add(Me.Label19)
        Me.DataUpdation.Controls.Add(Me.txtvendquote)
        Me.DataUpdation.Controls.Add(Me.Label11)
        Me.DataUpdation.Controls.Add(Me.Label12)
        Me.DataUpdation.Controls.Add(Me.txtvendorref)
        Me.DataUpdation.Controls.Add(Me.txtvendorcontact)
        Me.DataUpdation.Controls.Add(Me.txtspecial)
        Me.DataUpdation.Controls.Add(Me.l)
        Me.DataUpdation.Controls.Add(Me.ComboBoxPriceStatus)
        Me.DataUpdation.Controls.Add(Me.Label17)
        Me.DataUpdation.Controls.Add(Me.RadioButton3P)
        Me.DataUpdation.Controls.Add(Me.RadioButtonGroup)
        Me.DataUpdation.Controls.Add(Me.ComboBoxuom)
        Me.DataUpdation.Controls.Add(Me.Label10)
        Me.DataUpdation.Controls.Add(Me.Label54)
        Me.DataUpdation.Controls.Add(Me.ComboBoxItemSource)
        Me.DataUpdation.Controls.Add(Me.Label46)
        Me.DataUpdation.Controls.Add(Me.txtRecVend)
        Me.DataUpdation.Controls.Add(Me.Label45)
        Me.DataUpdation.Controls.Add(Me.Label15)
        Me.DataUpdation.Controls.Add(Me.Label42)
        Me.DataUpdation.Controls.Add(Me.txtDimension)
        Me.DataUpdation.Controls.Add(Me.txtCustDesc)
        Me.DataUpdation.Controls.Add(Me.ComboBoxFSYesNo)
        Me.DataUpdation.Controls.Add(Me.Label40)
        Me.DataUpdation.Controls.Add(Me.Label39)
        Me.DataUpdation.Controls.Add(Me.Label16)
        Me.DataUpdation.Controls.Add(Me.txtCustPart)
        Me.DataUpdation.Controls.Add(Me.txtMaterial)
        Me.DataUpdation.Controls.Add(Me.txtDetailSpecial)
        Me.DataUpdation.Controls.Add(Me.Label13)
        Me.DataUpdation.Controls.Add(Me.Label14)
        Me.DataUpdation.Controls.Add(Me.txtstockavble)
        Me.DataUpdation.Controls.Add(Me.txttooling)
        Me.DataUpdation.Controls.Add(Me.Label43)
        Me.DataUpdation.Controls.Add(Me.txtdetailremarks)
        Me.DataUpdation.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataUpdation.ForeColor = System.Drawing.Color.Firebrick
        Me.DataUpdation.Location = New System.Drawing.Point(8, 520)
        Me.DataUpdation.Name = "DataUpdation"
        Me.DataUpdation.Size = New System.Drawing.Size(1272, 176)
        Me.DataUpdation.TabIndex = 116
        Me.DataUpdation.TabStop = False
        Me.DataUpdation.Text = "Details"
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Red
        Me.Label18.Location = New System.Drawing.Point(624, 128)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(104, 16)
        Me.Label18.TabIndex = 206
        Me.Label18.Text = "Special Remarks"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Red
        Me.Label19.Location = New System.Drawing.Point(376, 128)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(184, 16)
        Me.Label19.TabIndex = 205
        Me.Label19.Text = "Vendor Quote Reference"
        '
        'txtvendquote
        '
        Me.txtvendquote.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvendquote.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtvendquote.Location = New System.Drawing.Point(376, 144)
        Me.txtvendquote.MaxLength = 100
        Me.txtvendquote.Name = "txtvendquote"
        Me.txtvendquote.Size = New System.Drawing.Size(240, 20)
        Me.txtvendquote.TabIndex = 18
        Me.txtvendquote.Text = ""
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Red
        Me.Label11.Location = New System.Drawing.Point(144, 128)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(152, 16)
        Me.Label11.TabIndex = 203
        Me.Label11.Text = "Contact person / Designation"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Red
        Me.Label12.Location = New System.Drawing.Point(8, 128)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(136, 16)
        Me.Label12.TabIndex = 202
        Me.Label12.Text = "Vendor ID/ Name"
        '
        'txtvendorref
        '
        Me.txtvendorref.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvendorref.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtvendorref.Location = New System.Drawing.Point(8, 144)
        Me.txtvendorref.MaxLength = 100
        Me.txtvendorref.Name = "txtvendorref"
        Me.txtvendorref.Size = New System.Drawing.Size(128, 20)
        Me.txtvendorref.TabIndex = 16
        Me.txtvendorref.Text = ""
        '
        'txtvendorcontact
        '
        Me.txtvendorcontact.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvendorcontact.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtvendorcontact.Location = New System.Drawing.Point(144, 144)
        Me.txtvendorcontact.MaxLength = 100
        Me.txtvendorcontact.Name = "txtvendorcontact"
        Me.txtvendorcontact.Size = New System.Drawing.Size(224, 20)
        Me.txtvendorcontact.TabIndex = 17
        Me.txtvendorcontact.Text = ""
        '
        'txtspecial
        '
        Me.txtspecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtspecial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtspecial.Location = New System.Drawing.Point(624, 144)
        Me.txtspecial.MaxLength = 100
        Me.txtspecial.Name = "txtspecial"
        Me.txtspecial.Size = New System.Drawing.Size(448, 20)
        Me.txtspecial.TabIndex = 19
        Me.txtspecial.Text = ""
        '
        'l
        '
        Me.l.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(192, Byte))
        Me.l.Location = New System.Drawing.Point(-8, 88)
        Me.l.Name = "l"
        Me.l.Size = New System.Drawing.Size(1280, 8)
        Me.l.TabIndex = 197
        Me.l.TabStop = False
        '
        'ComboBoxPriceStatus
        '
        Me.ComboBoxPriceStatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxPriceStatus.Items.AddRange(New Object() {"Accepted-Line Open", "Accepted-Line Closed", "Rejected"})
        Me.ComboBoxPriceStatus.Location = New System.Drawing.Point(1080, 144)
        Me.ComboBoxPriceStatus.Name = "ComboBoxPriceStatus"
        Me.ComboBoxPriceStatus.Size = New System.Drawing.Size(168, 22)
        Me.ComboBoxPriceStatus.TabIndex = 20
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Blue
        Me.Label17.Location = New System.Drawing.Point(8, 104)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(56, 16)
        Me.Label17.TabIndex = 192
        Me.Label17.Text = "Price from"
        '
        'RadioButton3P
        '
        Me.RadioButton3P.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.RadioButton3P.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton3P.ForeColor = System.Drawing.Color.Blue
        Me.RadioButton3P.Location = New System.Drawing.Point(136, 104)
        Me.RadioButton3P.Name = "RadioButton3P"
        Me.RadioButton3P.Size = New System.Drawing.Size(72, 16)
        Me.RadioButton3P.TabIndex = 7
        Me.RadioButton3P.Text = "3rd party"
        '
        'RadioButtonGroup
        '
        Me.RadioButtonGroup.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.RadioButtonGroup.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonGroup.ForeColor = System.Drawing.Color.Blue
        Me.RadioButtonGroup.Location = New System.Drawing.Point(72, 104)
        Me.RadioButtonGroup.Name = "RadioButtonGroup"
        Me.RadioButtonGroup.Size = New System.Drawing.Size(56, 16)
        Me.RadioButtonGroup.TabIndex = 6
        Me.RadioButtonGroup.Text = "Group"
        '
        'ComboBoxuom
        '
        Me.ComboBoxuom.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxuom.Items.AddRange(New Object() {"EA", "Pcs", "Sets", "Mtrs", "Cms", "Ltrs", "Ft", "Sheet", "Length"})
        Me.ComboBoxuom.Location = New System.Drawing.Point(1024, 32)
        Me.ComboBoxuom.Name = "ComboBoxuom"
        Me.ComboBoxuom.Size = New System.Drawing.Size(64, 22)
        Me.ComboBoxuom.TabIndex = 166
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(48, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 16)
        Me.Label10.TabIndex = 183
        Me.Label10.Text = "Avbl FS"
        '
        'Label54
        '
        Me.Label54.BackColor = System.Drawing.Color.Transparent
        Me.Label54.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label54.ForeColor = System.Drawing.Color.Black
        Me.Label54.Location = New System.Drawing.Point(128, 16)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(56, 16)
        Me.Label54.TabIndex = 182
        Me.Label54.Text = "Source"
        '
        'ComboBoxItemSource
        '
        Me.ComboBoxItemSource.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxItemSource.Items.AddRange(New Object() {"Procure", "Mfg in Blore"})
        Me.ComboBoxItemSource.Location = New System.Drawing.Point(120, 32)
        Me.ComboBoxItemSource.Name = "ComboBoxItemSource"
        Me.ComboBoxItemSource.Size = New System.Drawing.Size(80, 22)
        Me.ComboBoxItemSource.TabIndex = 161
        '
        'Label46
        '
        Me.Label46.BackColor = System.Drawing.Color.Transparent
        Me.Label46.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label46.ForeColor = System.Drawing.Color.Black
        Me.Label46.Location = New System.Drawing.Point(1112, 16)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(104, 16)
        Me.Label46.TabIndex = 181
        Me.Label46.Text = "Recom. Vendor"
        '
        'txtRecVend
        '
        Me.txtRecVend.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRecVend.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRecVend.Location = New System.Drawing.Point(1088, 32)
        Me.txtRecVend.MaxLength = 50
        Me.txtRecVend.Name = "txtRecVend"
        Me.txtRecVend.Size = New System.Drawing.Size(168, 20)
        Me.txtRecVend.TabIndex = 167
        Me.txtRecVend.Text = ""
        '
        'Label45
        '
        Me.Label45.BackColor = System.Drawing.Color.Transparent
        Me.Label45.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label45.ForeColor = System.Drawing.Color.Black
        Me.Label45.Location = New System.Drawing.Point(1024, 16)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(40, 16)
        Me.Label45.TabIndex = 180
        Me.Label45.Text = "Uom"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(416, 64)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(104, 16)
        Me.Label15.TabIndex = 179
        Me.Label15.Text = "Special Instructions"
        '
        'Label42
        '
        Me.Label42.BackColor = System.Drawing.Color.Transparent
        Me.Label42.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label42.ForeColor = System.Drawing.Color.Black
        Me.Label42.Location = New System.Drawing.Point(8, 64)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(48, 16)
        Me.Label42.TabIndex = 178
        Me.Label42.Text = "Material"
        '
        'txtDimension
        '
        Me.txtDimension.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDimension.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDimension.Location = New System.Drawing.Point(808, 32)
        Me.txtDimension.Name = "txtDimension"
        Me.txtDimension.Size = New System.Drawing.Size(208, 20)
        Me.txtDimension.TabIndex = 168
        Me.txtDimension.Text = ""
        '
        'txtCustDesc
        '
        Me.txtCustDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustDesc.Location = New System.Drawing.Point(528, 32)
        Me.txtCustDesc.MaxLength = 80
        Me.txtCustDesc.Name = "txtCustDesc"
        Me.txtCustDesc.Size = New System.Drawing.Size(272, 20)
        Me.txtCustDesc.TabIndex = 165
        Me.txtCustDesc.Text = ""
        '
        'ComboBoxFSYesNo
        '
        Me.ComboBoxFSYesNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxFSYesNo.Items.AddRange(New Object() {"Yes", "No"})
        Me.ComboBoxFSYesNo.Location = New System.Drawing.Point(8, 32)
        Me.ComboBoxFSYesNo.Name = "ComboBoxFSYesNo"
        Me.ComboBoxFSYesNo.Size = New System.Drawing.Size(104, 22)
        Me.ComboBoxFSYesNo.TabIndex = 160
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(0, Byte), CType(4, Byte))
        Me.Label40.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.ForeColor = System.Drawing.Color.Black
        Me.Label40.Location = New System.Drawing.Point(808, 16)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(64, 16)
        Me.Label40.TabIndex = 176
        Me.Label40.Text = "Dimension"
        '
        'Label39
        '
        Me.Label39.BackColor = System.Drawing.Color.Transparent
        Me.Label39.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.ForeColor = System.Drawing.Color.Black
        Me.Label39.Location = New System.Drawing.Point(216, 16)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(176, 16)
        Me.Label39.TabIndex = 175
        Me.Label39.Text = "Customer Part No."
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.Transparent
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(528, 16)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(152, 16)
        Me.Label16.TabIndex = 174
        Me.Label16.Text = "Cust Part Description"
        '
        'txtCustPart
        '
        Me.txtCustPart.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustPart.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustPart.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustPart.Location = New System.Drawing.Point(216, 32)
        Me.txtCustPart.MaxLength = 50
        Me.txtCustPart.Name = "txtCustPart"
        Me.txtCustPart.Size = New System.Drawing.Size(304, 20)
        Me.txtCustPart.TabIndex = 164
        Me.txtCustPart.Text = ""
        '
        'txtMaterial
        '
        Me.txtMaterial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMaterial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterial.Location = New System.Drawing.Point(56, 64)
        Me.txtMaterial.MaxLength = 80
        Me.txtMaterial.Name = "txtMaterial"
        Me.txtMaterial.Size = New System.Drawing.Size(360, 20)
        Me.txtMaterial.TabIndex = 169
        Me.txtMaterial.Text = ""
        '
        'txtDetailSpecial
        '
        Me.txtDetailSpecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDetailSpecial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDetailSpecial.Location = New System.Drawing.Point(528, 64)
        Me.txtDetailSpecial.MaxLength = 100
        Me.txtDetailSpecial.Name = "txtDetailSpecial"
        Me.txtDetailSpecial.Size = New System.Drawing.Size(728, 20)
        Me.txtDetailSpecial.TabIndex = 170
        Me.txtDetailSpecial.Text = ""
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Blue
        Me.Label13.Location = New System.Drawing.Point(504, 104)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 16)
        Me.Label13.TabIndex = 158
        Me.Label13.Text = "Tooling Cost"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Blue
        Me.Label14.Location = New System.Drawing.Point(208, 104)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 16)
        Me.Label14.TabIndex = 157
        Me.Label14.Text = "Stock Avlbe"
        '
        'txtstockavble
        '
        Me.txtstockavble.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtstockavble.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtstockavble.Location = New System.Drawing.Point(288, 104)
        Me.txtstockavble.MaxLength = 100
        Me.txtstockavble.Name = "txtstockavble"
        Me.txtstockavble.Size = New System.Drawing.Size(216, 20)
        Me.txtstockavble.TabIndex = 13
        Me.txtstockavble.Text = ""
        '
        'txttooling
        '
        Me.txttooling.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txttooling.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttooling.Location = New System.Drawing.Point(576, 104)
        Me.txttooling.MaxLength = 100
        Me.txttooling.Name = "txttooling"
        Me.txttooling.Size = New System.Drawing.Size(232, 20)
        Me.txttooling.TabIndex = 14
        Me.txttooling.Text = ""
        '
        'Label43
        '
        Me.Label43.BackColor = System.Drawing.Color.Transparent
        Me.Label43.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label43.ForeColor = System.Drawing.Color.Blue
        Me.Label43.Location = New System.Drawing.Point(808, 104)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(48, 16)
        Me.Label43.TabIndex = 138
        Me.Label43.Text = "Remarks"
        '
        'txtdetailremarks
        '
        Me.txtdetailremarks.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtdetailremarks.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdetailremarks.Location = New System.Drawing.Point(856, 104)
        Me.txtdetailremarks.MaxLength = 100
        Me.txtdetailremarks.Name = "txtdetailremarks"
        Me.txtdetailremarks.Size = New System.Drawing.Size(400, 20)
        Me.txtdetailremarks.TabIndex = 15
        Me.txtdetailremarks.Text = ""
        '
        'RFQCompleted
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1320, 701)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.DataUpdation)
        Me.Controls.Add(Me.datagridRFQCompleted)
        Me.Name = "RFQCompleted"
        Me.Text = "RFQ Completed"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.datagridRFQCompleted, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.DataUpdation.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub RFQCompleted_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        listloadCertificate()

        RFQCompleted()

        If rfqcomp = "sales" Then

            txtvendorref.Visible = False
            txtvendorcontact.Visible = False
            txtvendquote.Visible = False
            txtspecial.Visible = False

            Label12.Visible = False
            Label11.Visible = False
            Label19.Visible = False
            Label18.Visible = False

        End If


    End Sub

    Private Sub RFQCompleted()
        datagridRFQCompleted.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConfigurationManager.AppSettings("ConnectionString"))
        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet


        strSQL = "SELECT  SNo,  RegNo, [Reg.Date], CustomerID, CustomerName, PartNumber, PartDescription, Qty_Type, Qty, Price, Currency, Type, MOQ, SPU," & _
        "LeadTime FROM TSS_Enquiry_Price_Completed ORDER BY RegNo, PartNumber, Qty"


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)



        datagridRFQCompleted.DataSource = stockDC.Tables(0)
        cnSQL.Close()
        datagridRFQCompleted.Expand(-1)
        'listloadCertificate()


    End Sub

    Private Sub datagridRFQCompleted_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles datagridRFQCompleted.Navigate

    End Sub

    Private Sub datagridRFQCompleted_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridRFQCompleted.CurrentCellChanged
        Dim b As Integer
        'Dim custid As String
        b = datagridRFQCompleted.CurrentCell.ColumnNumber()

        If b = 0 Then
            'clearpricedetails()
            'datagridRFQCompleted

            txtrecno.Text = datagridRFQCompleted.Item(datagridRFQCompleted.CurrentCell)
            txtregno.Text = datagridRFQCompleted.Item(datagridRFQCompleted.CurrentCell.RowNumber, 1)



            Dim cnSQL1 As SqlConnection = New SqlConnection(ConfigurationManager.AppSettings("ConnectionString"))
            Dim cmSQL1 As SqlCommand
            Dim drSQL1 As SqlDataReader
            Dim strSQL1 As String

            strSQL1 = "SELECT     Class, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, Special_instructions, SlNo, FS_Yes_NO," & _
            "Part_Source,  CustPartNumber, CustPartDescription, uom, RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req, Status," & _
            "Source_Mtrl,  Tooling_Cost, Stock_Avble, Remarks, Vendor_Ref, Name, Vendor_Quote, Special_Remarks FROM   TSS_Enquiry_Price_Completed  where " & _
            "SNo =  " & txtrecno.Text & ""

            cnSQL1.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
            drSQL1 = cmSQL1.ExecuteReader()


            If drSQL1.Read() Then

                ComboBoxCustomerClass.Text = drSQL1.Item(0)
                ComboBoxcsr.Text = drSQL1.Item(1)
                ComboBoxisr.Text = drSQL1.Item(2)
                ComboBoxtssiseg.Text = drSQL1.Item(3)
                ComboBoxtssseg.Text = drSQL1.Item(4)

                If drSQL1.Item(5) = "YES" Then
                    RadioButtonExisting.Checked = True
                    RadioButtonNew.Checked = False
                Else
                    RadioButtonNew.Checked = True
                    RadioButtonExisting.Checked = False

                End If

                If drSQL1.Item(6) = "YES" Then
                    RadioButtondocyes.Checked = True
                    RadioButtondocno.Checked = False
                Else
                    RadioButtondocyes.Checked = False
                    RadioButtondocno.Checked = True
                End If

                txtdocdetails.Text = drSQL1.Item(7)

                txtSpecialInst.Text = drSQL1.Item(8)

                ComboBoxFSYesNo.Text = drSQL1.Item(10)
                ComboBoxItemSource.Text = drSQL1.Item(11)
                txtCustPart.Text = drSQL1.Item(12)
                txtCustDesc.Text = drSQL1.Item(13)
                ComboBoxuom.Text = drSQL1.Item(14)
                txtRecVend.Text = drSQL1.Item(15)
                txtDimension.Text = drSQL1.Item(16)
                txtMaterial.Text = drSQL1.Item(17)
                txtDetailSpecial.Text = drSQL1.Item(18)
                If drSQL1.Item(19) = "01-01-1900" Then
                    dtpenqduedt.Checked = False
                    dtpenqduedt.Value = "01-01-1900"
                Else
                    dtpenqduedt.Checked = False
                    dtpenqduedt.Value = drSQL1.Item(19)

                End If

                ComboBoxPriceStatus.Text = (drSQL1.Item(22))

                If (drSQL1.Item(23)) = "Group" Then
                    RadioButtonGroup.Checked = True
                    RadioButton3P.Checked = False
                ElseIf (drSQL1.Item(23)) = "3rdParty" Then
                    RadioButtonGroup.Checked = False
                    RadioButton3P.Checked = True
                End If

                txtstockavble.Text = (drSQL1.Item(25))
                txttooling.Text = (drSQL1.Item(24))

                txtdetailremarks.Text = (drSQL1.Item(26))


                If rfqcomp = "purchase" Then


                    txtvendorref.Text = (drSQL1.Item(27))
                    txtvendorcontact.Text = (drSQL1.Item(28))
                    txtvendquote.Text = (drSQL1.Item(29))
                    txtspecial.Text = (drSQL1.Item(30))

                End If

                EditCertDetails()

            Else

                MsgBox("Error occured", vbInformation)
                Exit Sub

            End If

        Else
            MsgBox("Click on Sno ", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub txtvendorref_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvendorref.TextChanged

    End Sub
    Private Sub EditCertDetails()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConfigurationManager.AppSettings("ConnectionString"))
        'Dim cnSQL As SqlConnection = New SqlConnection(System.Configuration!System.Configuration.configurationManager.AppSettings[ConnectionString as string()}




        'SqlConnection connection=new SqlConnection(ConfigurationManager.AppSettings["real"].ToString());




        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert As String
        'Dim b As Integer

        strsql = "Select Certificates from ENQ_EnqWise_Certificates where Enq_Reg_NO = " & txtregno.Text & " "

        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        i = CheckedListBoxCertificate.Items.Count

        Do While drSQL1.Read()
            cert = drSQL1.Item(0)
            a = 0
            Do While a < i

                If cert = CheckedListBoxCertificate.Items(a) Then

                    CheckedListBoxCertificate.SetItemChecked(a, True)

                    a = i
                Else

                    a = a + 1

                End If


            Loop

        Loop

    End Sub
    Private Sub listloadCertificate()


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConfigurationManager.AppSettings("ConnectionString"))
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        'Dim a As ListView


        cnSQL1.Open()
        strSQL1 = "SELECT Certificates,Int_code FROM ENQ_Certificates " & _
                 "WHERE  Status = 'A'"
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        Dim ColumnValue As String = Nothing
        Do While drSQL1.Read()

            ColumnValue = (drSQL1.GetValue(0)).ToString
            CheckedListBoxCertificate.Items.Add(ColumnValue)
            '  ListBoxCertificate.Sorted = True

            'ListBoxCertificate.DisplayMember = "Certificates"
            CheckedListBoxCertificate.ValueMember = "Int_code"

        Loop
    End Sub


End Class
