Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms

Public Class PartCreation

    Inherits System.Windows.Forms.Form

    Private ConnectionString As String
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents dtpRegDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtCustName As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtCustcity As System.Windows.Forms.TextBox
    Friend WithEvents txtCustID As System.Windows.Forms.TextBox
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents txtDocDetails As System.Windows.Forms.TextBox
    Friend WithEvents txtRegNo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonDomestic As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonExport As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents RBCustomerExisting As System.Windows.Forms.RadioButton
    Friend WithEvents RBCustomerNew As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonDocYES As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonDocNo As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DTPEnquDue As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents txtSpecial As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents RBTenderYes As System.Windows.Forms.RadioButton
    Friend WithEvents RBTenderNo As System.Windows.Forms.RadioButton
    Friend WithEvents txtDetailIntcode As System.Windows.Forms.TextBox
    Friend WithEvents txtTSSISeg As System.Windows.Forms.TextBox
    Friend WithEvents txtISR As System.Windows.Forms.TextBox
    Friend WithEvents txtCSR As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtTSSSeg As System.Windows.Forms.TextBox
    Friend WithEvents txtUOM As System.Windows.Forms.TextBox
    Friend WithEvents txtInventoryAc As System.Windows.Forms.TextBox
    Friend WithEvents txtProdLine As System.Windows.Forms.TextBox
    Friend WithEvents txtBuyer As System.Windows.Forms.TextBox
    Friend WithEvents txtPlanner As System.Windows.Forms.TextBox
    Friend WithEvents txtMB As System.Windows.Forms.TextBox
    Friend WithEvents txtPartSource As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtSpInsPurchaseDept As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Public stockDA As SqlDataAdapter = New SqlDataAdapter
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"


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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents datagridPartCreation As System.Windows.Forms.DataGrid
    Friend WithEvents DataUpdation As System.Windows.Forms.GroupBox
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents txtDimension As System.Windows.Forms.TextBox
    Friend WithEvents txtCustDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents txtPartNo As System.Windows.Forms.TextBox
    Friend WithEvents txtCustPart As System.Windows.Forms.TextBox
    Friend WithEvents txtPartDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtslno As System.Windows.Forms.TextBox
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents txtMaterial As System.Windows.Forms.TextBox
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents txtAplSpecialInst As System.Windows.Forms.TextBox
    Friend WithEvents txtSpeNote2 As System.Windows.Forms.TextBox
    Friend WithEvents txtSpNote1 As System.Windows.Forms.TextBox
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents txtChildDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtfix As System.Windows.Forms.TextBox
    Friend WithEvents txtInsp As System.Windows.Forms.TextBox
    Friend WithEvents txtrun As System.Windows.Forms.TextBox
    Friend WithEvents sp1 As System.Windows.Forms.Label
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents lblPlanner As System.Windows.Forms.Label
    Friend WithEvents lblChildItem As System.Windows.Forms.Label
    Friend WithEvents lblmb As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents GroupBoxSelect As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonRefresh As System.Windows.Forms.Button
    Friend WithEvents RadioButtonAll As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonPending As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtoncompleted As System.Windows.Forms.RadioButton
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtPartAppApl As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBoxSelect = New System.Windows.Forms.GroupBox()
        Me.ButtonRefresh = New System.Windows.Forms.Button()
        Me.RadioButtonAll = New System.Windows.Forms.RadioButton()
        Me.RadioButtonPending = New System.Windows.Forms.RadioButton()
        Me.RadioButtoncompleted = New System.Windows.Forms.RadioButton()
        Me.datagridPartCreation = New System.Windows.Forms.DataGrid()
        Me.DataUpdation = New System.Windows.Forms.GroupBox()
        Me.txtMB = New System.Windows.Forms.TextBox()
        Me.txtPartSource = New System.Windows.Forms.TextBox()
        Me.txtInventoryAc = New System.Windows.Forms.TextBox()
        Me.txtProdLine = New System.Windows.Forms.TextBox()
        Me.txtBuyer = New System.Windows.Forms.TextBox()
        Me.txtPlanner = New System.Windows.Forms.TextBox()
        Me.txtUOM = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtPartAppApl = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtAplSpecialInst = New System.Windows.Forms.TextBox()
        Me.txtSpeNote2 = New System.Windows.Forms.TextBox()
        Me.txtSpNote1 = New System.Windows.Forms.TextBox()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.txtChildDesc = New System.Windows.Forms.TextBox()
        Me.txtfix = New System.Windows.Forms.TextBox()
        Me.txtInsp = New System.Windows.Forms.TextBox()
        Me.txtrun = New System.Windows.Forms.TextBox()
        Me.sp1 = New System.Windows.Forms.Label()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.lblPlanner = New System.Windows.Forms.Label()
        Me.lblChildItem = New System.Windows.Forms.Label()
        Me.lblmb = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtDimension = New System.Windows.Forms.TextBox()
        Me.txtCustDesc = New System.Windows.Forms.TextBox()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.txtPartNo = New System.Windows.Forms.TextBox()
        Me.txtCustPart = New System.Windows.Forms.TextBox()
        Me.txtPartDesc = New System.Windows.Forms.TextBox()
        Me.txtslno = New System.Windows.Forms.TextBox()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.txtMaterial = New System.Windows.Forms.TextBox()
        Me.txtRemarks = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.dtpRegDt = New System.Windows.Forms.DateTimePicker()
        Me.txtCustName = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtCustcity = New System.Windows.Forms.TextBox()
        Me.txtCustID = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.txtDocDetails = New System.Windows.Forms.TextBox()
        Me.txtRegNo = New System.Windows.Forms.TextBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonDomestic = New System.Windows.Forms.RadioButton()
        Me.RadioButtonExport = New System.Windows.Forms.RadioButton()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.RBCustomerExisting = New System.Windows.Forms.RadioButton()
        Me.RBCustomerNew = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonDocYES = New System.Windows.Forms.RadioButton()
        Me.RadioButtonDocNo = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DTPEnquDue = New System.Windows.Forms.DateTimePicker()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.txtSpecial = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.RBTenderYes = New System.Windows.Forms.RadioButton()
        Me.RBTenderNo = New System.Windows.Forms.RadioButton()
        Me.txtDetailIntcode = New System.Windows.Forms.TextBox()
        Me.txtTSSISeg = New System.Windows.Forms.TextBox()
        Me.txtISR = New System.Windows.Forms.TextBox()
        Me.txtCSR = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtSpInsPurchaseDept = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtTSSSeg = New System.Windows.Forms.TextBox()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBoxSelect.SuspendLayout()
        CType(Me.datagridPartCreation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DataUpdation.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.GroupBoxSelect)
        Me.GroupBox2.Controls.Add(Me.datagridPartCreation)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox2.Location = New System.Drawing.Point(45, 10)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1425, 368)
        Me.GroupBox2.TabIndex = 95
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Part Number Creation Details"
        '
        'GroupBoxSelect
        '
        Me.GroupBoxSelect.Controls.Add(Me.ButtonRefresh)
        Me.GroupBoxSelect.Controls.Add(Me.RadioButtonAll)
        Me.GroupBoxSelect.Controls.Add(Me.RadioButtonPending)
        Me.GroupBoxSelect.Controls.Add(Me.RadioButtoncompleted)
        Me.GroupBoxSelect.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBoxSelect.ForeColor = System.Drawing.Color.Black
        Me.GroupBoxSelect.Location = New System.Drawing.Point(216, 5)
        Me.GroupBoxSelect.Name = "GroupBoxSelect"
        Me.GroupBoxSelect.Size = New System.Drawing.Size(515, 36)
        Me.GroupBoxSelect.TabIndex = 173
        Me.GroupBoxSelect.TabStop = False
        Me.GroupBoxSelect.Text = "Status"
        '
        'ButtonRefresh
        '
        Me.ButtonRefresh.BackColor = System.Drawing.Color.DodgerBlue
        Me.ButtonRefresh.ForeColor = System.Drawing.Color.White
        Me.ButtonRefresh.Location = New System.Drawing.Point(390, 9)
        Me.ButtonRefresh.Name = "ButtonRefresh"
        Me.ButtonRefresh.Size = New System.Drawing.Size(87, 27)
        Me.ButtonRefresh.TabIndex = 116
        Me.ButtonRefresh.Text = "Refresh"
        Me.ButtonRefresh.UseVisualStyleBackColor = False
        '
        'RadioButtonAll
        '
        Me.RadioButtonAll.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonAll.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonAll.Location = New System.Drawing.Point(325, 12)
        Me.RadioButtonAll.Name = "RadioButtonAll"
        Me.RadioButtonAll.Size = New System.Drawing.Size(75, 22)
        Me.RadioButtonAll.TabIndex = 115
        Me.RadioButtonAll.Text = "All"
        '
        'RadioButtonPending
        '
        Me.RadioButtonPending.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonPending.ForeColor = System.Drawing.Color.Red
        Me.RadioButtonPending.Location = New System.Drawing.Point(92, 11)
        Me.RadioButtonPending.Name = "RadioButtonPending"
        Me.RadioButtonPending.Size = New System.Drawing.Size(84, 22)
        Me.RadioButtonPending.TabIndex = 113
        Me.RadioButtonPending.Text = "Pending"
        '
        'RadioButtoncompleted
        '
        Me.RadioButtoncompleted.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtoncompleted.ForeColor = System.Drawing.Color.Black
        Me.RadioButtoncompleted.Location = New System.Drawing.Point(188, 11)
        Me.RadioButtoncompleted.Name = "RadioButtoncompleted"
        Me.RadioButtoncompleted.Size = New System.Drawing.Size(93, 22)
        Me.RadioButtoncompleted.TabIndex = 114
        Me.RadioButtoncompleted.Text = "Completed"
        '
        'datagridPartCreation
        '
        Me.datagridPartCreation.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.datagridPartCreation.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridPartCreation.CaptionVisible = False
        Me.datagridPartCreation.DataMember = ""
        Me.datagridPartCreation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridPartCreation.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridPartCreation.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.datagridPartCreation.Location = New System.Drawing.Point(19, 45)
        Me.datagridPartCreation.Name = "datagridPartCreation"
        Me.datagridPartCreation.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.datagridPartCreation.ParentRowsVisible = False
        Me.datagridPartCreation.PreferredColumnWidth = 85
        Me.datagridPartCreation.ReadOnly = True
        Me.datagridPartCreation.RowHeadersVisible = False
        Me.datagridPartCreation.Size = New System.Drawing.Size(1375, 312)
        Me.datagridPartCreation.TabIndex = 0
        '
        'DataUpdation
        '
        Me.DataUpdation.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataUpdation.Controls.Add(Me.txtMB)
        Me.DataUpdation.Controls.Add(Me.txtPartSource)
        Me.DataUpdation.Controls.Add(Me.txtInventoryAc)
        Me.DataUpdation.Controls.Add(Me.txtProdLine)
        Me.DataUpdation.Controls.Add(Me.txtBuyer)
        Me.DataUpdation.Controls.Add(Me.txtPlanner)
        Me.DataUpdation.Controls.Add(Me.txtUOM)
        Me.DataUpdation.Controls.Add(Me.Label9)
        Me.DataUpdation.Controls.Add(Me.txtPartAppApl)
        Me.DataUpdation.Controls.Add(Me.Label8)
        Me.DataUpdation.Controls.Add(Me.Label6)
        Me.DataUpdation.Controls.Add(Me.Label4)
        Me.DataUpdation.Controls.Add(Me.txtAplSpecialInst)
        Me.DataUpdation.Controls.Add(Me.txtSpeNote2)
        Me.DataUpdation.Controls.Add(Me.txtSpNote1)
        Me.DataUpdation.Controls.Add(Me.Label65)
        Me.DataUpdation.Controls.Add(Me.txtChildDesc)
        Me.DataUpdation.Controls.Add(Me.txtfix)
        Me.DataUpdation.Controls.Add(Me.txtInsp)
        Me.DataUpdation.Controls.Add(Me.txtrun)
        Me.DataUpdation.Controls.Add(Me.sp1)
        Me.DataUpdation.Controls.Add(Me.Label59)
        Me.DataUpdation.Controls.Add(Me.Label5)
        Me.DataUpdation.Controls.Add(Me.Label60)
        Me.DataUpdation.Controls.Add(Me.lblPlanner)
        Me.DataUpdation.Controls.Add(Me.lblChildItem)
        Me.DataUpdation.Controls.Add(Me.lblmb)
        Me.DataUpdation.Controls.Add(Me.Label43)
        Me.DataUpdation.Controls.Add(Me.Label42)
        Me.DataUpdation.Controls.Add(Me.txtDimension)
        Me.DataUpdation.Controls.Add(Me.txtCustDesc)
        Me.DataUpdation.Controls.Add(Me.Label41)
        Me.DataUpdation.Controls.Add(Me.Label40)
        Me.DataUpdation.Controls.Add(Me.Label39)
        Me.DataUpdation.Controls.Add(Me.Label38)
        Me.DataUpdation.Controls.Add(Me.Label37)
        Me.DataUpdation.Controls.Add(Me.Label36)
        Me.DataUpdation.Controls.Add(Me.txtPartNo)
        Me.DataUpdation.Controls.Add(Me.txtCustPart)
        Me.DataUpdation.Controls.Add(Me.txtPartDesc)
        Me.DataUpdation.Controls.Add(Me.txtslno)
        Me.DataUpdation.Controls.Add(Me.btnUpdate)
        Me.DataUpdation.Controls.Add(Me.txtMaterial)
        Me.DataUpdation.Controls.Add(Me.txtRemarks)
        Me.DataUpdation.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataUpdation.ForeColor = System.Drawing.Color.Firebrick
        Me.DataUpdation.Location = New System.Drawing.Point(45, 567)
        Me.DataUpdation.Name = "DataUpdation"
        Me.DataUpdation.Size = New System.Drawing.Size(1425, 287)
        Me.DataUpdation.TabIndex = 98
        Me.DataUpdation.TabStop = False
        Me.DataUpdation.Text = "Details"
        '
        'txtMB
        '
        Me.txtMB.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMB.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMB.ForeColor = System.Drawing.Color.Black
        Me.txtMB.Location = New System.Drawing.Point(163, 38)
        Me.txtMB.Name = "txtMB"
        Me.txtMB.ReadOnly = True
        Me.txtMB.Size = New System.Drawing.Size(62, 23)
        Me.txtMB.TabIndex = 207
        '
        'txtPartSource
        '
        Me.txtPartSource.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartSource.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartSource.ForeColor = System.Drawing.Color.Black
        Me.txtPartSource.Location = New System.Drawing.Point(63, 36)
        Me.txtPartSource.Name = "txtPartSource"
        Me.txtPartSource.ReadOnly = True
        Me.txtPartSource.Size = New System.Drawing.Size(93, 23)
        Me.txtPartSource.TabIndex = 206
        '
        'txtInventoryAc
        '
        Me.txtInventoryAc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtInventoryAc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInventoryAc.Location = New System.Drawing.Point(516, 114)
        Me.txtInventoryAc.Name = "txtInventoryAc"
        Me.txtInventoryAc.Size = New System.Drawing.Size(166, 23)
        Me.txtInventoryAc.TabIndex = 205
        '
        'txtProdLine
        '
        Me.txtProdLine.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtProdLine.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProdLine.Location = New System.Drawing.Point(322, 114)
        Me.txtProdLine.Name = "txtProdLine"
        Me.txtProdLine.Size = New System.Drawing.Size(187, 23)
        Me.txtProdLine.TabIndex = 204
        '
        'txtBuyer
        '
        Me.txtBuyer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBuyer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBuyer.Location = New System.Drawing.Point(84, 114)
        Me.txtBuyer.Name = "txtBuyer"
        Me.txtBuyer.Size = New System.Drawing.Size(69, 23)
        Me.txtBuyer.TabIndex = 203
        '
        'txtPlanner
        '
        Me.txtPlanner.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPlanner.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPlanner.Location = New System.Drawing.Point(10, 114)
        Me.txtPlanner.Name = "txtPlanner"
        Me.txtPlanner.Size = New System.Drawing.Size(70, 23)
        Me.txtPlanner.TabIndex = 202
        '
        'txtUOM
        '
        Me.txtUOM.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUOM.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUOM.Location = New System.Drawing.Point(1346, 34)
        Me.txtUOM.Name = "txtUOM"
        Me.txtUOM.Size = New System.Drawing.Size(71, 23)
        Me.txtUOM.TabIndex = 201
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(308, 196)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(257, 18)
        Me.Label9.TabIndex = 200
        Me.Label9.Text = "Special instructions to Customer Sup."
        '
        'txtPartAppApl
        '
        Me.txtPartAppApl.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtPartAppApl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartAppApl.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartAppApl.Location = New System.Drawing.Point(308, 168)
        Me.txtPartAppApl.Name = "txtPartAppApl"
        Me.txtPartAppApl.Size = New System.Drawing.Size(371, 23)
        Me.txtPartAppApl.TabIndex = 199
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(78, 94)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 16)
        Me.Label8.TabIndex = 198
        Me.Label8.Text = "Buyer"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Blue
        Me.Label6.Location = New System.Drawing.Point(1357, 15)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(76, 19)
        Me.Label6.TabIndex = 197
        Me.Label6.Text = "Uom"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(63, 15)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(84, 19)
        Me.Label4.TabIndex = 196
        Me.Label4.Text = "Part Source"
        '
        'txtAplSpecialInst
        '
        Me.txtAplSpecialInst.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtAplSpecialInst.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAplSpecialInst.Location = New System.Drawing.Point(304, 217)
        Me.txtAplSpecialInst.MaxLength = 200
        Me.txtAplSpecialInst.Multiline = True
        Me.txtAplSpecialInst.Name = "txtAplSpecialInst"
        Me.txtAplSpecialInst.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtAplSpecialInst.Size = New System.Drawing.Size(375, 51)
        Me.txtAplSpecialInst.TabIndex = 190
        '
        'txtSpeNote2
        '
        Me.txtSpeNote2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSpeNote2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpeNote2.Location = New System.Drawing.Point(689, 177)
        Me.txtSpeNote2.Multiline = True
        Me.txtSpeNote2.Name = "txtSpeNote2"
        Me.txtSpeNote2.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSpeNote2.Size = New System.Drawing.Size(714, 68)
        Me.txtSpeNote2.TabIndex = 189
        '
        'txtSpNote1
        '
        Me.txtSpNote1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSpNote1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpNote1.Location = New System.Drawing.Point(689, 114)
        Me.txtSpNote1.Multiline = True
        Me.txtSpNote1.Name = "txtSpNote1"
        Me.txtSpNote1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSpNote1.Size = New System.Drawing.Size(714, 56)
        Me.txtSpNote1.TabIndex = 188
        '
        'Label65
        '
        Me.Label65.AutoSize = True
        Me.Label65.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label65.ForeColor = System.Drawing.Color.Black
        Me.Label65.Location = New System.Drawing.Point(308, 143)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(211, 18)
        Me.Label65.TabIndex = 187
        Me.Label65.Text = "PartNo. Approved by Apl. Dept."
        '
        'txtChildDesc
        '
        Me.txtChildDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChildDesc.Location = New System.Drawing.Point(10, 165)
        Me.txtChildDesc.Multiline = True
        Me.txtChildDesc.Name = "txtChildDesc"
        Me.txtChildDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtChildDesc.Size = New System.Drawing.Size(282, 108)
        Me.txtChildDesc.TabIndex = 185
        '
        'txtfix
        '
        Me.txtfix.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtfix.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtfix.Location = New System.Drawing.Point(216, 114)
        Me.txtfix.Name = "txtfix"
        Me.txtfix.Size = New System.Drawing.Size(45, 23)
        Me.txtfix.TabIndex = 184
        '
        'txtInsp
        '
        Me.txtInsp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtInsp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInsp.Location = New System.Drawing.Point(268, 114)
        Me.txtInsp.Name = "txtInsp"
        Me.txtInsp.Size = New System.Drawing.Size(47, 23)
        Me.txtInsp.TabIndex = 183
        '
        'txtrun
        '
        Me.txtrun.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtrun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtrun.Location = New System.Drawing.Point(163, 114)
        Me.txtrun.Name = "txtrun"
        Me.txtrun.Size = New System.Drawing.Size(46, 23)
        Me.txtrun.TabIndex = 182
        '
        'sp1
        '
        Me.sp1.AutoSize = True
        Me.sp1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sp1.ForeColor = System.Drawing.Color.Black
        Me.sp1.Location = New System.Drawing.Point(686, 94)
        Me.sp1.Name = "sp1"
        Me.sp1.Size = New System.Drawing.Size(168, 16)
        Me.sp1.TabIndex = 177
        Me.sp1.Text = "Sp. Note 1 and Sp. Note2"
        '
        'Label59
        '
        Me.Label59.AutoSize = True
        Me.Label59.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label59.ForeColor = System.Drawing.Color.Black
        Me.Label59.Location = New System.Drawing.Point(512, 94)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(93, 16)
        Me.Label59.TabIndex = 176
        Me.Label59.Text = "Inventory  A/c"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(331, 94)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(69, 16)
        Me.Label5.TabIndex = 175
        Me.Label5.Text = "Prod LIne"
        '
        'Label60
        '
        Me.Label60.AutoSize = True
        Me.Label60.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label60.ForeColor = System.Drawing.Color.Black
        Me.Label60.Location = New System.Drawing.Point(160, 94)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(164, 16)
        Me.Label60.TabIndex = 174
        Me.Label60.Text = "Lead Time [Run Fix Insp]"
        '
        'lblPlanner
        '
        Me.lblPlanner.AutoSize = True
        Me.lblPlanner.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlanner.ForeColor = System.Drawing.Color.Black
        Me.lblPlanner.Location = New System.Drawing.Point(9, 94)
        Me.lblPlanner.Name = "lblPlanner"
        Me.lblPlanner.Size = New System.Drawing.Size(57, 16)
        Me.lblPlanner.TabIndex = 172
        Me.lblPlanner.Text = "Planner"
        '
        'lblChildItem
        '
        Me.lblChildItem.AutoSize = True
        Me.lblChildItem.Location = New System.Drawing.Point(7, 143)
        Me.lblChildItem.Name = "lblChildItem"
        Me.lblChildItem.Size = New System.Drawing.Size(119, 18)
        Me.lblChildItem.TabIndex = 171
        Me.lblChildItem.Text = "Child Item Desc"
        '
        'lblmb
        '
        Me.lblmb.AutoSize = True
        Me.lblmb.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblmb.ForeColor = System.Drawing.Color.Black
        Me.lblmb.Location = New System.Drawing.Point(160, 17)
        Me.lblmb.Name = "lblmb"
        Me.lblmb.Size = New System.Drawing.Size(32, 16)
        Me.lblmb.TabIndex = 170
        Me.lblmb.Text = "M/B"
        '
        'Label43
        '
        Me.Label43.BackColor = System.Drawing.Color.Transparent
        Me.Label43.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label43.ForeColor = System.Drawing.Color.Black
        Me.Label43.Location = New System.Drawing.Point(654, 68)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(60, 15)
        Me.Label43.TabIndex = 138
        Me.Label43.Text = "Remarks"
        '
        'Label42
        '
        Me.Label42.BackColor = System.Drawing.Color.Transparent
        Me.Label42.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label42.ForeColor = System.Drawing.Color.Black
        Me.Label42.Location = New System.Drawing.Point(341, 68)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(56, 15)
        Me.Label42.TabIndex = 136
        Me.Label42.Text = "Material"
        '
        'txtDimension
        '
        Me.txtDimension.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDimension.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDimension.Location = New System.Drawing.Point(82, 66)
        Me.txtDimension.Name = "txtDimension"
        Me.txtDimension.Size = New System.Drawing.Size(252, 23)
        Me.txtDimension.TabIndex = 135
        '
        'txtCustDesc
        '
        Me.txtCustDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustDesc.Location = New System.Drawing.Point(1063, 38)
        Me.txtCustDesc.MaxLength = 20
        Me.txtCustDesc.Name = "txtCustDesc"
        Me.txtCustDesc.Size = New System.Drawing.Size(276, 23)
        Me.txtCustDesc.TabIndex = 134
        '
        'Label41
        '
        Me.Label41.BackColor = System.Drawing.Color.Transparent
        Me.Label41.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label41.ForeColor = System.Drawing.Color.Black
        Me.Label41.Location = New System.Drawing.Point(9, 17)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(47, 19)
        Me.Label41.TabIndex = 132
        Me.Label41.Text = "Sl.No."
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.Color.Transparent
        Me.Label40.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.ForeColor = System.Drawing.Color.Black
        Me.Label40.Location = New System.Drawing.Point(9, 66)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(94, 19)
        Me.Label40.TabIndex = 131
        Me.Label40.Text = "Dimension"
        '
        'Label39
        '
        Me.Label39.BackColor = System.Drawing.Color.Transparent
        Me.Label39.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.ForeColor = System.Drawing.Color.Black
        Me.Label39.Location = New System.Drawing.Point(792, 18)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(187, 20)
        Me.Label39.TabIndex = 130
        Me.Label39.Text = "Customer Part No."
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.Color.Transparent
        Me.Label38.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.ForeColor = System.Drawing.Color.Black
        Me.Label38.Location = New System.Drawing.Point(1063, 18)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(177, 20)
        Me.Label38.TabIndex = 129
        Me.Label38.Text = "Cust Part Description"
        '
        'Label37
        '
        Me.Label37.BackColor = System.Drawing.Color.Transparent
        Me.Label37.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.ForeColor = System.Drawing.Color.Black
        Me.Label37.Location = New System.Drawing.Point(512, 18)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(187, 20)
        Me.Label37.TabIndex = 128
        Me.Label37.Text = "Description"
        '
        'Label36
        '
        Me.Label36.BackColor = System.Drawing.Color.Transparent
        Me.Label36.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.ForeColor = System.Drawing.Color.Blue
        Me.Label36.Location = New System.Drawing.Point(233, 13)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(243, 20)
        Me.Label36.TabIndex = 127
        Me.Label36.Text = "Part No."
        '
        'txtPartNo
        '
        Me.txtPartNo.BackColor = System.Drawing.Color.White
        Me.txtPartNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartNo.Location = New System.Drawing.Point(232, 36)
        Me.txtPartNo.Name = "txtPartNo"
        Me.txtPartNo.Size = New System.Drawing.Size(229, 23)
        Me.txtPartNo.TabIndex = 119
        '
        'txtCustPart
        '
        Me.txtCustPart.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustPart.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustPart.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustPart.Location = New System.Drawing.Point(792, 38)
        Me.txtCustPart.Name = "txtCustPart"
        Me.txtCustPart.Size = New System.Drawing.Size(261, 23)
        Me.txtCustPart.TabIndex = 121
        '
        'txtPartDesc
        '
        Me.txtPartDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartDesc.Location = New System.Drawing.Point(468, 38)
        Me.txtPartDesc.Name = "txtPartDesc"
        Me.txtPartDesc.Size = New System.Drawing.Size(309, 23)
        Me.txtPartDesc.TabIndex = 120
        '
        'txtslno
        '
        Me.txtslno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtslno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtslno.ForeColor = System.Drawing.Color.Black
        Me.txtslno.Location = New System.Drawing.Point(9, 36)
        Me.txtslno.Name = "txtslno"
        Me.txtslno.ReadOnly = True
        Me.txtslno.Size = New System.Drawing.Size(38, 23)
        Me.txtslno.TabIndex = 118
        '
        'btnUpdate
        '
        Me.btnUpdate.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.btnUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.ForeColor = System.Drawing.Color.Black
        Me.btnUpdate.Location = New System.Drawing.Point(1338, 253)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(65, 29)
        Me.btnUpdate.TabIndex = 124
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = False
        '
        'txtMaterial
        '
        Me.txtMaterial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMaterial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterial.Location = New System.Drawing.Point(404, 64)
        Me.txtMaterial.MaxLength = 20
        Me.txtMaterial.Name = "txtMaterial"
        Me.txtMaterial.Size = New System.Drawing.Size(243, 23)
        Me.txtMaterial.TabIndex = 122
        '
        'txtRemarks
        '
        Me.txtRemarks.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRemarks.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRemarks.Location = New System.Drawing.Point(724, 66)
        Me.txtRemarks.MaxLength = 12
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(693, 23)
        Me.txtRemarks.TabIndex = 123
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(9, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(87, 28)
        Me.Label2.TabIndex = 100
        Me.Label2.Text = "Enq.Reg. No."
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(212, 22)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(28, 29)
        Me.Label7.TabIndex = 101
        Me.Label7.Text = "Dt."
        '
        'dtpRegDt
        '
        Me.dtpRegDt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpRegDt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpRegDt.Location = New System.Drawing.Point(243, 22)
        Me.dtpRegDt.Name = "dtpRegDt"
        Me.dtpRegDt.Size = New System.Drawing.Size(149, 23)
        Me.dtpRegDt.TabIndex = 108
        '
        'txtCustName
        '
        Me.txtCustName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustName.Location = New System.Drawing.Point(261, 64)
        Me.txtCustName.Name = "txtCustName"
        Me.txtCustName.Size = New System.Drawing.Size(453, 23)
        Me.txtCustName.TabIndex = 75
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(212, 63)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(49, 29)
        Me.Label10.TabIndex = 76
        Me.Label10.Text = "Name"
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(721, 67)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(47, 24)
        Me.Label15.TabIndex = 84
        Me.Label15.Text = "City"
        '
        'txtCustcity
        '
        Me.txtCustcity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustcity.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustcity.Location = New System.Drawing.Point(763, 64)
        Me.txtCustcity.Name = "txtCustcity"
        Me.txtCustcity.Size = New System.Drawing.Size(370, 23)
        Me.txtCustcity.TabIndex = 85
        '
        'txtCustID
        '
        Me.txtCustID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustID.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustID.Location = New System.Drawing.Point(103, 58)
        Me.txtCustID.Name = "txtCustID"
        Me.txtCustID.Size = New System.Drawing.Size(102, 26)
        Me.txtCustID.TabIndex = 115
        '
        'Label34
        '
        Me.Label34.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.ForeColor = System.Drawing.Color.Black
        Me.Label34.Location = New System.Drawing.Point(1281, 35)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(122, 28)
        Me.Label34.TabIndex = 133
        Me.Label34.Text = "Document Details"
        '
        'txtDocDetails
        '
        Me.txtDocDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDocDetails.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDocDetails.Location = New System.Drawing.Point(1136, 69)
        Me.txtDocDetails.Multiline = True
        Me.txtDocDetails.Name = "txtDocDetails"
        Me.txtDocDetails.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDocDetails.Size = New System.Drawing.Size(267, 94)
        Me.txtDocDetails.TabIndex = 134
        '
        'txtRegNo
        '
        Me.txtRegNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRegNo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegNo.Location = New System.Drawing.Point(103, 24)
        Me.txtRegNo.Name = "txtRegNo"
        Me.txtRegNo.Size = New System.Drawing.Size(102, 26)
        Me.txtRegNo.TabIndex = 140
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.RadioButtonDomestic)
        Me.GroupBox6.Controls.Add(Me.RadioButtonExport)
        Me.GroupBox6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox6.ForeColor = System.Drawing.Color.Black
        Me.GroupBox6.Location = New System.Drawing.Point(567, 18)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(182, 43)
        Me.GroupBox6.TabIndex = 169
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Market Type"
        '
        'RadioButtonDomestic
        '
        Me.RadioButtonDomestic.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonDomestic.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonDomestic.Location = New System.Drawing.Point(7, 15)
        Me.RadioButtonDomestic.Name = "RadioButtonDomestic"
        Me.RadioButtonDomestic.Size = New System.Drawing.Size(84, 24)
        Me.RadioButtonDomestic.TabIndex = 52
        Me.RadioButtonDomestic.Text = "Domestic"
        '
        'RadioButtonExport
        '
        Me.RadioButtonExport.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonExport.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonExport.Location = New System.Drawing.Point(92, 13)
        Me.RadioButtonExport.Name = "RadioButtonExport"
        Me.RadioButtonExport.Size = New System.Drawing.Size(67, 27)
        Me.RadioButtonExport.TabIndex = 53
        Me.RadioButtonExport.Text = "Export"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.RBCustomerExisting)
        Me.GroupBox7.Controls.Add(Me.RBCustomerNew)
        Me.GroupBox7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox7.ForeColor = System.Drawing.Color.Black
        Me.GroupBox7.Location = New System.Drawing.Point(404, 18)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(156, 43)
        Me.GroupBox7.TabIndex = 170
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Customer"
        '
        'RBCustomerExisting
        '
        Me.RBCustomerExisting.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBCustomerExisting.ForeColor = System.Drawing.Color.Black
        Me.RBCustomerExisting.Location = New System.Drawing.Point(69, 15)
        Me.RBCustomerExisting.Name = "RBCustomerExisting"
        Me.RBCustomerExisting.Size = New System.Drawing.Size(72, 20)
        Me.RBCustomerExisting.TabIndex = 13
        Me.RBCustomerExisting.Text = "Existing"
        '
        'RBCustomerNew
        '
        Me.RBCustomerNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBCustomerNew.ForeColor = System.Drawing.Color.Black
        Me.RBCustomerNew.Location = New System.Drawing.Point(7, 16)
        Me.RBCustomerNew.Name = "RBCustomerNew"
        Me.RBCustomerNew.Size = New System.Drawing.Size(65, 19)
        Me.RBCustomerNew.TabIndex = 12
        Me.RBCustomerNew.Text = "New"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(9, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 29)
        Me.Label1.TabIndex = 171
        Me.Label1.Text = "Cust Id"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButtonDocYES)
        Me.GroupBox1.Controls.Add(Me.RadioButtonDocNo)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(1136, 18)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(138, 46)
        Me.GroupBox1.TabIndex = 172
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Document Uploaded"
        '
        'RadioButtonDocYES
        '
        Me.RadioButtonDocYES.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonDocYES.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonDocYES.Location = New System.Drawing.Point(12, 16)
        Me.RadioButtonDocYES.Name = "RadioButtonDocYES"
        Me.RadioButtonDocYES.Size = New System.Drawing.Size(61, 29)
        Me.RadioButtonDocYES.TabIndex = 113
        Me.RadioButtonDocYES.Text = "Yes"
        '
        'RadioButtonDocNo
        '
        Me.RadioButtonDocNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonDocNo.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonDocNo.Location = New System.Drawing.Point(80, 15)
        Me.RadioButtonDocNo.Name = "RadioButtonDocNo"
        Me.RadioButtonDocNo.Size = New System.Drawing.Size(52, 29)
        Me.RadioButtonDocNo.TabIndex = 114
        Me.RadioButtonDocNo.Text = "No"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(884, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(95, 29)
        Me.Label3.TabIndex = 173
        Me.Label3.Text = "Enquiry Due on"
        '
        'DTPEnquDue
        '
        Me.DTPEnquDue.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTPEnquDue.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTPEnquDue.Location = New System.Drawing.Point(982, 30)
        Me.DTPEnquDue.Name = "DTPEnquDue"
        Me.DTPEnquDue.Size = New System.Drawing.Size(147, 23)
        Me.DTPEnquDue.TabIndex = 174
        '
        'Label52
        '
        Me.Label52.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label52.ForeColor = System.Drawing.Color.Black
        Me.Label52.Location = New System.Drawing.Point(349, 92)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(176, 20)
        Me.Label52.TabIndex = 176
        Me.Label52.Text = "Special Inst for pur/App.Dept."
        '
        'txtSpecial
        '
        Me.txtSpecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSpecial.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpecial.Location = New System.Drawing.Point(344, 109)
        Me.txtSpecial.MaxLength = 500
        Me.txtSpecial.Multiline = True
        Me.txtSpecial.Name = "txtSpecial"
        Me.txtSpecial.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSpecial.Size = New System.Drawing.Size(370, 66)
        Me.txtSpecial.TabIndex = 175
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.RBTenderYes)
        Me.GroupBox4.Controls.Add(Me.RBTenderNo)
        Me.GroupBox4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.Color.Black
        Me.GroupBox4.Location = New System.Drawing.Point(756, 22)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(127, 39)
        Me.GroupBox4.TabIndex = 180
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Tender"
        '
        'RBTenderYes
        '
        Me.RBTenderYes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBTenderYes.ForeColor = System.Drawing.Color.Red
        Me.RBTenderYes.Location = New System.Drawing.Point(7, 16)
        Me.RBTenderYes.Name = "RBTenderYes"
        Me.RBTenderYes.Size = New System.Drawing.Size(56, 19)
        Me.RBTenderYes.TabIndex = 177
        Me.RBTenderYes.Text = "Yes"
        '
        'RBTenderNo
        '
        Me.RBTenderNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBTenderNo.ForeColor = System.Drawing.Color.Red
        Me.RBTenderNo.Location = New System.Drawing.Point(70, 13)
        Me.RBTenderNo.Name = "RBTenderNo"
        Me.RBTenderNo.Size = New System.Drawing.Size(51, 20)
        Me.RBTenderNo.TabIndex = 178
        Me.RBTenderNo.Text = "No"
        '
        'txtDetailIntcode
        '
        Me.txtDetailIntcode.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtDetailIntcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDetailIntcode.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDetailIntcode.Location = New System.Drawing.Point(654, 0)
        Me.txtDetailIntcode.Name = "txtDetailIntcode"
        Me.txtDetailIntcode.Size = New System.Drawing.Size(269, 26)
        Me.txtDetailIntcode.TabIndex = 181
        '
        'txtTSSISeg
        '
        Me.txtTSSISeg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTSSISeg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTSSISeg.Location = New System.Drawing.Point(154, 118)
        Me.txtTSSISeg.Name = "txtTSSISeg"
        Me.txtTSSISeg.Size = New System.Drawing.Size(73, 23)
        Me.txtTSSISeg.TabIndex = 182
        '
        'txtISR
        '
        Me.txtISR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtISR.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtISR.Location = New System.Drawing.Point(87, 118)
        Me.txtISR.Name = "txtISR"
        Me.txtISR.Size = New System.Drawing.Size(60, 23)
        Me.txtISR.TabIndex = 183
        '
        'txtCSR
        '
        Me.txtCSR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSR.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCSR.Location = New System.Drawing.Point(13, 118)
        Me.txtCSR.Name = "txtCSR"
        Me.txtCSR.Size = New System.Drawing.Size(67, 23)
        Me.txtCSR.TabIndex = 184
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox3.Controls.Add(Me.txtSpInsPurchaseDept)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.Label11)
        Me.GroupBox3.Controls.Add(Me.txtTSSSeg)
        Me.GroupBox3.Controls.Add(Me.txtCSR)
        Me.GroupBox3.Controls.Add(Me.txtISR)
        Me.GroupBox3.Controls.Add(Me.txtTSSISeg)
        Me.GroupBox3.Controls.Add(Me.txtDetailIntcode)
        Me.GroupBox3.Controls.Add(Me.GroupBox4)
        Me.GroupBox3.Controls.Add(Me.txtSpecial)
        Me.GroupBox3.Controls.Add(Me.Label52)
        Me.GroupBox3.Controls.Add(Me.DTPEnquDue)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.GroupBox1)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.GroupBox7)
        Me.GroupBox3.Controls.Add(Me.GroupBox6)
        Me.GroupBox3.Controls.Add(Me.txtRegNo)
        Me.GroupBox3.Controls.Add(Me.txtDocDetails)
        Me.GroupBox3.Controls.Add(Me.Label34)
        Me.GroupBox3.Controls.Add(Me.txtCustID)
        Me.GroupBox3.Controls.Add(Me.txtCustcity)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.txtCustName)
        Me.GroupBox3.Controls.Add(Me.dtpRegDt)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.Black
        Me.GroupBox3.Location = New System.Drawing.Point(45, 385)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1425, 175)
        Me.GroupBox3.TabIndex = 106
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Customer Details"
        '
        'txtSpInsPurchaseDept
        '
        Me.txtSpInsPurchaseDept.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSpInsPurchaseDept.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpInsPurchaseDept.Location = New System.Drawing.Point(724, 109)
        Me.txtSpInsPurchaseDept.MaxLength = 500
        Me.txtSpInsPurchaseDept.Multiline = True
        Me.txtSpInsPurchaseDept.Name = "txtSpInsPurchaseDept"
        Me.txtSpInsPurchaseDept.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSpInsPurchaseDept.Size = New System.Drawing.Size(405, 66)
        Me.txtSpInsPurchaseDept.TabIndex = 190
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(721, 92)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(202, 20)
        Me.Label12.TabIndex = 189
        Me.Label12.Text = "Special Inst by Purchase Dept."
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(15, 95)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(327, 19)
        Me.Label11.TabIndex = 188
        Me.Label11.Text = "CSR                ISR            TSSISeg          TSSSeg"
        '
        'txtTSSSeg
        '
        Me.txtTSSSeg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTSSSeg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTSSSeg.Location = New System.Drawing.Point(237, 118)
        Me.txtTSSSeg.Name = "txtTSSSeg"
        Me.txtTSSSeg.Size = New System.Drawing.Size(81, 23)
        Me.txtTSSSeg.TabIndex = 187
        '
        'PartCreation
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 17)
        Me.BackColor = System.Drawing.Color.CornflowerBlue
        Me.ClientSize = New System.Drawing.Size(1291, 791)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.DataUpdation)
        Me.Controls.Add(Me.GroupBox2)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.Name = "PartCreation"
        Me.Text = "Part Creation"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBoxSelect.ResumeLayout(False)
        CType(Me.datagridPartCreation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DataUpdation.ResumeLayout(False)
        Me.DataUpdation.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub datagridPartCreation_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridPartCreation.CurrentCellChanged


        Dim b As Integer
        b = datagridPartCreation.CurrentCell.ColumnNumber()

        If b = 0 Then

            txtPartAppApl.Text = ""
            txtAplSpecialInst.Text = ""
            'clearcustomerdata()
            'ClearPartDatat()

            txtDetailIntcode.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 0)

            txtRegNo.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 1)
            txtslno.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 2)
            dtpRegDt.Value = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 3)
            txtCustID.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 4)
            txtCustName.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 5)
            txtCustcity.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 6)
            'part source  7
            txtPartNo.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 8)
            txtPartDesc.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 9)
            txtCustPart.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 10)
            txtCustDesc.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 11)
            txtUOM.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 12)
            txtDimension.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 13)
            txtMaterial.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 14)
            txtSpecial.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 15)
            If datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 16) = "01-01-1900" Then
                DTPEnquDue.Checked = False
                DTPEnquDue.Value = "01-01-1900"
            Else
                DTPEnquDue.Value = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 16)
            End If
            txtCSR.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 18)
            txtISR.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 19)
            txtTSSISeg.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 20)
            txtTSSSeg.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 21)

            If Trim(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 22)) = "YES" Then
                RBCustomerExisting.Checked = True
                RBCustomerNew.Checked = False
            ElseIf Trim(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 22)) = "NO" Then
                RBCustomerExisting.Checked = False
                RBCustomerNew.Checked = True
            End If

            If Trim(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 23)) = "YES" Then
                RadioButtonDocYES.Checked = True
                RadioButtonDocNo.Checked = False
            ElseIf Trim(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 23)) = "NO" Then
                RadioButtonDocNo.Checked = True
                RadioButtonDocYES.Checked = False
            End If

            txtDocDetails.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 24)

            txtRemarks.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 25)

            txtPartSource.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 26)

            If Trim(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 27)) = "YES" Then
                RBTenderYes.Checked = True
                RBTenderNo.Checked = False
            ElseIf Trim(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 27)) = "NO" Then
                RBTenderYes.Checked = False
                RBTenderNo.Checked = True
            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 28)) Then
                RadioButtonDomestic.Checked = True
                RadioButtonExport.Checked = False

            ElseIf Trim(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 28)) = "Domestic" Then

                RadioButtonDomestic.Checked = True
                RadioButtonExport.Checked = False
            ElseIf Trim(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 28)) = "Export" Then
                RadioButtonDomestic.Checked = False
                RadioButtonExport.Checked = True
            End If

            '            strSQL = "SELECT  Enq_Detail_code, RegNo, SlNo, [Reg.Date], CustomerID, CustomerName,City, Part_Source, PartNumber, PartDescription,10
            'CustPartNumber, CustPartDescription, uom,  Dimension, Material, Special, Enq_Due_date,Req, CSR, ISR,20
            'TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details,Special_instructions,  FS_Yes_NO, Tender_YesNo, MarketType,Item_Type, 30
            '           Planner, Buyer, Lead_Run, Lead_Fix, Lead_Insp, Child_desc, prod_Line, Inv_Ac, Sp_note1, sp_note2 " & _40
            '         " FROM ENQ_Parts_Created_Q where Part_No_Appl_Sug = '-' order by  RegNo, SlNo "
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 29)) Then
                txtMB.Text = "-"
            Else
                txtMB.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 29)

            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 30)) Then
                txtPlanner.Text = "-"
            Else
                txtPlanner.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 30)
            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 31)) Then
                txtBuyer.Text = "-"
            Else
                txtBuyer.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 31)
            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 32)) Then
                txtrun.Text = 0
            Else
                txtrun.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 32)
            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 33)) Then
                txtfix.Text = 0
            Else
                txtfix.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 33)
            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 34)) Then
                txtInsp.Text = 0
            Else
                txtInsp.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 34)
            End If

            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 35)) Then
                txtChildDesc.Text = "-"
            Else
                txtChildDesc.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 35)
            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 36)) Then
                txtProdLine.Text = 0
            Else
                txtProdLine.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 36)
            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 37)) Then
                txtInventoryAc.Text = "-"
            Else
                txtInventoryAc.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 37)
            End If
            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 38)) Then
                txtInsp.Text = "-"
            Else
                txtSpNote1.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 38)
            End If

            If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 39)) Then
                txtSpeNote2.Text = "-"
            Else
                txtSpeNote2.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 39)
            End If


            If RadioButtonPending.Checked = True Then

                If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 40)) Then
                    txtSpInsPurchaseDept.Text = "-"
                Else


                    txtSpInsPurchaseDept.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 40)
                End If


            End If

            If RadioButtoncompleted.Checked = True Or RadioButtonAll.Checked = True Then
                txtPartAppApl.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 40)
                txtAplSpecialInst.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 41)

                If IsDBNull(datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 42)) Then
                    txtSpInsPurchaseDept.Text = "-"
                Else
                    txtSpInsPurchaseDept.Text = datagridPartCreation.Item(datagridPartCreation.CurrentCell.RowNumber, 42)
                End If


            End If



            'FS_Yes_NO, Tender_YesNo, MarketType, Item_Type, Planner, Buyer, Lead_Run, Lead_Fix, Lead_Insp, Child_desc, prod_Line, Inv_Ac, Sp_note1, sp_note2




        Else
            MsgBox("Click on Detail code ", vbInformation)
            Exit Sub

        End If


    End Sub

    Private Sub datagridStock_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles datagridPartCreation.Navigate

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonDocYES.CheckedChanged

    End Sub

    Private Sub Label36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label36.Click

    End Sub

    Private Sub txtLotDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPartNo.TextChanged

    End Sub

    Private Sub ButtonRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRefresh.Click


        If RadioButtonAll.Checked = True Or RadioButtoncompleted.Checked = True Then
            btnUpdate.Enabled = False
        Else
            btnUpdate.Enabled = True

        End If

        datagridPartCreation.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        Dim stockDC As DataSet = New DataSet

        If RadioButtonPending.Checked = True Then

            strSQL = "SELECT  Enq_Detail_code, RegNo, SlNo, [Reg.Date], CustomerID, CustomerName,City, Part_Source, PartNumber, PartDescription, CustPartNumber, CustPartDescription," & _
            "uom,  Dimension, Material, Special_instructions, Enq_Due_date,Req, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
            "Special, Part_Source, Tender_YesNo, MarketType,Item_Type, Planner, Buyer, Lead_Run, Lead_Fix, Lead_Insp, Child_desc, prod_Line, Inv_Ac, Sp_note1, sp_note2,specialpur " & _
            " FROM ENQ_Parts_Created_Q where (Part_No_Appl_Sug = '-'or Part_No_Appl_Sug = '') order by  RegNo, SlNo "

        ElseIf RadioButtoncompleted.Checked = True Then
            strSQL = "SELECT  Enq_Detail_code, RegNo, SlNo, [Reg.Date], CustomerID, CustomerName,City, Part_Source, PartNumber, PartDescription, CustPartNumber, CustPartDescription," & _
               "uom,  Dimension, Material, Special_instructions, Enq_Due_date,Req, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
               "Special,  Part_Source, Tender_YesNo, MarketType,Item_Type, Planner, Buyer, Lead_Run, Lead_Fix, Lead_Insp, Child_desc, prod_Line, Inv_Ac, Sp_note1, sp_note2,Special_Inst_Apl,Part_No_Appl_Sug,specialpur " & _
               " FROM ENQ_Parts_Created_Q where len(Part_No_Appl_Sug)>6 order by  RegNo, SlNo "

            'dbo.ENQ_Details.Special_Inst_Apl, ISNULL(dbo.ENQ_Details.Part_No_Appl_Sug, '-') AS Part_No_Appl_Sug

        ElseIf RadioButtonAll.Checked = True Then
            strSQL = "SELECT  Enq_Detail_code, RegNo, SlNo, [Reg.Date], CustomerID, CustomerName,City, Part_Source, PartNumber, PartDescription, CustPartNumber, CustPartDescription," & _
           "uom,  Dimension, Material, Special_instructions, Enq_Due_date,Req, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
           "Special,  Part_Source,Tender_YesNo, MarketType,Item_Type, Planner, Buyer, Lead_Run, Lead_Fix, Lead_Insp, Child_desc, prod_Line, Inv_Ac, Sp_note1, sp_note2,Special_Inst_Apl,Part_No_Appl_Sug,specialpur " & _
           " FROM ENQ_Parts_Created_Q order by  RegNo, SlNo "

        End If

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        stockDAC.Fill(stockDC)

        datagridPartCreation.DataSource = stockDC.Tables(0)
        'cnSQL.Close()
        datagridPartCreation.Expand(-1)


    End Sub

    Private Sub ComboBoxItemSource_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        curdate = System.DateTime.Now()

        txtPartAppApl.Text = UCase(txtPartAppApl.Text)

        If Len(Trim(txtPartAppApl.Text)) >= 5 Or Len(Trim(txtAplSpecialInst.Text)) >= 3 Then
            strsql = "update ENQ_Details  set Item_Created_Date = '" & curdate & "',Item_Created_By = '" & username & "',Part_No_Appl_Sug = '" & txtPartAppApl.Text & "', Special_Inst_Apl ='" & txtAplSpecialInst.Text & "' WHERE Enq_Detail_code = " & txtDetailIntcode.Text & ""



            cnSQL.Open()
            cmSQL = New SqlCommand(strsql, cnSQL)

        End If

        If cmSQL.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot Save PartNumber. " & strsql, MsgBoxStyle.Exclamation, "Error!")
            Application.Exit()

        Else
            MsgBox("Part Number saved ", vbInformation)
            Exit Sub

        End If


    End Sub

    Private Sub txtMaterial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaterial.TextChanged

    End Sub

    Private Sub txtPartAppApl_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPartAppApl.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtPartAppApl_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPartAppApl.TextChanged

    End Sub

    Private Sub txtAplSpecialInst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAplSpecialInst.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtAplSpecialInst_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAplSpecialInst.TextChanged

    End Sub
End Class
