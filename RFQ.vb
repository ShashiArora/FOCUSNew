Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports outlook = Microsoft.Office.Interop.Outlook
Imports Excel = Microsoft.Office.Interop.Excel



'Imports SoftBrands.FourthShift.Transaction
'Imports CrystalDecisions.CrystalReports.Engine
Imports System.Windows.Forms


Public Class RFQ

 
    Inherits System.Windows.Forms.Form

    Private ConnectionString As String
    Public stockDA As SqlDataAdapter = New SqlDataAdapter
    Public countqty As Integer
    Public countcerqty As Integer
    Public rfqmode As String
    Friend WithEvents ToolCost As System.Windows.Forms.GroupBox
    Friend WithEvents RBToolNo As System.Windows.Forms.RadioButton
    Friend WithEvents RBToolYes As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtPurSpecial As System.Windows.Forms.TextBox
    Friend WithEvents ToolDetails As System.Windows.Forms.GroupBox
    Friend WithEvents ProtoTotal As System.Windows.Forms.TextBox
    Friend WithEvents lblTotalCost As System.Windows.Forms.Label
    Friend WithEvents ProdTotalCost As System.Windows.Forms.TextBox
    Friend WithEvents ProdLeadTime As System.Windows.Forms.TextBox
    Friend WithEvents ProtoLeadTime As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents ProdCustShare As System.Windows.Forms.TextBox
    Friend WithEvents ProtoCustShare As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ProdLifeofTool As System.Windows.Forms.TextBox
    Friend WithEvents ProtoLifeofTool As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents ProdQty As System.Windows.Forms.TextBox
    Friend WithEvents ProtoQty As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents LabelToolClose As System.Windows.Forms.Label
    Friend WithEvents lblForwardApl As System.Windows.Forms.Label
    Friend WithEvents txtreasonforwarding As System.Windows.Forms.TextBox
    Friend WithEvents txtMov As System.Windows.Forms.TextBox
    Friend WithEvents lblMov As System.Windows.Forms.Label
    Protected WithEvents DataGridCertificateCharges As System.Windows.Forms.DataGrid
    Friend WithEvents LatestAction As System.Windows.Forms.TextBox
    Friend WithEvents ToolFrameOpen As System.Windows.Forms.CheckBox
    Friend WithEvents lblProd As System.Windows.Forms.Label
    Friend WithEvents lblProto As System.Windows.Forms.Label
    Friend WithEvents txtLatestAction As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtVendorName As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Vendor As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonVendorNo As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonVendorYes As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3P As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonGroup As System.Windows.Forms.RadioButton
    Friend WithEvents DataGridVendor As System.Windows.Forms.DataGrid
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonDomestic As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonExport As System.Windows.Forms.RadioButton
    Friend WithEvents txtTSSSeg As System.Windows.Forms.TextBox
    Friend WithEvents txtTSSISeg As System.Windows.Forms.TextBox
    Friend WithEvents TXTCL3 As System.Windows.Forms.TextBox
    Friend WithEvents TXTCL1 As System.Windows.Forms.TextBox
    Friend WithEvents txtISR As System.Windows.Forms.TextBox
    Friend WithEvents txtCSR As System.Windows.Forms.TextBox
    Friend WithEvents datagridEnquiryPending As System.Windows.Forms.DataGridView
    Friend WithEvents datagridEnquiryPending1 As System.Windows.Forms.DataGrid
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnMail As System.Windows.Forms.Button
    Friend WithEvents BtnHistory As System.Windows.Forms.Button
    Friend WithEvents txtintcode As System.Windows.Forms.TextBox
    Friend WithEvents txtcustcode As System.Windows.Forms.TextBox
    Friend WithEvents txtQuoteRef3 As System.Windows.Forms.TextBox
    Friend WithEvents txtQuoteRef2 As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtQuoteRef1 As System.Windows.Forms.TextBox
    Friend WithEvents BtnQuoteRefSave As System.Windows.Forms.Button
    Friend WithEvents BtnClear As System.Windows.Forms.Button
    Friend WithEvents Btnsearch As System.Windows.Forms.Button
    Friend WithEvents txtReg As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents txtPartN As System.Windows.Forms.TextBox
    Friend WithEvents txtCustN As System.Windows.Forms.TextBox
    Friend WithEvents txtCID As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents txtAltMtrl As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents RadioButtonFactory As System.Windows.Forms.RadioButton
    Public multiple As String

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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DataUpdation As System.Windows.Forms.GroupBox
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
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
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents txtpart As System.Windows.Forms.TextBox
    Friend WithEvents txtCustPart As System.Windows.Forms.TextBox
    Friend WithEvents txtPartDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtslno As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterial As System.Windows.Forms.TextBox
    Friend WithEvents txtDetailSpecial As System.Windows.Forms.TextBox
    Friend WithEvents BtnAddPrice As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxPriceStatus As System.Windows.Forms.ComboBox
    Friend WithEvents l As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtSpecialInst As System.Windows.Forms.TextBox
    Friend WithEvents txtdocdetails As System.Windows.Forms.TextBox
    Friend WithEvents RadioButtondocno As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtondocyes As System.Windows.Forms.RadioButton
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomer As System.Windows.Forms.TextBox
    Friend WithEvents DtpEnqRegDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtRegNo As System.Windows.Forms.TextBox
    Friend WithEvents txtenqdetailintcode As System.Windows.Forms.TextBox
    Friend WithEvents dtpenqduedt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtvendquote As System.Windows.Forms.TextBox
    Friend WithEvents txtvendorref As System.Windows.Forms.TextBox
    Friend WithEvents txtvendorcontact As System.Windows.Forms.TextBox
    Friend WithEvents txtprogressdetails As System.Windows.Forms.TextBox
    Friend WithEvents txtstockavble As System.Windows.Forms.TextBox
    Friend WithEvents txtSPU As System.Windows.Forms.TextBox
    Friend WithEvents txtMOQ As System.Windows.Forms.TextBox
    Friend WithEvents txtLeadTime As System.Windows.Forms.TextBox
    Friend WithEvents txtdetailremarks As System.Windows.Forms.TextBox
    Friend WithEvents btnRFQSave As System.Windows.Forms.Button
    Friend WithEvents comboboxcurrency As System.Windows.Forms.ComboBox
    Friend WithEvents comboboxstocktype As System.Windows.Forms.ComboBox
    Friend WithEvents txtRFQIntcode As System.Windows.Forms.TextBox
    Protected WithEvents DataGridQty As System.Windows.Forms.DataGrid
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonNew As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonExisting As System.Windows.Forms.RadioButton
    Friend WithEvents txtitemstatus As System.Windows.Forms.TextBox
    Friend WithEvents btnnext As System.Windows.Forms.Button
    Protected WithEvents DatagridMultiprices As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RFQ))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtcustcode = New System.Windows.Forms.TextBox()
        Me.txtintcode = New System.Windows.Forms.TextBox()
        Me.BtnHistory = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.txtTSSSeg = New System.Windows.Forms.TextBox()
        Me.txtTSSISeg = New System.Windows.Forms.TextBox()
        Me.TXTCL3 = New System.Windows.Forms.TextBox()
        Me.TXTCL1 = New System.Windows.Forms.TextBox()
        Me.txtISR = New System.Windows.Forms.TextBox()
        Me.txtCSR = New System.Windows.Forms.TextBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonDomestic = New System.Windows.Forms.RadioButton()
        Me.RadioButtonExport = New System.Windows.Forms.RadioButton()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.DatagridMultiprices = New System.Windows.Forms.DataGrid()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonNew = New System.Windows.Forms.RadioButton()
        Me.RadioButtonExisting = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpenqduedt = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSpecialInst = New System.Windows.Forms.TextBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.txtdocdetails = New System.Windows.Forms.TextBox()
        Me.RadioButtondocno = New System.Windows.Forms.RadioButton()
        Me.RadioButtondocyes = New System.Windows.Forms.RadioButton()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtCustomer = New System.Windows.Forms.TextBox()
        Me.txtenqdetailintcode = New System.Windows.Forms.TextBox()
        Me.DtpEnqRegDt = New System.Windows.Forms.DateTimePicker()
        Me.txtRegNo = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.Btnsearch = New System.Windows.Forms.Button()
        Me.txtReg = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.txtPartN = New System.Windows.Forms.TextBox()
        Me.txtCustN = New System.Windows.Forms.TextBox()
        Me.txtCID = New System.Windows.Forms.TextBox()
        Me.datagridEnquiryPending = New System.Windows.Forms.DataGridView()
        Me.DataUpdation = New System.Windows.Forms.GroupBox()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.txtAltMtrl = New System.Windows.Forms.TextBox()
        Me.txtQuoteRef3 = New System.Windows.Forms.TextBox()
        Me.BtnMail = New System.Windows.Forms.Button()
        Me.datagridEnquiryPending1 = New System.Windows.Forms.DataGrid()
        Me.DataGridVendor = New System.Windows.Forms.DataGrid()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonFactory = New System.Windows.Forms.RadioButton()
        Me.RadioButton3P = New System.Windows.Forms.RadioButton()
        Me.RadioButtonGroup = New System.Windows.Forms.RadioButton()
        Me.Vendor = New System.Windows.Forms.GroupBox()
        Me.RadioButtonVendorYes = New System.Windows.Forms.RadioButton()
        Me.RadioButtonVendorNo = New System.Windows.Forms.RadioButton()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtVendorName = New System.Windows.Forms.TextBox()
        Me.ToolDetails = New System.Windows.Forms.GroupBox()
        Me.lblProd = New System.Windows.Forms.Label()
        Me.lblProto = New System.Windows.Forms.Label()
        Me.LabelToolClose = New System.Windows.Forms.Label()
        Me.ProdLifeofTool = New System.Windows.Forms.TextBox()
        Me.ProtoLifeofTool = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.ProdQty = New System.Windows.Forms.TextBox()
        Me.ProtoQty = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.ProdLeadTime = New System.Windows.Forms.TextBox()
        Me.ProtoLeadTime = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.ProdCustShare = New System.Windows.Forms.TextBox()
        Me.ProtoCustShare = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.ProdTotalCost = New System.Windows.Forms.TextBox()
        Me.ProtoTotal = New System.Windows.Forms.TextBox()
        Me.lblTotalCost = New System.Windows.Forms.Label()
        Me.txtLatestAction = New System.Windows.Forms.TextBox()
        Me.DataGridCertificateCharges = New System.Windows.Forms.DataGrid()
        Me.LatestAction = New System.Windows.Forms.TextBox()
        Me.txtMov = New System.Windows.Forms.TextBox()
        Me.lblMov = New System.Windows.Forms.Label()
        Me.lblForwardApl = New System.Windows.Forms.Label()
        Me.txtreasonforwarding = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtPurSpecial = New System.Windows.Forms.TextBox()
        Me.ToolCost = New System.Windows.Forms.GroupBox()
        Me.ToolFrameOpen = New System.Windows.Forms.CheckBox()
        Me.RBToolNo = New System.Windows.Forms.RadioButton()
        Me.RBToolYes = New System.Windows.Forms.RadioButton()
        Me.btnnext = New System.Windows.Forms.Button()
        Me.txtitemstatus = New System.Windows.Forms.TextBox()
        Me.txtRFQIntcode = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.txtvendquote = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtvendorref = New System.Windows.Forms.TextBox()
        Me.txtvendorcontact = New System.Windows.Forms.TextBox()
        Me.txtprogressdetails = New System.Windows.Forms.TextBox()
        Me.l = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ComboBoxPriceStatus = New System.Windows.Forms.ComboBox()
        Me.DataGridQty = New System.Windows.Forms.DataGrid()
        Me.BtnAddPrice = New System.Windows.Forms.Button()
        Me.ComboBoxuom = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.ComboBoxItemSource = New System.Windows.Forms.ComboBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.txtRecVend = New System.Windows.Forms.TextBox()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtDimension = New System.Windows.Forms.TextBox()
        Me.txtCustDesc = New System.Windows.Forms.TextBox()
        Me.ComboBoxFSYesNo = New System.Windows.Forms.ComboBox()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.txtpart = New System.Windows.Forms.TextBox()
        Me.txtCustPart = New System.Windows.Forms.TextBox()
        Me.txtPartDesc = New System.Windows.Forms.TextBox()
        Me.txtslno = New System.Windows.Forms.TextBox()
        Me.txtMaterial = New System.Windows.Forms.TextBox()
        Me.txtDetailSpecial = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtstockavble = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.comboboxcurrency = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.comboboxstocktype = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSPU = New System.Windows.Forms.TextBox()
        Me.txtMOQ = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.txtLeadTime = New System.Windows.Forms.TextBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.BtnDelete = New System.Windows.Forms.Button()
        Me.btnRFQSave = New System.Windows.Forms.Button()
        Me.txtdetailremarks = New System.Windows.Forms.TextBox()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.BtnQuoteRefSave = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtQuoteRef1 = New System.Windows.Forms.TextBox()
        Me.txtQuoteRef2 = New System.Windows.Forms.TextBox()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.DatagridMultiprices, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.datagridEnquiryPending, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DataUpdation.SuspendLayout()
        CType(Me.datagridEnquiryPending1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridVendor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.Vendor.SuspendLayout()
        Me.ToolDetails.SuspendLayout()
        CType(Me.DataGridCertificateCharges, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolCost.SuspendLayout()
        CType(Me.DataGridQty, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(558, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(139, 16)
        Me.Label2.TabIndex = 112
        Me.Label2.Text = "Enquiry Reg. No. and Dt"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox3.Controls.Add(Me.txtcustcode)
        Me.GroupBox3.Controls.Add(Me.txtintcode)
        Me.GroupBox3.Controls.Add(Me.BtnHistory)
        Me.GroupBox3.Controls.Add(Me.GroupBox5)
        Me.GroupBox3.Controls.Add(Me.txtTSSSeg)
        Me.GroupBox3.Controls.Add(Me.txtTSSISeg)
        Me.GroupBox3.Controls.Add(Me.TXTCL3)
        Me.GroupBox3.Controls.Add(Me.TXTCL1)
        Me.GroupBox3.Controls.Add(Me.txtISR)
        Me.GroupBox3.Controls.Add(Me.txtCSR)
        Me.GroupBox3.Controls.Add(Me.GroupBox6)
        Me.GroupBox3.Controls.Add(Me.Label23)
        Me.GroupBox3.Controls.Add(Me.DatagridMultiprices)
        Me.GroupBox3.Controls.Add(Me.GroupBox1)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.dtpenqduedt)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.txtSpecialInst)
        Me.GroupBox3.Controls.Add(Me.Label35)
        Me.GroupBox3.Controls.Add(Me.txtdocdetails)
        Me.GroupBox3.Controls.Add(Me.RadioButtondocno)
        Me.GroupBox3.Controls.Add(Me.RadioButtondocyes)
        Me.GroupBox3.Controls.Add(Me.txtCity)
        Me.GroupBox3.Controls.Add(Me.txtCustomer)
        Me.GroupBox3.Controls.Add(Me.txtenqdetailintcode)
        Me.GroupBox3.Controls.Add(Me.DtpEnqRegDt)
        Me.GroupBox3.Controls.Add(Me.txtRegNo)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.Black
        Me.GroupBox3.Location = New System.Drawing.Point(17, 281)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1271, 91)
        Me.GroupBox3.TabIndex = 115
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Customer Details"
        '
        'txtcustcode
        '
        Me.txtcustcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcustcode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtcustcode.Enabled = False
        Me.txtcustcode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcustcode.Location = New System.Drawing.Point(423, 9)
        Me.txtcustcode.MaxLength = 50
        Me.txtcustcode.Name = "txtcustcode"
        Me.txtcustcode.Size = New System.Drawing.Size(34, 20)
        Me.txtcustcode.TabIndex = 235
        Me.txtcustcode.Visible = False
        '
        'txtintcode
        '
        Me.txtintcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtintcode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtintcode.Enabled = False
        Me.txtintcode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtintcode.Location = New System.Drawing.Point(387, 7)
        Me.txtintcode.MaxLength = 50
        Me.txtintcode.Name = "txtintcode"
        Me.txtintcode.Size = New System.Drawing.Size(31, 20)
        Me.txtintcode.TabIndex = 234
        Me.txtintcode.Visible = False
        '
        'BtnHistory
        '
        Me.BtnHistory.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.BtnHistory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnHistory.ForeColor = System.Drawing.Color.Red
        Me.BtnHistory.Location = New System.Drawing.Point(457, 7)
        Me.BtnHistory.Name = "BtnHistory"
        Me.BtnHistory.Size = New System.Drawing.Size(86, 24)
        Me.BtnHistory.TabIndex = 233
        Me.BtnHistory.Text = "History"
        Me.BtnHistory.UseVisualStyleBackColor = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(575, 36)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(692, 41)
        Me.GroupBox5.TabIndex = 216
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Tool Cost"
        Me.GroupBox5.Visible = False
        '
        'txtTSSSeg
        '
        Me.txtTSSSeg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTSSSeg.Enabled = False
        Me.txtTSSSeg.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTSSSeg.Location = New System.Drawing.Point(1210, 66)
        Me.txtTSSSeg.Name = "txtTSSSeg"
        Me.txtTSSSeg.Size = New System.Drawing.Size(57, 20)
        Me.txtTSSSeg.TabIndex = 214
        '
        'txtTSSISeg
        '
        Me.txtTSSISeg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTSSISeg.Enabled = False
        Me.txtTSSISeg.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTSSISeg.Location = New System.Drawing.Point(1133, 66)
        Me.txtTSSISeg.Name = "txtTSSISeg"
        Me.txtTSSISeg.Size = New System.Drawing.Size(57, 20)
        Me.txtTSSISeg.TabIndex = 213
        '
        'TXTCL3
        '
        Me.TXTCL3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TXTCL3.Enabled = False
        Me.TXTCL3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TXTCL3.Location = New System.Drawing.Point(1210, 35)
        Me.TXTCL3.Name = "TXTCL3"
        Me.TXTCL3.Size = New System.Drawing.Size(57, 20)
        Me.TXTCL3.TabIndex = 212
        '
        'TXTCL1
        '
        Me.TXTCL1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TXTCL1.Enabled = False
        Me.TXTCL1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TXTCL1.Location = New System.Drawing.Point(1133, 38)
        Me.TXTCL1.Name = "TXTCL1"
        Me.TXTCL1.Size = New System.Drawing.Size(57, 20)
        Me.TXTCL1.TabIndex = 211
        '
        'txtISR
        '
        Me.txtISR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtISR.Enabled = False
        Me.txtISR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtISR.Location = New System.Drawing.Point(1210, 9)
        Me.txtISR.Name = "txtISR"
        Me.txtISR.Size = New System.Drawing.Size(57, 20)
        Me.txtISR.TabIndex = 210
        '
        'txtCSR
        '
        Me.txtCSR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCSR.Enabled = False
        Me.txtCSR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCSR.Location = New System.Drawing.Point(1133, 10)
        Me.txtCSR.Name = "txtCSR"
        Me.txtCSR.Size = New System.Drawing.Size(57, 20)
        Me.txtCSR.TabIndex = 209
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.RadioButtonDomestic)
        Me.GroupBox6.Controls.Add(Me.RadioButtonExport)
        Me.GroupBox6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox6.ForeColor = System.Drawing.Color.Black
        Me.GroupBox6.Location = New System.Drawing.Point(143, 0)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(217, 28)
        Me.GroupBox6.TabIndex = 208
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Market Type"
        '
        'RadioButtonDomestic
        '
        Me.RadioButtonDomestic.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonDomestic.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonDomestic.Location = New System.Drawing.Point(68, 8)
        Me.RadioButtonDomestic.Name = "RadioButtonDomestic"
        Me.RadioButtonDomestic.Size = New System.Drawing.Size(72, 18)
        Me.RadioButtonDomestic.TabIndex = 52
        Me.RadioButtonDomestic.Text = "Domestic"
        '
        'RadioButtonExport
        '
        Me.RadioButtonExport.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonExport.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonExport.Location = New System.Drawing.Point(147, 10)
        Me.RadioButtonExport.Name = "RadioButtonExport"
        Me.RadioButtonExport.Size = New System.Drawing.Size(65, 16)
        Me.RadioButtonExport.TabIndex = 53
        Me.RadioButtonExport.Text = "Export"
        '
        'Label23
        '
        Me.Label23.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.Black
        Me.Label23.Location = New System.Drawing.Point(702, 8)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(81, 18)
        Me.Label23.TabIndex = 197
        Me.Label23.Text = "Enq, Det Code"
        '
        'DatagridMultiprices
        '
        Me.DatagridMultiprices.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DatagridMultiprices.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatagridMultiprices.CaptionForeColor = System.Drawing.Color.Black
        Me.DatagridMultiprices.CaptionVisible = False
        Me.DatagridMultiprices.DataMember = ""
        Me.DatagridMultiprices.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatagridMultiprices.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatagridMultiprices.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DatagridMultiprices.Location = New System.Drawing.Point(120, 8)
        Me.DatagridMultiprices.Name = "DatagridMultiprices"
        Me.DatagridMultiprices.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DatagridMultiprices.ParentRowsVisible = False
        Me.DatagridMultiprices.PreferredColumnWidth = 85
        Me.DatagridMultiprices.RowHeadersVisible = False
        Me.DatagridMultiprices.RowHeaderWidth = 20
        Me.DatagridMultiprices.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DatagridMultiprices.Size = New System.Drawing.Size(28, 10)
        Me.DatagridMultiprices.TabIndex = 196
        Me.DatagridMultiprices.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButtonNew)
        Me.GroupBox1.Controls.Add(Me.RadioButtonExisting)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 19)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(119, 38)
        Me.GroupBox1.TabIndex = 165
        Me.GroupBox1.TabStop = False
        '
        'RadioButtonNew
        '
        Me.RadioButtonNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonNew.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonNew.Location = New System.Drawing.Point(67, 12)
        Me.RadioButtonNew.Name = "RadioButtonNew"
        Me.RadioButtonNew.Size = New System.Drawing.Size(46, 17)
        Me.RadioButtonNew.TabIndex = 140
        Me.RadioButtonNew.Text = "New"
        '
        'RadioButtonExisting
        '
        Me.RadioButtonExisting.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonExisting.ForeColor = System.Drawing.Color.Black
        Me.RadioButtonExisting.Location = New System.Drawing.Point(7, 10)
        Me.RadioButtonExisting.Name = "RadioButtonExisting"
        Me.RadioButtonExisting.Size = New System.Drawing.Size(63, 19)
        Me.RadioButtonExisting.TabIndex = 139
        Me.RadioButtonExisting.Text = "Existing"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(802, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 163
        Me.Label3.Text = "Enq. Due Dt."
        '
        'dtpenqduedt
        '
        Me.dtpenqduedt.CalendarFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpenqduedt.Checked = False
        Me.dtpenqduedt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpenqduedt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpenqduedt.Location = New System.Drawing.Point(783, 27)
        Me.dtpenqduedt.Name = "dtpenqduedt"
        Me.dtpenqduedt.ShowCheckBox = True
        Me.dtpenqduedt.Size = New System.Drawing.Size(105, 20)
        Me.dtpenqduedt.TabIndex = 162
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(8, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 16)
        Me.Label1.TabIndex = 147
        Me.Label1.Text = "Special Instructions"
        '
        'txtSpecialInst
        '
        Me.txtSpecialInst.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSpecialInst.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpecialInst.Location = New System.Drawing.Point(110, 55)
        Me.txtSpecialInst.Multiline = True
        Me.txtSpecialInst.Name = "txtSpecialInst"
        Me.txtSpecialInst.Size = New System.Drawing.Size(758, 31)
        Me.txtSpecialInst.TabIndex = 146
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label35.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.ForeColor = System.Drawing.Color.Black
        Me.Label35.Location = New System.Drawing.Point(893, 8)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(110, 19)
        Me.Label35.TabIndex = 135
        Me.Label35.Text = "Document uploaded:"
        '
        'txtdocdetails
        '
        Me.txtdocdetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtdocdetails.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdocdetails.ForeColor = System.Drawing.Color.Red
        Me.txtdocdetails.Location = New System.Drawing.Point(893, 29)
        Me.txtdocdetails.Multiline = True
        Me.txtdocdetails.Name = "txtdocdetails"
        Me.txtdocdetails.Size = New System.Drawing.Size(235, 57)
        Me.txtdocdetails.TabIndex = 134
        '
        'RadioButtondocno
        '
        Me.RadioButtondocno.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.RadioButtondocno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtondocno.ForeColor = System.Drawing.Color.Black
        Me.RadioButtondocno.Location = New System.Drawing.Point(1062, 12)
        Me.RadioButtondocno.Name = "RadioButtondocno"
        Me.RadioButtondocno.Size = New System.Drawing.Size(65, 16)
        Me.RadioButtondocno.TabIndex = 114
        Me.RadioButtondocno.Text = "No"
        Me.RadioButtondocno.UseVisualStyleBackColor = False
        '
        'RadioButtondocyes
        '
        Me.RadioButtondocyes.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.RadioButtondocyes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtondocyes.ForeColor = System.Drawing.Color.Black
        Me.RadioButtondocyes.Location = New System.Drawing.Point(1008, 10)
        Me.RadioButtondocyes.Name = "RadioButtondocyes"
        Me.RadioButtondocyes.Size = New System.Drawing.Size(49, 16)
        Me.RadioButtondocyes.TabIndex = 113
        Me.RadioButtondocyes.Text = "Yes"
        Me.RadioButtondocyes.UseVisualStyleBackColor = False
        '
        'txtCity
        '
        Me.txtCity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCity.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCity.Location = New System.Drawing.Point(398, 30)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(150, 20)
        Me.txtCity.TabIndex = 85
        '
        'txtCustomer
        '
        Me.txtCustomer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustomer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomer.Location = New System.Drawing.Point(128, 30)
        Me.txtCustomer.Name = "txtCustomer"
        Me.txtCustomer.Size = New System.Drawing.Size(265, 20)
        Me.txtCustomer.TabIndex = 75
        '
        'txtenqdetailintcode
        '
        Me.txtenqdetailintcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtenqdetailintcode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtenqdetailintcode.Enabled = False
        Me.txtenqdetailintcode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtenqdetailintcode.Location = New System.Drawing.Point(705, 29)
        Me.txtenqdetailintcode.MaxLength = 50
        Me.txtenqdetailintcode.Name = "txtenqdetailintcode"
        Me.txtenqdetailintcode.Size = New System.Drawing.Size(78, 20)
        Me.txtenqdetailintcode.TabIndex = 207
        '
        'DtpEnqRegDt
        '
        Me.DtpEnqRegDt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DtpEnqRegDt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpEnqRegDt.Location = New System.Drawing.Point(610, 29)
        Me.DtpEnqRegDt.Name = "DtpEnqRegDt"
        Me.DtpEnqRegDt.Size = New System.Drawing.Size(97, 20)
        Me.DtpEnqRegDt.TabIndex = 116
        '
        'txtRegNo
        '
        Me.txtRegNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRegNo.Enabled = False
        Me.txtRegNo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegNo.Location = New System.Drawing.Point(552, 30)
        Me.txtRegNo.Name = "txtRegNo"
        Me.txtRegNo.Size = New System.Drawing.Size(56, 22)
        Me.txtRegNo.TabIndex = 114
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.BtnClear)
        Me.GroupBox2.Controls.Add(Me.Btnsearch)
        Me.GroupBox2.Controls.Add(Me.txtReg)
        Me.GroupBox2.Controls.Add(Me.Label30)
        Me.GroupBox2.Controls.Add(Me.Label29)
        Me.GroupBox2.Controls.Add(Me.Label28)
        Me.GroupBox2.Controls.Add(Me.Label27)
        Me.GroupBox2.Controls.Add(Me.txtPartN)
        Me.GroupBox2.Controls.Add(Me.txtCustN)
        Me.GroupBox2.Controls.Add(Me.txtCID)
        Me.GroupBox2.Controls.Add(Me.datagridEnquiryPending)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox2.Location = New System.Drawing.Point(17, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1271, 267)
        Me.GroupBox2.TabIndex = 109
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "1"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Silver
        Me.Button1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Red
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(1092, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 24)
        Me.Button1.TabIndex = 246
        Me.Button1.Text = "Load"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'BtnClear
        '
        Me.BtnClear.BackColor = System.Drawing.Color.Silver
        Me.BtnClear.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnClear.ForeColor = System.Drawing.Color.Red
        Me.BtnClear.Image = CType(resources.GetObject("BtnClear.Image"), System.Drawing.Image)
        Me.BtnClear.Location = New System.Drawing.Point(998, 12)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(79, 24)
        Me.BtnClear.TabIndex = 245
        Me.BtnClear.Text = "Clear"
        Me.BtnClear.UseVisualStyleBackColor = False
        '
        'Btnsearch
        '
        Me.Btnsearch.BackColor = System.Drawing.Color.Silver
        Me.Btnsearch.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btnsearch.ForeColor = System.Drawing.Color.Red
        Me.Btnsearch.Image = CType(resources.GetObject("Btnsearch.Image"), System.Drawing.Image)
        Me.Btnsearch.Location = New System.Drawing.Point(913, 11)
        Me.Btnsearch.Name = "Btnsearch"
        Me.Btnsearch.Size = New System.Drawing.Size(69, 25)
        Me.Btnsearch.TabIndex = 244
        Me.Btnsearch.Text = "Filter"
        Me.Btnsearch.UseVisualStyleBackColor = False
        '
        'txtReg
        '
        Me.txtReg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtReg.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtReg.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReg.Location = New System.Drawing.Point(73, 16)
        Me.txtReg.MaxLength = 50
        Me.txtReg.Name = "txtReg"
        Me.txtReg.Size = New System.Drawing.Size(75, 20)
        Me.txtReg.TabIndex = 243
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.Color.Transparent
        Me.Label30.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.ForeColor = System.Drawing.Color.Black
        Me.Label30.Location = New System.Drawing.Point(13, 20)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(55, 16)
        Me.Label30.TabIndex = 242
        Me.Label30.Text = "Reg No."
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.Color.Transparent
        Me.Label29.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.Color.Black
        Me.Label29.Location = New System.Drawing.Point(620, 21)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(68, 18)
        Me.Label29.TabIndex = 241
        Me.Label29.Text = "Part Number"
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.Color.Transparent
        Me.Label28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.Color.Black
        Me.Label28.Location = New System.Drawing.Point(297, 20)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(66, 18)
        Me.Label28.TabIndex = 240
        Me.Label28.Text = "Cust Name"
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.Color.Transparent
        Me.Label27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.Black
        Me.Label27.Location = New System.Drawing.Point(160, 20)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(40, 16)
        Me.Label27.TabIndex = 239
        Me.Label27.Text = "Cust Id"
        '
        'txtPartN
        '
        Me.txtPartN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartN.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPartN.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartN.Location = New System.Drawing.Point(693, 16)
        Me.txtPartN.MaxLength = 50
        Me.txtPartN.Name = "txtPartN"
        Me.txtPartN.Size = New System.Drawing.Size(189, 20)
        Me.txtPartN.TabIndex = 238
        '
        'txtCustN
        '
        Me.txtCustN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustN.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustN.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustN.Location = New System.Drawing.Point(370, 16)
        Me.txtCustN.MaxLength = 50
        Me.txtCustN.Name = "txtCustN"
        Me.txtCustN.Size = New System.Drawing.Size(238, 20)
        Me.txtCustN.TabIndex = 237
        '
        'txtCID
        '
        Me.txtCID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCID.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCID.Location = New System.Drawing.Point(207, 16)
        Me.txtCID.MaxLength = 50
        Me.txtCID.Name = "txtCID"
        Me.txtCID.Size = New System.Drawing.Size(83, 20)
        Me.txtCID.TabIndex = 235
        '
        'datagridEnquiryPending
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.datagridEnquiryPending.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.datagridEnquiryPending.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.Blue
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.datagridEnquiryPending.DefaultCellStyle = DataGridViewCellStyle2
        Me.datagridEnquiryPending.Location = New System.Drawing.Point(17, 42)
        Me.datagridEnquiryPending.Name = "datagridEnquiryPending"
        Me.datagridEnquiryPending.Size = New System.Drawing.Size(1235, 219)
        Me.datagridEnquiryPending.TabIndex = 18
        '
        'DataUpdation
        '
        Me.DataUpdation.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataUpdation.Controls.Add(Me.Label31)
        Me.DataUpdation.Controls.Add(Me.txtAltMtrl)
        Me.DataUpdation.Controls.Add(Me.txtQuoteRef3)
        Me.DataUpdation.Controls.Add(Me.BtnMail)
        Me.DataUpdation.Controls.Add(Me.datagridEnquiryPending1)
        Me.DataUpdation.Controls.Add(Me.DataGridVendor)
        Me.DataUpdation.Controls.Add(Me.GroupBox4)
        Me.DataUpdation.Controls.Add(Me.Vendor)
        Me.DataUpdation.Controls.Add(Me.Label25)
        Me.DataUpdation.Controls.Add(Me.Label24)
        Me.DataUpdation.Controls.Add(Me.txtVendorName)
        Me.DataUpdation.Controls.Add(Me.ToolDetails)
        Me.DataUpdation.Controls.Add(Me.txtLatestAction)
        Me.DataUpdation.Controls.Add(Me.DataGridCertificateCharges)
        Me.DataUpdation.Controls.Add(Me.LatestAction)
        Me.DataUpdation.Controls.Add(Me.txtMov)
        Me.DataUpdation.Controls.Add(Me.lblMov)
        Me.DataUpdation.Controls.Add(Me.lblForwardApl)
        Me.DataUpdation.Controls.Add(Me.txtreasonforwarding)
        Me.DataUpdation.Controls.Add(Me.Label7)
        Me.DataUpdation.Controls.Add(Me.txtPurSpecial)
        Me.DataUpdation.Controls.Add(Me.ToolCost)
        Me.DataUpdation.Controls.Add(Me.btnnext)
        Me.DataUpdation.Controls.Add(Me.txtitemstatus)
        Me.DataUpdation.Controls.Add(Me.txtRFQIntcode)
        Me.DataUpdation.Controls.Add(Me.Label18)
        Me.DataUpdation.Controls.Add(Me.Label19)
        Me.DataUpdation.Controls.Add(Me.txtvendquote)
        Me.DataUpdation.Controls.Add(Me.Label11)
        Me.DataUpdation.Controls.Add(Me.txtvendorref)
        Me.DataUpdation.Controls.Add(Me.txtvendorcontact)
        Me.DataUpdation.Controls.Add(Me.txtprogressdetails)
        Me.DataUpdation.Controls.Add(Me.l)
        Me.DataUpdation.Controls.Add(Me.Label5)
        Me.DataUpdation.Controls.Add(Me.ComboBoxPriceStatus)
        Me.DataUpdation.Controls.Add(Me.DataGridQty)
        Me.DataUpdation.Controls.Add(Me.BtnAddPrice)
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
        Me.DataUpdation.Controls.Add(Me.Label41)
        Me.DataUpdation.Controls.Add(Me.Label40)
        Me.DataUpdation.Controls.Add(Me.Label39)
        Me.DataUpdation.Controls.Add(Me.Label16)
        Me.DataUpdation.Controls.Add(Me.Label37)
        Me.DataUpdation.Controls.Add(Me.Label36)
        Me.DataUpdation.Controls.Add(Me.txtpart)
        Me.DataUpdation.Controls.Add(Me.txtCustPart)
        Me.DataUpdation.Controls.Add(Me.txtPartDesc)
        Me.DataUpdation.Controls.Add(Me.txtslno)
        Me.DataUpdation.Controls.Add(Me.txtMaterial)
        Me.DataUpdation.Controls.Add(Me.txtDetailSpecial)
        Me.DataUpdation.Controls.Add(Me.Label14)
        Me.DataUpdation.Controls.Add(Me.txtstockavble)
        Me.DataUpdation.Controls.Add(Me.Label9)
        Me.DataUpdation.Controls.Add(Me.comboboxcurrency)
        Me.DataUpdation.Controls.Add(Me.Label8)
        Me.DataUpdation.Controls.Add(Me.comboboxstocktype)
        Me.DataUpdation.Controls.Add(Me.Label6)
        Me.DataUpdation.Controls.Add(Me.txtSPU)
        Me.DataUpdation.Controls.Add(Me.txtMOQ)
        Me.DataUpdation.Controls.Add(Me.Label4)
        Me.DataUpdation.Controls.Add(Me.Label43)
        Me.DataUpdation.Controls.Add(Me.txtLeadTime)
        Me.DataUpdation.Controls.Add(Me.Label38)
        Me.DataUpdation.Controls.Add(Me.BtnDelete)
        Me.DataUpdation.Controls.Add(Me.btnRFQSave)
        Me.DataUpdation.Controls.Add(Me.txtdetailremarks)
        Me.DataUpdation.Controls.Add(Me.GroupBox7)
        Me.DataUpdation.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataUpdation.ForeColor = System.Drawing.Color.Firebrick
        Me.DataUpdation.Location = New System.Drawing.Point(17, 378)
        Me.DataUpdation.Name = "DataUpdation"
        Me.DataUpdation.Size = New System.Drawing.Size(1271, 340)
        Me.DataUpdation.TabIndex = 110
        Me.DataUpdation.TabStop = False
        Me.DataUpdation.Text = "Details"
        '
        'Label31
        '
        Me.Label31.BackColor = System.Drawing.Color.Transparent
        Me.Label31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.ForeColor = System.Drawing.Color.Blue
        Me.Label31.Location = New System.Drawing.Point(403, 162)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(60, 32)
        Me.Label31.TabIndex = 239
        Me.Label31.Text = "Alternative Material"
        '
        'txtAltMtrl
        '
        Me.txtAltMtrl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAltMtrl.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAltMtrl.Location = New System.Drawing.Point(467, 162)
        Me.txtAltMtrl.MaxLength = 50
        Me.txtAltMtrl.Name = "txtAltMtrl"
        Me.txtAltMtrl.Size = New System.Drawing.Size(116, 20)
        Me.txtAltMtrl.TabIndex = 30
        '
        'txtQuoteRef3
        '
        Me.txtQuoteRef3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtQuoteRef3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuoteRef3.Location = New System.Drawing.Point(863, 305)
        Me.txtQuoteRef3.MaxLength = 100
        Me.txtQuoteRef3.Multiline = True
        Me.txtQuoteRef3.Name = "txtQuoteRef3"
        Me.txtQuoteRef3.Size = New System.Drawing.Size(309, 23)
        Me.txtQuoteRef3.TabIndex = 23
        '
        'BtnMail
        '
        Me.BtnMail.BackColor = System.Drawing.Color.Silver
        Me.BtnMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnMail.ForeColor = System.Drawing.Color.Black
        Me.BtnMail.Location = New System.Drawing.Point(1222, 212)
        Me.BtnMail.Name = "BtnMail"
        Me.BtnMail.Size = New System.Drawing.Size(48, 24)
        Me.BtnMail.TabIndex = 232
        Me.BtnMail.Text = "Mail"
        Me.BtnMail.UseVisualStyleBackColor = False
        Me.BtnMail.Visible = False
        '
        'datagridEnquiryPending1
        '
        Me.datagridEnquiryPending1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.datagridEnquiryPending1.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridEnquiryPending1.CaptionVisible = False
        Me.datagridEnquiryPending1.DataMember = ""
        Me.datagridEnquiryPending1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridEnquiryPending1.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridEnquiryPending1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.datagridEnquiryPending1.Location = New System.Drawing.Point(918, 9)
        Me.datagridEnquiryPending1.Name = "datagridEnquiryPending1"
        Me.datagridEnquiryPending1.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.datagridEnquiryPending1.ParentRowsVisible = False
        Me.datagridEnquiryPending1.PreferredColumnWidth = 85
        Me.datagridEnquiryPending1.ReadOnly = True
        Me.datagridEnquiryPending1.RowHeadersVisible = False
        Me.datagridEnquiryPending1.Size = New System.Drawing.Size(12, 17)
        Me.datagridEnquiryPending1.TabIndex = 231
        '
        'DataGridVendor
        '
        Me.DataGridVendor.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridVendor.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridVendor.CaptionVisible = False
        Me.DataGridVendor.DataMember = ""
        Me.DataGridVendor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridVendor.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridVendor.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridVendor.Location = New System.Drawing.Point(705, 9)
        Me.DataGridVendor.Name = "DataGridVendor"
        Me.DataGridVendor.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridVendor.ParentRowsVisible = False
        Me.DataGridVendor.PreferredColumnWidth = 85
        Me.DataGridVendor.ReadOnly = True
        Me.DataGridVendor.RowHeadersVisible = False
        Me.DataGridVendor.Size = New System.Drawing.Size(48, 17)
        Me.DataGridVendor.TabIndex = 230
        Me.DataGridVendor.Visible = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.RadioButtonFactory)
        Me.GroupBox4.Controls.Add(Me.RadioButton3P)
        Me.GroupBox4.Controls.Add(Me.RadioButtonGroup)
        Me.GroupBox4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.Color.Black
        Me.GroupBox4.Location = New System.Drawing.Point(8, 174)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(90, 77)
        Me.GroupBox4.TabIndex = 229
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Price From"
        '
        'RadioButtonFactory
        '
        Me.RadioButtonFactory.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.RadioButtonFactory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonFactory.ForeColor = System.Drawing.Color.Blue
        Me.RadioButtonFactory.Location = New System.Drawing.Point(7, 51)
        Me.RadioButtonFactory.Name = "RadioButtonFactory"
        Me.RadioButtonFactory.Size = New System.Drawing.Size(76, 20)
        Me.RadioButtonFactory.TabIndex = 9
        Me.RadioButtonFactory.Text = "Factory"
        Me.RadioButtonFactory.UseVisualStyleBackColor = False
        '
        'RadioButton3P
        '
        Me.RadioButton3P.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.RadioButton3P.Font = New System.Drawing.Font("Arial", 8.25!)
        Me.RadioButton3P.ForeColor = System.Drawing.Color.Blue
        Me.RadioButton3P.Location = New System.Drawing.Point(7, 32)
        Me.RadioButton3P.Name = "RadioButton3P"
        Me.RadioButton3P.Size = New System.Drawing.Size(76, 22)
        Me.RadioButton3P.TabIndex = 8
        Me.RadioButton3P.Text = "Non Group"
        Me.RadioButton3P.UseVisualStyleBackColor = False
        '
        'RadioButtonGroup
        '
        Me.RadioButtonGroup.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.RadioButtonGroup.Font = New System.Drawing.Font("Arial", 8.25!)
        Me.RadioButtonGroup.ForeColor = System.Drawing.Color.Blue
        Me.RadioButtonGroup.Location = New System.Drawing.Point(7, 16)
        Me.RadioButtonGroup.Name = "RadioButtonGroup"
        Me.RadioButtonGroup.Size = New System.Drawing.Size(55, 18)
        Me.RadioButtonGroup.TabIndex = 7
        Me.RadioButtonGroup.Text = "Group"
        Me.RadioButtonGroup.UseVisualStyleBackColor = False
        '
        'Vendor
        '
        Me.Vendor.Controls.Add(Me.RadioButtonVendorYes)
        Me.Vendor.Controls.Add(Me.RadioButtonVendorNo)
        Me.Vendor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Vendor.ForeColor = System.Drawing.Color.Red
        Me.Vendor.Location = New System.Drawing.Point(103, 189)
        Me.Vendor.Name = "Vendor"
        Me.Vendor.Size = New System.Drawing.Size(134, 36)
        Me.Vendor.TabIndex = 228
        Me.Vendor.TabStop = False
        Me.Vendor.Text = "Vendor Existing in FS"
        '
        'RadioButtonVendorYes
        '
        Me.RadioButtonVendorYes.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.RadioButtonVendorYes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonVendorYes.ForeColor = System.Drawing.Color.Red
        Me.RadioButtonVendorYes.Location = New System.Drawing.Point(7, 16)
        Me.RadioButtonVendorYes.Name = "RadioButtonVendorYes"
        Me.RadioButtonVendorYes.Size = New System.Drawing.Size(46, 18)
        Me.RadioButtonVendorYes.TabIndex = 116
        Me.RadioButtonVendorYes.Text = "Yes"
        Me.RadioButtonVendorYes.UseVisualStyleBackColor = False
        '
        'RadioButtonVendorNo
        '
        Me.RadioButtonVendorNo.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.RadioButtonVendorNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonVendorNo.ForeColor = System.Drawing.Color.Red
        Me.RadioButtonVendorNo.Location = New System.Drawing.Point(60, 16)
        Me.RadioButtonVendorNo.Name = "RadioButtonVendorNo"
        Me.RadioButtonVendorNo.Size = New System.Drawing.Size(48, 18)
        Me.RadioButtonVendorNo.TabIndex = 115
        Me.RadioButtonVendorNo.Text = "No"
        Me.RadioButtonVendorNo.UseVisualStyleBackColor = False
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.Color.Transparent
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.Red
        Me.Label25.Location = New System.Drawing.Point(103, 233)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(25, 18)
        Me.Label25.TabIndex = 226
        Me.Label25.Text = "I D"
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.Color.Transparent
        Me.Label24.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.Red
        Me.Label24.Location = New System.Drawing.Point(238, 212)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(97, 13)
        Me.Label24.TabIndex = 225
        Me.Label24.Text = "Vendor Name"
        '
        'txtVendorName
        '
        Me.txtVendorName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtVendorName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorName.Location = New System.Drawing.Point(185, 231)
        Me.txtVendorName.MaxLength = 100
        Me.txtVendorName.Name = "txtVendorName"
        Me.txtVendorName.Size = New System.Drawing.Size(177, 20)
        Me.txtVendorName.TabIndex = 224
        '
        'ToolDetails
        '
        Me.ToolDetails.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ToolDetails.Controls.Add(Me.lblProd)
        Me.ToolDetails.Controls.Add(Me.lblProto)
        Me.ToolDetails.Controls.Add(Me.LabelToolClose)
        Me.ToolDetails.Controls.Add(Me.ProdLifeofTool)
        Me.ToolDetails.Controls.Add(Me.ProtoLifeofTool)
        Me.ToolDetails.Controls.Add(Me.Label22)
        Me.ToolDetails.Controls.Add(Me.ProdQty)
        Me.ToolDetails.Controls.Add(Me.ProtoQty)
        Me.ToolDetails.Controls.Add(Me.Label21)
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
        Me.ToolDetails.Location = New System.Drawing.Point(575, 99)
        Me.ToolDetails.Name = "ToolDetails"
        Me.ToolDetails.Size = New System.Drawing.Size(90, 17)
        Me.ToolDetails.TabIndex = 218
        Me.ToolDetails.TabStop = False
        Me.ToolDetails.Text = "Tool Details"
        Me.ToolDetails.Visible = False
        '
        'lblProd
        '
        Me.lblProd.AutoSize = True
        Me.lblProd.Font = New System.Drawing.Font("Arial", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProd.ForeColor = System.Drawing.Color.Blue
        Me.lblProd.Location = New System.Drawing.Point(7, 63)
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
        Me.lblProto.Location = New System.Drawing.Point(3, 28)
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
        'ProdLifeofTool
        '
        Me.ProdLifeofTool.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdLifeofTool.Location = New System.Drawing.Point(485, 61)
        Me.ProdLifeofTool.Name = "ProdLifeofTool"
        Me.ProdLifeofTool.Size = New System.Drawing.Size(113, 21)
        Me.ProdLifeofTool.TabIndex = 28
        '
        'ProtoLifeofTool
        '
        Me.ProtoLifeofTool.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProtoLifeofTool.Location = New System.Drawing.Point(487, 29)
        Me.ProtoLifeofTool.Name = "ProtoLifeofTool"
        Me.ProtoLifeofTool.Size = New System.Drawing.Size(113, 21)
        Me.ProtoLifeofTool.TabIndex = 23
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Blue
        Me.Label22.Location = New System.Drawing.Point(482, 12)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(124, 14)
        Me.Label22.TabIndex = 14
        Me.Label22.Text = "Life of  Tool (Months)"
        '
        'ProdQty
        '
        Me.ProdQty.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdQty.Location = New System.Drawing.Point(402, 60)
        Me.ProdQty.Name = "ProdQty"
        Me.ProdQty.Size = New System.Drawing.Size(78, 21)
        Me.ProdQty.TabIndex = 27
        '
        'ProtoQty
        '
        Me.ProtoQty.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProtoQty.Location = New System.Drawing.Point(402, 29)
        Me.ProtoQty.Name = "ProtoQty"
        Me.ProtoQty.Size = New System.Drawing.Size(78, 21)
        Me.ProtoQty.TabIndex = 22
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Blue
        Me.Label21.Location = New System.Drawing.Point(407, 12)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(25, 14)
        Me.Label21.TabIndex = 11
        Me.Label21.Text = "Qty"
        '
        'ProdLeadTime
        '
        Me.ProdLeadTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdLeadTime.Location = New System.Drawing.Point(297, 60)
        Me.ProdLeadTime.Name = "ProdLeadTime"
        Me.ProdLeadTime.Size = New System.Drawing.Size(83, 21)
        Me.ProdLeadTime.TabIndex = 26
        '
        'ProtoLeadTime
        '
        Me.ProtoLeadTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProtoLeadTime.Location = New System.Drawing.Point(297, 28)
        Me.ProtoLeadTime.Name = "ProtoLeadTime"
        Me.ProtoLeadTime.Size = New System.Drawing.Size(83, 21)
        Me.ProtoLeadTime.TabIndex = 21
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Blue
        Me.Label20.Location = New System.Drawing.Point(293, 11)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(102, 14)
        Me.Label20.TabIndex = 8
        Me.Label20.Text = "Lead Time (Days)"
        '
        'ProdCustShare
        '
        Me.ProdCustShare.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdCustShare.Location = New System.Drawing.Point(198, 61)
        Me.ProdCustShare.Name = "ProdCustShare"
        Me.ProdCustShare.Size = New System.Drawing.Size(84, 21)
        Me.ProdCustShare.TabIndex = 25
        '
        'ProtoCustShare
        '
        Me.ProtoCustShare.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProtoCustShare.Location = New System.Drawing.Point(197, 29)
        Me.ProtoCustShare.Name = "ProtoCustShare"
        Me.ProtoCustShare.Size = New System.Drawing.Size(83, 21)
        Me.ProtoCustShare.TabIndex = 20
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Blue
        Me.Label13.Location = New System.Drawing.Point(197, 11)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(87, 14)
        Me.Label13.TabIndex = 5
        Me.Label13.Text = "Cust Cost(INR)"
        '
        'ProdTotalCost
        '
        Me.ProdTotalCost.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdTotalCost.Location = New System.Drawing.Point(107, 61)
        Me.ProdTotalCost.Name = "ProdTotalCost"
        Me.ProdTotalCost.Size = New System.Drawing.Size(80, 21)
        Me.ProdTotalCost.TabIndex = 24
        '
        'ProtoTotal
        '
        Me.ProtoTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProtoTotal.Location = New System.Drawing.Point(107, 29)
        Me.ProtoTotal.Name = "ProtoTotal"
        Me.ProtoTotal.Size = New System.Drawing.Size(80, 21)
        Me.ProtoTotal.TabIndex = 19
        '
        'lblTotalCost
        '
        Me.lblTotalCost.AutoSize = True
        Me.lblTotalCost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalCost.ForeColor = System.Drawing.Color.Blue
        Me.lblTotalCost.Location = New System.Drawing.Point(103, 11)
        Me.lblTotalCost.Name = "lblTotalCost"
        Me.lblTotalCost.Size = New System.Drawing.Size(90, 14)
        Me.lblTotalCost.TabIndex = 2
        Me.lblTotalCost.Text = "Total Cost (INR)"
        '
        'txtLatestAction
        '
        Me.txtLatestAction.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtLatestAction.ForeColor = System.Drawing.Color.Fuchsia
        Me.txtLatestAction.Location = New System.Drawing.Point(10, 102)
        Me.txtLatestAction.Name = "txtLatestAction"
        Me.txtLatestAction.Size = New System.Drawing.Size(102, 21)
        Me.txtLatestAction.TabIndex = 223
        Me.txtLatestAction.Tag = ""
        Me.txtLatestAction.Text = "Latest Action By:"
        '
        'DataGridCertificateCharges
        '
        Me.DataGridCertificateCharges.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridCertificateCharges.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCertificateCharges.CaptionForeColor = System.Drawing.Color.Black
        Me.DataGridCertificateCharges.CaptionVisible = False
        Me.DataGridCertificateCharges.DataMember = ""
        Me.DataGridCertificateCharges.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCertificateCharges.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridCertificateCharges.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridCertificateCharges.Location = New System.Drawing.Point(668, 110)
        Me.DataGridCertificateCharges.Name = "DataGridCertificateCharges"
        Me.DataGridCertificateCharges.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridCertificateCharges.ParentRowsVisible = False
        Me.DataGridCertificateCharges.PreferredColumnWidth = 85
        Me.DataGridCertificateCharges.RowHeadersVisible = False
        Me.DataGridCertificateCharges.RowHeaderWidth = 20
        Me.DataGridCertificateCharges.Size = New System.Drawing.Size(260, 99)
        Me.DataGridCertificateCharges.TabIndex = 38
        '
        'LatestAction
        '
        Me.LatestAction.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.LatestAction.ForeColor = System.Drawing.Color.Fuchsia
        Me.LatestAction.Location = New System.Drawing.Point(108, 102)
        Me.LatestAction.Name = "LatestAction"
        Me.LatestAction.Size = New System.Drawing.Size(60, 21)
        Me.LatestAction.TabIndex = 215
        '
        'txtMov
        '
        Me.txtMov.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMov.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMov.Location = New System.Drawing.Point(178, 137)
        Me.txtMov.MaxLength = 20
        Me.txtMov.Name = "txtMov"
        Me.txtMov.Size = New System.Drawing.Size(59, 20)
        Me.txtMov.TabIndex = 10
        '
        'lblMov
        '
        Me.lblMov.BackColor = System.Drawing.Color.Transparent
        Me.lblMov.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMov.ForeColor = System.Drawing.Color.Red
        Me.lblMov.Location = New System.Drawing.Point(182, 122)
        Me.lblMov.Name = "lblMov"
        Me.lblMov.Size = New System.Drawing.Size(36, 14)
        Me.lblMov.TabIndex = 222
        Me.lblMov.Text = "MOV"
        '
        'lblForwardApl
        '
        Me.lblForwardApl.BackColor = System.Drawing.Color.Transparent
        Me.lblForwardApl.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblForwardApl.ForeColor = System.Drawing.Color.Blue
        Me.lblForwardApl.Location = New System.Drawing.Point(438, 272)
        Me.lblForwardApl.Name = "lblForwardApl"
        Me.lblForwardApl.Size = New System.Drawing.Size(124, 14)
        Me.lblForwardApl.TabIndex = 220
        Me.lblForwardApl.Text = "Reason for forwarding"
        Me.lblForwardApl.Visible = False
        '
        'txtreasonforwarding
        '
        Me.txtreasonforwarding.BackColor = System.Drawing.Color.White
        Me.txtreasonforwarding.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtreasonforwarding.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtreasonforwarding.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.txtreasonforwarding.Location = New System.Drawing.Point(567, 270)
        Me.txtreasonforwarding.MaxLength = 100
        Me.txtreasonforwarding.Name = "txtreasonforwarding"
        Me.txtreasonforwarding.Size = New System.Drawing.Size(450, 20)
        Me.txtreasonforwarding.TabIndex = 36
        Me.txtreasonforwarding.Visible = False
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(8, 254)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 16)
        Me.Label7.TabIndex = 217
        Me.Label7.Text = "Special Notes if any"
        '
        'txtPurSpecial
        '
        Me.txtPurSpecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPurSpecial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPurSpecial.Location = New System.Drawing.Point(8, 270)
        Me.txtPurSpecial.MaxLength = 100
        Me.txtPurSpecial.Name = "txtPurSpecial"
        Me.txtPurSpecial.Size = New System.Drawing.Size(425, 20)
        Me.txtPurSpecial.TabIndex = 35
        '
        'ToolCost
        '
        Me.ToolCost.Controls.Add(Me.ToolFrameOpen)
        Me.ToolCost.Controls.Add(Me.RBToolNo)
        Me.ToolCost.Controls.Add(Me.RBToolYes)
        Me.ToolCost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolCost.Location = New System.Drawing.Point(583, 122)
        Me.ToolCost.Name = "ToolCost"
        Me.ToolCost.Size = New System.Drawing.Size(82, 87)
        Me.ToolCost.TabIndex = 215
        Me.ToolCost.TabStop = False
        Me.ToolCost.Text = "Tool Cost"
        '
        'ToolFrameOpen
        '
        Me.ToolFrameOpen.AutoSize = True
        Me.ToolFrameOpen.Font = New System.Drawing.Font("Arial", 6.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolFrameOpen.Location = New System.Drawing.Point(7, 60)
        Me.ToolFrameOpen.Name = "ToolFrameOpen"
        Me.ToolFrameOpen.Size = New System.Drawing.Size(44, 14)
        Me.ToolFrameOpen.TabIndex = 17
        Me.ToolFrameOpen.Text = "Open"
        Me.ToolFrameOpen.UseVisualStyleBackColor = True
        '
        'RBToolNo
        '
        Me.RBToolNo.AutoSize = True
        Me.RBToolNo.ForeColor = System.Drawing.Color.Blue
        Me.RBToolNo.Location = New System.Drawing.Point(7, 35)
        Me.RBToolNo.Name = "RBToolNo"
        Me.RBToolNo.Size = New System.Drawing.Size(39, 18)
        Me.RBToolNo.TabIndex = 16
        Me.RBToolNo.TabStop = True
        Me.RBToolNo.Text = "No"
        Me.RBToolNo.UseVisualStyleBackColor = True
        '
        'RBToolYes
        '
        Me.RBToolYes.AutoSize = True
        Me.RBToolYes.ForeColor = System.Drawing.Color.Blue
        Me.RBToolYes.Location = New System.Drawing.Point(7, 16)
        Me.RBToolYes.Name = "RBToolYes"
        Me.RBToolYes.Size = New System.Drawing.Size(45, 18)
        Me.RBToolYes.TabIndex = 15
        Me.RBToolYes.TabStop = True
        Me.RBToolYes.Text = "Yes"
        Me.RBToolYes.UseVisualStyleBackColor = True
        '
        'btnnext
        '
        Me.btnnext.BackColor = System.Drawing.Color.Silver
        Me.btnnext.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnnext.ForeColor = System.Drawing.Color.Black
        Me.btnnext.Location = New System.Drawing.Point(1258, 102)
        Me.btnnext.Name = "btnnext"
        Me.btnnext.Size = New System.Drawing.Size(14, 24)
        Me.btnnext.TabIndex = 210
        Me.btnnext.Text = "Next Price"
        Me.btnnext.UseVisualStyleBackColor = False
        Me.btnnext.Visible = False
        '
        'txtitemstatus
        '
        Me.txtitemstatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtitemstatus.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtitemstatus.Enabled = False
        Me.txtitemstatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtitemstatus.Location = New System.Drawing.Point(1217, 0)
        Me.txtitemstatus.MaxLength = 50
        Me.txtitemstatus.Name = "txtitemstatus"
        Me.txtitemstatus.Size = New System.Drawing.Size(31, 20)
        Me.txtitemstatus.TabIndex = 209
        Me.txtitemstatus.Visible = False
        '
        'txtRFQIntcode
        '
        Me.txtRFQIntcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRFQIntcode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRFQIntcode.Location = New System.Drawing.Point(407, 103)
        Me.txtRFQIntcode.MaxLength = 20
        Me.txtRFQIntcode.Name = "txtRFQIntcode"
        Me.txtRFQIntcode.Size = New System.Drawing.Size(60, 20)
        Me.txtRFQIntcode.TabIndex = 208
        Me.txtRFQIntcode.Visible = False
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Red
        Me.Label18.Location = New System.Drawing.Point(787, 212)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(103, 16)
        Me.Label18.TabIndex = 206
        Me.Label18.Text = "Progress Status"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Red
        Me.Label19.Location = New System.Drawing.Point(607, 212)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(185, 16)
        Me.Label19.TabIndex = 205
        Me.Label19.Text = "Vendor Quote Reference"
        '
        'txtvendquote
        '
        Me.txtvendquote.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvendquote.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtvendquote.Location = New System.Drawing.Point(607, 231)
        Me.txtvendquote.MaxLength = 100
        Me.txtvendquote.Name = "txtvendquote"
        Me.txtvendquote.Size = New System.Drawing.Size(176, 20)
        Me.txtvendquote.TabIndex = 32
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Red
        Me.Label11.Location = New System.Drawing.Point(373, 212)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(154, 16)
        Me.Label11.TabIndex = 203
        Me.Label11.Text = "Contact person / Designation"
        '
        'txtvendorref
        '
        Me.txtvendorref.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvendorref.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtvendorref.Location = New System.Drawing.Point(128, 231)
        Me.txtvendorref.MaxLength = 100
        Me.txtvendorref.Name = "txtvendorref"
        Me.txtvendorref.Size = New System.Drawing.Size(55, 20)
        Me.txtvendorref.TabIndex = 30
        '
        'txtvendorcontact
        '
        Me.txtvendorcontact.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvendorcontact.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtvendorcontact.Location = New System.Drawing.Point(368, 231)
        Me.txtvendorcontact.MaxLength = 100
        Me.txtvendorcontact.Name = "txtvendorcontact"
        Me.txtvendorcontact.Size = New System.Drawing.Size(234, 20)
        Me.txtvendorcontact.TabIndex = 31
        '
        'txtprogressdetails
        '
        Me.txtprogressdetails.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtprogressdetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtprogressdetails.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtprogressdetails.ForeColor = System.Drawing.Color.Fuchsia
        Me.txtprogressdetails.Location = New System.Drawing.Point(788, 231)
        Me.txtprogressdetails.MaxLength = 100
        Me.txtprogressdetails.Name = "txtprogressdetails"
        Me.txtprogressdetails.Size = New System.Drawing.Size(429, 20)
        Me.txtprogressdetails.TabIndex = 33
        '
        'l
        '
        Me.l.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.l.Location = New System.Drawing.Point(0, 96)
        Me.l.Name = "l"
        Me.l.Size = New System.Drawing.Size(1280, 8)
        Me.l.TabIndex = 197
        Me.l.TabStop = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Blue
        Me.Label5.Location = New System.Drawing.Point(1018, 254)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 196
        Me.Label5.Text = "Status"
        '
        'ComboBoxPriceStatus
        '
        Me.ComboBoxPriceStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPriceStatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxPriceStatus.ItemHeight = 14
        Me.ComboBoxPriceStatus.Items.AddRange(New Object() {"Accepted-Line Open", "Accepted-Line Closed", "Rejected", "Released to Customer Sup"})
        Me.ComboBoxPriceStatus.Location = New System.Drawing.Point(1022, 269)
        Me.ComboBoxPriceStatus.Name = "ComboBoxPriceStatus"
        Me.ComboBoxPriceStatus.Size = New System.Drawing.Size(191, 22)
        Me.ComboBoxPriceStatus.TabIndex = 20
        '
        'DataGridQty
        '
        Me.DataGridQty.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridQty.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQty.CaptionForeColor = System.Drawing.Color.Black
        Me.DataGridQty.CaptionVisible = False
        Me.DataGridQty.DataMember = ""
        Me.DataGridQty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQty.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQty.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridQty.Location = New System.Drawing.Point(933, 107)
        Me.DataGridQty.Name = "DataGridQty"
        Me.DataGridQty.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridQty.ParentRowsVisible = False
        Me.DataGridQty.PreferredColumnWidth = 59
        Me.DataGridQty.RowHeadersVisible = False
        Me.DataGridQty.RowHeaderWidth = 20
        Me.DataGridQty.Size = New System.Drawing.Size(319, 102)
        Me.DataGridQty.TabIndex = 39
        '
        'BtnAddPrice
        '
        Me.BtnAddPrice.BackColor = System.Drawing.Color.Silver
        Me.BtnAddPrice.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddPrice.ForeColor = System.Drawing.Color.Black
        Me.BtnAddPrice.Location = New System.Drawing.Point(1258, 131)
        Me.BtnAddPrice.Name = "BtnAddPrice"
        Me.BtnAddPrice.Size = New System.Drawing.Size(14, 24)
        Me.BtnAddPrice.TabIndex = 193
        Me.BtnAddPrice.Text = "Add another Price"
        Me.BtnAddPrice.UseVisualStyleBackColor = False
        Me.BtnAddPrice.Visible = False
        '
        'ComboBoxuom
        '
        Me.ComboBoxuom.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxuom.Items.AddRange(New Object() {"EA", "Pcs", "Sets", "Mtrs", "Cms", "Ltrs", "Ft", "Sheet", "Length"})
        Me.ComboBoxuom.Location = New System.Drawing.Point(1023, 32)
        Me.ComboBoxuom.Name = "ComboBoxuom"
        Me.ComboBoxuom.Size = New System.Drawing.Size(65, 22)
        Me.ComboBoxuom.TabIndex = 166
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(53, 17)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(74, 16)
        Me.Label10.TabIndex = 183
        Me.Label10.Text = "Avbl FS"
        '
        'Label54
        '
        Me.Label54.BackColor = System.Drawing.Color.Transparent
        Me.Label54.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label54.ForeColor = System.Drawing.Color.Black
        Me.Label54.Location = New System.Drawing.Point(128, 17)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(55, 16)
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
        Me.Label46.Size = New System.Drawing.Size(105, 16)
        Me.Label46.TabIndex = 181
        Me.Label46.Text = "Recom. Vendor"
        '
        'txtRecVend
        '
        Me.txtRecVend.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRecVend.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRecVend.Location = New System.Drawing.Point(1093, 33)
        Me.txtRecVend.MaxLength = 50
        Me.txtRecVend.Name = "txtRecVend"
        Me.txtRecVend.Size = New System.Drawing.Size(169, 20)
        Me.txtRecVend.TabIndex = 167
        '
        'Label45
        '
        Me.Label45.BackColor = System.Drawing.Color.Transparent
        Me.Label45.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label45.ForeColor = System.Drawing.Color.Black
        Me.Label45.Location = New System.Drawing.Point(1023, 16)
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
        Me.Label15.Location = New System.Drawing.Point(617, 56)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(120, 16)
        Me.Label15.TabIndex = 179
        Me.Label15.Text = "Special Instructions"
        '
        'Label42
        '
        Me.Label42.BackColor = System.Drawing.Color.Transparent
        Me.Label42.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label42.ForeColor = System.Drawing.Color.Black
        Me.Label42.Location = New System.Drawing.Point(192, 56)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(88, 16)
        Me.Label42.TabIndex = 178
        Me.Label42.Text = "Material"
        '
        'txtDimension
        '
        Me.txtDimension.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDimension.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDimension.Location = New System.Drawing.Point(8, 72)
        Me.txtDimension.Name = "txtDimension"
        Me.txtDimension.Size = New System.Drawing.Size(175, 20)
        Me.txtDimension.TabIndex = 168
        '
        'txtCustDesc
        '
        Me.txtCustDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustDesc.Location = New System.Drawing.Point(777, 32)
        Me.txtCustDesc.MaxLength = 80
        Me.txtCustDesc.Name = "txtCustDesc"
        Me.txtCustDesc.Size = New System.Drawing.Size(240, 20)
        Me.txtCustDesc.TabIndex = 165
        '
        'ComboBoxFSYesNo
        '
        Me.ComboBoxFSYesNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxFSYesNo.Items.AddRange(New Object() {"Yes", "No"})
        Me.ComboBoxFSYesNo.Location = New System.Drawing.Point(48, 32)
        Me.ComboBoxFSYesNo.Name = "ComboBoxFSYesNo"
        Me.ComboBoxFSYesNo.Size = New System.Drawing.Size(64, 22)
        Me.ComboBoxFSYesNo.TabIndex = 160
        '
        'Label41
        '
        Me.Label41.BackColor = System.Drawing.Color.Transparent
        Me.Label41.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label41.ForeColor = System.Drawing.Color.Black
        Me.Label41.Location = New System.Drawing.Point(8, 16)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(40, 16)
        Me.Label41.TabIndex = 177
        Me.Label41.Text = "Sl.No."
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(4, Byte), Integer))
        Me.Label40.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.ForeColor = System.Drawing.Color.Black
        Me.Label40.Location = New System.Drawing.Point(8, 56)
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
        Me.Label39.Location = New System.Drawing.Point(592, 13)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(115, 19)
        Me.Label39.TabIndex = 175
        Me.Label39.Text = "Customer Part No."
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.Transparent
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(777, 16)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(151, 16)
        Me.Label16.TabIndex = 174
        Me.Label16.Text = "Cust Part Description"
        '
        'Label37
        '
        Me.Label37.BackColor = System.Drawing.Color.Transparent
        Me.Label37.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.ForeColor = System.Drawing.Color.Black
        Me.Label37.Location = New System.Drawing.Point(368, 16)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(135, 16)
        Me.Label37.TabIndex = 173
        Me.Label37.Text = "Description"
        '
        'Label36
        '
        Me.Label36.BackColor = System.Drawing.Color.Transparent
        Me.Label36.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.ForeColor = System.Drawing.Color.Black
        Me.Label36.Location = New System.Drawing.Point(208, 16)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(64, 16)
        Me.Label36.TabIndex = 172
        Me.Label36.Text = "Part No."
        '
        'txtpart
        '
        Me.txtpart.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpart.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtpart.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpart.Location = New System.Drawing.Point(208, 32)
        Me.txtpart.MaxLength = 50
        Me.txtpart.Name = "txtpart"
        Me.txtpart.Size = New System.Drawing.Size(152, 20)
        Me.txtpart.TabIndex = 162
        '
        'txtCustPart
        '
        Me.txtCustPart.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustPart.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustPart.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustPart.Location = New System.Drawing.Point(552, 32)
        Me.txtCustPart.MaxLength = 50
        Me.txtCustPart.Name = "txtCustPart"
        Me.txtCustPart.Size = New System.Drawing.Size(216, 20)
        Me.txtCustPart.TabIndex = 164
        '
        'txtPartDesc
        '
        Me.txtPartDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartDesc.Location = New System.Drawing.Point(368, 32)
        Me.txtPartDesc.MaxLength = 50
        Me.txtPartDesc.Name = "txtPartDesc"
        Me.txtPartDesc.Size = New System.Drawing.Size(175, 20)
        Me.txtPartDesc.TabIndex = 163
        '
        'txtslno
        '
        Me.txtslno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtslno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtslno.ForeColor = System.Drawing.Color.Black
        Me.txtslno.Location = New System.Drawing.Point(8, 32)
        Me.txtslno.Name = "txtslno"
        Me.txtslno.ReadOnly = True
        Me.txtslno.Size = New System.Drawing.Size(32, 20)
        Me.txtslno.TabIndex = 159
        '
        'txtMaterial
        '
        Me.txtMaterial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMaterial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterial.Location = New System.Drawing.Point(192, 72)
        Me.txtMaterial.MaxLength = 80
        Me.txtMaterial.Name = "txtMaterial"
        Me.txtMaterial.Size = New System.Drawing.Size(416, 20)
        Me.txtMaterial.TabIndex = 169
        '
        'txtDetailSpecial
        '
        Me.txtDetailSpecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDetailSpecial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDetailSpecial.Location = New System.Drawing.Point(617, 72)
        Me.txtDetailSpecial.MaxLength = 100
        Me.txtDetailSpecial.Name = "txtDetailSpecial"
        Me.txtDetailSpecial.Size = New System.Drawing.Size(640, 20)
        Me.txtDetailSpecial.TabIndex = 170
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Blue
        Me.Label14.Location = New System.Drawing.Point(463, 122)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(80, 13)
        Me.Label14.TabIndex = 157
        Me.Label14.Text = "Stock Avlbe"
        '
        'txtstockavble
        '
        Me.txtstockavble.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtstockavble.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtstockavble.Location = New System.Drawing.Point(467, 137)
        Me.txtstockavble.MaxLength = 100
        Me.txtstockavble.Name = "txtstockavble"
        Me.txtstockavble.Size = New System.Drawing.Size(116, 20)
        Me.txtstockavble.TabIndex = 14
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Red
        Me.Label9.Location = New System.Drawing.Point(373, 122)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(84, 13)
        Me.Label9.TabIndex = 150
        Me.Label9.Text = "Vend Currency"
        '
        'comboboxcurrency
        '
        Me.comboboxcurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.comboboxcurrency.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comboboxcurrency.Items.AddRange(New Object() {"INR", "EUR", "USD", "SGD", "SEK", "JPY", "GBP", "DKK", "CAD", "CHF"})
        Me.comboboxcurrency.Location = New System.Drawing.Point(372, 135)
        Me.comboboxcurrency.Name = "comboboxcurrency"
        Me.comboboxcurrency.Size = New System.Drawing.Size(80, 22)
        Me.comboboxcurrency.TabIndex = 13
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Blue
        Me.Label8.Location = New System.Drawing.Point(312, 123)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 13)
        Me.Label8.TabIndex = 148
        Me.Label8.Text = "Type"
        '
        'comboboxstocktype
        '
        Me.comboboxstocktype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.comboboxstocktype.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comboboxstocktype.Items.AddRange(New Object() {"MTO", "MTC", "MTS", "None"})
        Me.comboboxstocktype.Location = New System.Drawing.Point(310, 136)
        Me.comboboxstocktype.Name = "comboboxstocktype"
        Me.comboboxstocktype.Size = New System.Drawing.Size(58, 22)
        Me.comboboxstocktype.TabIndex = 12
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Blue
        Me.Label6.Location = New System.Drawing.Point(253, 123)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(27, 13)
        Me.Label6.TabIndex = 146
        Me.Label6.Text = "SPU"
        '
        'txtSPU
        '
        Me.txtSPU.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSPU.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSPU.Location = New System.Drawing.Point(242, 138)
        Me.txtSPU.MaxLength = 20
        Me.txtSPU.Name = "txtSPU"
        Me.txtSPU.Size = New System.Drawing.Size(61, 20)
        Me.txtSPU.TabIndex = 11
        '
        'txtMOQ
        '
        Me.txtMOQ.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMOQ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMOQ.Location = New System.Drawing.Point(128, 136)
        Me.txtMOQ.MaxLength = 20
        Me.txtMOQ.Name = "txtMOQ"
        Me.txtMOQ.Size = New System.Drawing.Size(49, 20)
        Me.txtMOQ.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Blue
        Me.Label4.Location = New System.Drawing.Point(140, 122)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 141
        Me.Label4.Text = "MOQ"
        '
        'Label43
        '
        Me.Label43.BackColor = System.Drawing.Color.Transparent
        Me.Label43.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label43.ForeColor = System.Drawing.Color.Blue
        Me.Label43.Location = New System.Drawing.Point(7, 159)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(141, 23)
        Me.Label43.TabIndex = 138
        Me.Label43.Text = "Special Instructions for CS"
        '
        'txtLeadTime
        '
        Me.txtLeadTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLeadTime.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLeadTime.Location = New System.Drawing.Point(7, 136)
        Me.txtLeadTime.MaxLength = 100
        Me.txtLeadTime.Name = "txtLeadTime"
        Me.txtLeadTime.Size = New System.Drawing.Size(115, 20)
        Me.txtLeadTime.TabIndex = 8
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.Color.Transparent
        Me.Label38.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.ForeColor = System.Drawing.Color.Blue
        Me.Label38.Location = New System.Drawing.Point(7, 122)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(141, 20)
        Me.Label38.TabIndex = 129
        Me.Label38.Text = "Product Lead Time(days)"
        '
        'BtnDelete
        '
        Me.BtnDelete.BackColor = System.Drawing.Color.Silver
        Me.BtnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDelete.ForeColor = System.Drawing.Color.Black
        Me.BtnDelete.Location = New System.Drawing.Point(1222, 267)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(48, 24)
        Me.BtnDelete.TabIndex = 126
        Me.BtnDelete.Text = "Delete"
        Me.BtnDelete.UseVisualStyleBackColor = False
        Me.BtnDelete.Visible = False
        '
        'btnRFQSave
        '
        Me.btnRFQSave.BackColor = System.Drawing.Color.Silver
        Me.btnRFQSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRFQSave.ForeColor = System.Drawing.Color.Black
        Me.btnRFQSave.Location = New System.Drawing.Point(1222, 242)
        Me.btnRFQSave.Name = "btnRFQSave"
        Me.btnRFQSave.Size = New System.Drawing.Size(48, 24)
        Me.btnRFQSave.TabIndex = 37
        Me.btnRFQSave.Text = "Save"
        Me.btnRFQSave.UseVisualStyleBackColor = False
        '
        'txtdetailremarks
        '
        Me.txtdetailremarks.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtdetailremarks.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdetailremarks.Location = New System.Drawing.Point(155, 159)
        Me.txtdetailremarks.MaxLength = 100
        Me.txtdetailremarks.Multiline = True
        Me.txtdetailremarks.Name = "txtdetailremarks"
        Me.txtdetailremarks.Size = New System.Drawing.Size(247, 26)
        Me.txtdetailremarks.TabIndex = 29
        '
        'GroupBox7
        '
        Me.GroupBox7.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox7.Controls.Add(Me.Label26)
        Me.GroupBox7.Controls.Add(Me.Label17)
        Me.GroupBox7.Controls.Add(Me.BtnQuoteRefSave)
        Me.GroupBox7.Controls.Add(Me.Label12)
        Me.GroupBox7.Controls.Add(Me.txtQuoteRef1)
        Me.GroupBox7.Controls.Add(Me.txtQuoteRef2)
        Me.GroupBox7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox7.ForeColor = System.Drawing.Color.Red
        Me.GroupBox7.Location = New System.Drawing.Point(8, 292)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(1259, 42)
        Me.GroupBox7.TabIndex = 240
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Vendor Quote References"
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.Transparent
        Me.Label26.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Red
        Me.Label26.Location = New System.Drawing.Point(830, 13)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(28, 26)
        Me.Label26.TabIndex = 236
        Me.Label26.Text = "3"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.Transparent
        Me.Label17.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Red
        Me.Label17.Location = New System.Drawing.Point(417, 16)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(18, 26)
        Me.Label17.TabIndex = 235
        Me.Label17.Text = "2"
        '
        'BtnQuoteRefSave
        '
        Me.BtnQuoteRefSave.BackColor = System.Drawing.Color.Silver
        Me.BtnQuoteRefSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnQuoteRefSave.ForeColor = System.Drawing.Color.Black
        Me.BtnQuoteRefSave.Location = New System.Drawing.Point(1177, 8)
        Me.BtnQuoteRefSave.Name = "BtnQuoteRefSave"
        Me.BtnQuoteRefSave.Size = New System.Drawing.Size(80, 24)
        Me.BtnQuoteRefSave.TabIndex = 237
        Me.BtnQuoteRefSave.Text = "Q Ref Save"
        Me.BtnQuoteRefSave.UseVisualStyleBackColor = False
        Me.BtnQuoteRefSave.Visible = False
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Red
        Me.Label12.Location = New System.Drawing.Point(17, 16)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(23, 20)
        Me.Label12.TabIndex = 234
        Me.Label12.Text = "1"
        '
        'txtQuoteRef1
        '
        Me.txtQuoteRef1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtQuoteRef1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuoteRef1.Location = New System.Drawing.Point(47, 15)
        Me.txtQuoteRef1.MaxLength = 100
        Me.txtQuoteRef1.Multiline = True
        Me.txtQuoteRef1.Name = "txtQuoteRef1"
        Me.txtQuoteRef1.Size = New System.Drawing.Size(365, 21)
        Me.txtQuoteRef1.TabIndex = 21
        '
        'txtQuoteRef2
        '
        Me.txtQuoteRef2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtQuoteRef2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuoteRef2.Location = New System.Drawing.Point(442, 16)
        Me.txtQuoteRef2.MaxLength = 100
        Me.txtQuoteRef2.Multiline = True
        Me.txtQuoteRef2.Name = "txtQuoteRef2"
        Me.txtQuoteRef2.Size = New System.Drawing.Size(381, 20)
        Me.txtQuoteRef2.TabIndex = 22
        '
        'RFQ
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.CornflowerBlue
        Me.ClientSize = New System.Drawing.Size(1572, 852)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.DataUpdation)
        Me.Name = "RFQ"
        Me.Text = "Pricing"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        CType(Me.DatagridMultiprices, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.datagridEnquiryPending, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DataUpdation.ResumeLayout(False)
        Me.DataUpdation.PerformLayout()
        CType(Me.datagridEnquiryPending1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridVendor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.Vendor.ResumeLayout(False)
        Me.ToolDetails.ResumeLayout(False)
        Me.ToolDetails.PerformLayout()
        CType(Me.DataGridCertificateCharges, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolCost.ResumeLayout(False)
        Me.ToolCost.PerformLayout()
        CType(Me.DataGridQty, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If screentype = "COMP" Or screentype = "COMPP" Then
            GroupBox2.Text = "RFQ Completed"
        End If

        If screentype = "PEND" Or screentype = "PENDP" Then

            datagridEnquiryPending.Enabled = True

            RBToolNo.Checked = True
            RadioButtonGroup.Checked = True
            RadioButtonVendorYes.Checked = True


            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
            'Dim cmSQL As SqlCommand
            'Dim drSQL As SqlDataReader
            Dim strSQL As String


            Dim stockDC As DataSet = New DataSet

            If screentype = "PEND" Then

                GroupBox2.Text = "Request for Quotation"


                btnRFQSave.Visible = True



                strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, UserId, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                      "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1,Enq_Type, Enq_Forward,Status,Enq_Int_code " & _
                       "from TSS_Enquiry_Pending_Price order by RegNo,SlNo"




            ElseIf screentype = "PENDP" Then

                GroupBox2.Text = "Request for Quotation"
                btnRFQSave.Visible = True

                strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, UserId, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                         "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1,Enq_Type, Enq_Forward,Enq_Int_code " & _
                         "from TSS_Enquiry_Pending_Price where Enq_Type in ('Project','Project-Budgetary') and Enq_Forward = 'Forward to Apl. Dept' AND [Reg.Date] >= '11-01-2013' order by RegNo,SlNo"


            ElseIf screentype = "COMP" Then

                GroupBox2.Text = "RFQ Completed"
                btnRFQSave.Visible = False
                BtnDelete.Visible = False
                BtnQuoteRefSave.Visible = True

                'btnRFQSave.Visible = True
                'btnRFQSave.Name = "QRef Save"

                BtnDelete.Visible = False
                '            strSQL = "Select * from TSS_Enquiry_Price_Completed order by RegNo,SlNo"

                strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                         "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1 from TSS_Enquiry_Price_Completed order by RegNo,SlNo"

            ElseIf screentype = "COMPP" Then

                GroupBox2.Text = "RFQ Completed"

                btnRFQSave.Visible = False
                BtnDelete.Visible = False
                'btnRFQSave.Text = "QRefSave"

                strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
               "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1 from TSS_Enquiry_Price_Completed " & _
               "where Enq_Type in ('Project','Project-Budgetary') and Enq_Forward = 'Forward to Apl. Dept' AND [Reg.Date] >= '11-01-2013' order by RegNo,SlNo"



            End If


            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAC.TableMappings.Add("Table", "Enq")
            'get data
            stockDAC.Fill(stockDC)



            datagridEnquiryPending.DataSource = stockDC.Tables(0)
            'datagridEnquiryPending.Expand(-1)


            'colouring unattended enquiries

            'datagridEnquiryPending.Rows()
            ' Dim a As String


            If screentype = "PEND" Then

                For i As Integer = 0 To datagridEnquiryPending.RowCount - 2
                    '    a = ""
                    ' a = datagridEnquiryPending.CurrentRow.Cells("Status").Value.ToString

                    'a = datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString
                    'If IsDBNull(datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) Or Len(a) < 2 Then

                    '   If IsDBNull(datagridEnquiryPending.Rows(i).Cells("Status").ToString) Or Len(datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) < 2 Then


                    If (datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) <> "-" Then

                        ' datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Red
                        datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Black

                    ElseIf (datagridEnquiryPending.Rows(i).Cells("Enq_Type").Value.ToString) <> "Internal RFQ" Then
                        ' datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Black
                        datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Red

                    ElseIf (datagridEnquiryPending.Rows(i).Cells("Enq_Type").Value.ToString) = "Internal RFQ" Then
                        ' datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Black
                        datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Blue


                    End If


                Next

            End If



            'end of colouring


            cnSQL.Close()
            '        datagridEnquiryPending.Expand(-1)


1:

            listloadCertificate()

        End If


        '    Me.ReportViewer1.RefreshReport()
    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click

    End Sub

    Private Sub datagridStock_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub

    Private Sub datagridEnquieryPending_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles datagridEnquiryPending.CurrentCellChanged

        'txtRFQIntcode.Text = ""
        'RadioButtonGroup.Checked = True
        'comboboxcurrency.Text = "EUR"
        'RBToolNo.Checked = True

        'If DatagridMultiprices.Visible = True Then
        '    DatagridMultiprices.Visible = False
        'End If

        'rfqmode = ""
        'multiple = "NO"

        'Dim b As Integer
        ''Dim custid As String
        ''b = datagridEnquiryPending.CurrentCell.ColumnNumber()

        'If b = 0 Then
        '    clearpricedetails()

        '    txtenqdetailintcode.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell)


        '    txtRegNo.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 1)
        '    DtpEnqRegDt.Value = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 2)
        '    txtCustomer.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 9)
        '    txtCity.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 30)
        '    TXTCL3.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 11)
        '    TXTCL1.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 31)
        '    txtCSR.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 12)
        '    txtISR.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 13)
        '    txtTSSISeg.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 14)
        '    txtTSSSeg.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 15)
        '    If datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 16) = "YES" Then
        '        RadioButtonExisting.Checked = True
        '        RadioButtonNew.Checked = False
        '    Else
        '        RadioButtonNew.Checked = True
        '        RadioButtonExisting.Checked = False
        '    End If
        '    If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 10)) Then
        '        RadioButtonDomestic.Checked = True
        '        RadioButtonExport.Checked = False
        '    ElseIf Trim(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 10)) = "Domestic" Then
        '        RadioButtonDomestic.Checked = True
        '        RadioButtonExport.Checked = False
        '    ElseIf Trim(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 10)) = "Export" Then
        '        RadioButtonDomestic.Checked = False
        '        RadioButtonExport.Checked = True
        '    End If
        '    If datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 17) = "YES" Then
        '        RadioButtondocyes.Checked = True
        '        RadioButtondocno.Checked = False
        '    Else
        '        RadioButtondocyes.Checked = False
        '        RadioButtondocno.Checked = True
        '    End If

        '    If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 18)) Then
        '        txtdocdetails.Text = ""
        '    Else

        '        txtdocdetails.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 18)
        '    End If

        '    If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 19)) Then
        '        txtSpecialInst.Text = ""
        '    Else

        '        txtSpecialInst.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 19)
        '    End If


        '    txtslno.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 3)
        '    ComboBoxFSYesNo.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 20)
        '    ComboBoxItemSource.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 21)
        '    txtpart.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 4)
        '    txtPartDesc.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 5)
        '    txtCustPart.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 6)
        '    txtCustDesc.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 22)
        '    ComboBoxuom.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 7)

        '    If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 23)) Then
        '        txtRecVend.Text = ""
        '    Else

        '        txtRecVend.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 23)
        '    End If


        '    If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 24)) Then
        '        txtDimension.Text = ""
        '    Else


        '        txtDimension.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 24)

        '    End If

        '    If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 25)) Then
        '        txtMaterial.Text = ""
        '    Else

        '        txtMaterial.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 25)

        '    End If

        '    If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 26)) Then
        '        txtDetailSpecial.Text = ""

        '    Else

        '        txtDetailSpecial.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 26)
        '    End If

        '    If datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 27) = "01-01-1900" Then
        '        dtpenqduedt.Checked = False

        '        dtpenqduedt.Value = "01-01-1900"
        '    Else
        '        dtpenqduedt.Checked = True

        '        dtpenqduedt.Value = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 27)

        '    End If


        '    txtitemstatus.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 28)

        '    txtitemstatus.Text = Trim(txtitemstatus.Text)

        '    'EditCertDetails()

        '    If txtitemstatus.Text = "H" Or txtitemstatus.Text = "U" Or txtitemstatus.Text = "C" Or txtitemstatus.Text = "F" Then
        '        callrfqdetails()

        '    End If

        'Else

        '    MsgBox("Click on Detailcode ", vbInformation)
        '    Exit Sub
        'End If

        'If multiple = "YES" Then
        '    Exit Sub
        'Else
        '    clearqty()
        '    fillqty()

        '    ClearCertificate()
        '    fillcertificate()


        'End If



    End Sub

    Private Sub DataUpdation_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataUpdation.Enter

    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Label38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label38.Click

    End Sub
    Private Sub listloadCertificate()


        'Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cmSQL1 As SqlCommand
        'Dim drSQL1 As SqlDataReader
        'Dim strSQL1 As String

        ''Dim a As ListView


        'cnSQL1.Open()
        'strSQL1 = "SELECT Certificates,Int_code FROM ENQ_Certificates " & _
        '         "WHERE  Status = 'A'"
        'cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        'drSQL1 = cmSQL1.ExecuteReader()

        'Dim ColumnValue As String = Nothing
        'Do While drSQL1.Read()

        '    ColumnValue = (drSQL1.GetValue(0)).ToString
        '    CheckedListBoxCertificate.Items.Add(ColumnValue)
        '    '  ListBoxCertificate.Sorted = True

        '    'ListBoxCertificate.DisplayMember = "Certificates"
        '    CheckedListBoxCertificate.ValueMember = "Int_code"

        'Loop
    End Sub
    Private Sub EditCertDetails()

        'Dim strsql As String
        'Dim cmSQL As SqlCommand
        'Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim drSQL1 As SqlDataReader

        'cnSQL.Open()

        'Dim i As Integer
        'Dim a As Integer
        'Dim cert As String
        ''Dim b As Integer

        'strsql = "Select Certificates from ENQ_EnqWise_Certificates where Enq_Reg_NO = " & txtRegNo.Text & " "

        'cmSQL = New SqlCommand(strsql, cnSQL)
        'drSQL1 = cmSQL.ExecuteReader()

        'i = CheckedListBoxCertificate.Items.Count

        'Do While drSQL1.Read()
        '    cert = drSQL1.Item(0)
        '    a = 0
        '    Do While a < i

        '        If cert = CheckedListBoxCertificate.Items(a) Then

        '            CheckedListBoxCertificate.SetItemChecked(a, True)

        '            a = i
        '        Else

        '            a = a + 1

        '        End If


        '    Loop

        'Loop

    End Sub

    Private Sub RadioButtonExisting_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBoxPriceStatus_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxPriceStatus.LostFocus

        If ComboBoxPriceStatus.Text = "Forwarded to App Dept" Then
            lblForwardApl.Visible = True
            txtreasonforwarding.Visible = True

        End If
    End Sub

    Private Sub ComboBoxPriceStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxPriceStatus.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxPriceStatus_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxPriceStatus.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtLeadTime_ReadOnlyChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeadTime.ReadOnlyChanged

    End Sub

    Private Sub TextBox24_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLeadTime.TextChanged

    End Sub

    Private Sub TextBox24_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLeadTime.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtMOQ_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMOQ.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMOQ.TextChanged

    End Sub

    Private Sub TextBox6_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMOQ.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtSPU_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSPU.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSPU.TextChanged

    End Sub

    Private Sub TextBox8_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSPU.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles comboboxstocktype.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxType_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles comboboxstocktype.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click

    End Sub

    Private Sub ComboBoxCurrency_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles comboboxcurrency.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxCurrency_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles comboboxcurrency.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtstockavble.TextChanged

    End Sub

    Private Sub TextBox10_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtstockavble.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox12_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtsupinvdt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtdetailremarks.TextChanged

    End Sub

    Private Sub txtsupinvdt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtdetailremarks.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtvendorref_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtvendorref.DoubleClick
        fillvendorlist()

    End Sub

    Private Sub txtvendorref_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvendorref.TextChanged

    End Sub

    Private Sub txtvendorref_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtvendorref.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtvendorcontact_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvendorcontact.TextChanged

    End Sub

    Private Sub txtvendorcontact_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtvendorcontact.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtvendquote_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvendquote.TextChanged

    End Sub

    Private Sub txtvendquote_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtvendquote.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtspecial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtprogressdetails.TextChanged

    End Sub

    Private Sub txtspecial_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtprogressdetails.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRFQSave.Click
        ' Dim a As Integer
        'a = 0
        'If a = 0 Then
        'sendmail()
        'Exit Sub
        'End If

        Dim msgb As String
        Dim strsql As String
        Dim toolyesno As String
        strsql = ""
        msgb = MsgBox("Have you entered all the details, including price ? Are you sure of saving ?", vbYesNo)

        If msgb = vbNo Then
            Exit Sub
        End If

        If ComboBoxPriceStatus.Text = "Released to Customer Sup" Then
            If Len(txtdetailremarks.Text) < 3 Then

                MsgBox("Special Instruction to customer support should not be blank", vbInformation)
                Exit Sub

            End If
        End If

        If screentype = "PENDP" Then

            If ComboBoxPriceStatus.Text = "Rejected" Or ComboBoxPriceStatus.Text = "Released to Customer Sup" Then
                MsgBox("This option [Rejected or Releasd to Customer Sup] is not available for project enquiries during pricing!")
                Exit Sub
            End If
        End If

        'validation
        If ComboBoxPriceStatus.Text = "Forwarded to App Dept" Then
            If Len(txtreasonforwarding.Text) < 2 Then
                MsgBox("Reason for forwarding should be entered", vbInformation)
                Exit Sub

            End If
        End If

        If RBToolYes.Checked = False And RBToolNo.Checked = False Then
            MsgBox("Toooling Charges  Yes or No should be selected ", vbCritical)
            Exit Sub
        End If

        If RBToolYes.Checked = True Then
            If Val(ProtoTotal.Text) + Val(ProdTotalCost.Text) = 0 Then
                MsgBox("If Tooling Charges Yes, then Tooling cost should not be zero", vbCritical)
                Exit Sub
            End If
        End If

        If ComboBoxPriceStatus.Text = "" Then
            MsgBox("Line status should not be blank", vbInformation)
            Exit Sub
        End If

        If ComboBoxPriceStatus.Text = "Accepted-Line Closed" Then
            msgb = MsgBox("This line will be closed permanently,  Are you sure of saving ?", vbYesNo)

            If msgb = vbNo Then
                Exit Sub
            End If

        End If

        If RBToolYes.Checked = True Then
            toolyesno = "YES"
        ElseIf RBToolNo.Checked = True Then
            toolyesno = "NO"
        End If

        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim cmSQL2 As SqlCommand
        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)



        curdate = System.DateTime.Now()

        Dim price As String
        price = ""

        If RadioButton3P.Checked = True Then
            price = "3rdParty"
        ElseIf RadioButtonGroup.Checked = True Then
            price = "Group"
        ElseIf RadioButtonFactory.Checked = True Then
            price = "Factory"
        End If
        If rfqmode <> "addprice" And multiple = "NO" Then
            If txtRFQIntcode.Text = "" Then
                RFQcodegen()
            End If

        End If

        Saveqtydetails()
        Savecertificatedetails()


        If ProtoTotal.Text = "" Then
            ProtoTotal.Text = 0

        End If

        If ProtoCustShare.Text = "" Then
            ProtoCustShare.Text = 0
        End If

        If ProtoLeadTime.Text = "" Then
            ProtoLeadTime.Text = 0
        End If

        If ProtoQty.Text = "" Then
            ProtoQty.Text = 0
        End If

        If ProtoLifeofTool.Text = "" Then
            ProtoLifeofTool.Text = 0
        End If

        If ProdTotalCost.Text = "" Then
            ProdTotalCost.Text = 0
        End If

        If ProdCustShare.Text = "" Then
            ProdCustShare.Text = 0
        End If

        If ProdLeadTime.Text = "" Then
            ProdLeadTime.Text = 0
        End If

        If ProdQty.Text = "" Then
            ProdQty.Text = 0
        End If

        If ProdLifeofTool.Text = "" Then
            ProdLifeofTool.Text = 0
        End If

        If txtMOQ.Text = "" Then
            txtMOQ.Text = 0
        End If

        If txtSPU.Text = "" Then
            txtSPU.Text = 0
        End If
        If txtMov.Text = "" Then
            txtMov.Text = 0
        End If

        If txtitemstatus.Text = "P" And rfqmode = "" And multiple = "NO" Then

            strsql = "insert ENQ_RFQ_PriceDetails values(" & txtenqdetailintcode.Text & "," & txtRFQIntcode.Text & "," & txtRegNo.Text & "," & _
                   "'" & ComboBoxPriceStatus.Text & "','" & price & "'," & txtMOQ.Text & "," & txtSPU.Text & ",'" & txtLeadTime.Text & "'," & _
                  "'" & comboboxstocktype.Text & "','" & comboboxcurrency.Text & "','" & txtstockavble.Text & "'," & _
                 "'" & txtdetailremarks.Text & "','" & txtvendorref.Text & "','" & txtvendorcontact.Text & "','" & txtvendquote.Text & "','" & txtPurSpecial.Text & "'," & _
                "'" & curdate & "','" & curdate & "','" & username & "'," & _
                " " & txtMov.Text & ", '" & toolyesno & "'," & ProtoTotal.Text & ", " & ProtoCustShare.Text & ", " & ProtoLeadTime.Text & ", " & _
                " " & ProtoQty.Text & ", " & ProtoLifeofTool.Text & "," & ProdTotalCost.Text & ", " & ProdCustShare.Text & ", " & ProdLeadTime.Text & ", " & _
                " " & ProdQty.Text & ", " & ProdLifeofTool.Text & ",'" & txtprogressdetails.Text & "', '" & txtreasonforwarding.Text & "','" & txtVendorName.Text & "'," & _
                " '" & txtQuoteRef1.Text & "', '" & txtQuoteRef2.Text & "', '" & txtQuoteRef3.Text & "', '" & txtAltMtrl.Text & "')"


        ElseIf (txtitemstatus.Text = "H" Or txtitemstatus.Text = "C" Or txtitemstatus.Text = "F") And rfqmode = "" And (multiple = "YES" Or multiple = "NO") Then
            strsql = "update ENQ_RFQ_PriceDetails  set Status = '" & ComboBoxPriceStatus.Text & "'," & _
             "Source_Mtrl = '" & price & "', MOQ = " & txtMOQ.Text & "," & _
             "SPU = " & txtSPU.Text & ", LeadTime = '" & txtLeadTime.Text & "'," & _
             "Type = '" & comboboxstocktype.Text & "',Currency = '" & comboboxcurrency.Text & "'," & _
             "Stock_Avble = '" & txtstockavble.Text & "'," & _
             "Remarks = '" & txtdetailremarks.Text & "',Vendor_Ref = '" & txtvendorref.Text & "'," & _
             "Name = '" & txtvendorcontact.Text & "',Vendor_Quote = '" & txtvendquote.Text & "'," & _
             "Special_Remarks = '" & txtPurSpecial.Text & "',Date_Modify = '" & curdate & "',UserId = '" & username & "'," & _
             "Mov = " & txtMov.Text & ", Tools_YesNo = '" & toolyesno & "',Proto_TotalCost = " & ProtoTotal.Text & "," & _
             "Proto_CustCost = " & ProtoCustShare.Text & ", Proto_LeadTime = " & ProtoLeadTime.Text & "," & _
             "Proto_Qty = " & ProtoQty.Text & ", Prod_TotalCost = " & ProdTotalCost.Text & ",Prod_CustCost = " & ProdCustShare.Text & "," & _
             "Prod_LeadTime = " & ProdLeadTime.Text & ", Prod_Qty = " & ProdQty.Text & ", Prod_LifeTool = " & ProdLifeofTool.Text & "," & _
             "Progress_details = '" & txtprogressdetails.Text & "',Forwarding_reason = '" & txtreasonforwarding.Text & "', Vendor_Name = '" & txtVendorName.Text & "'," & _
             "QuoteRef1 = '" & txtQuoteRef1.Text & "', QuoteRef2 = '" & txtQuoteRef2.Text & "', QuoteRef3 = '" & txtQuoteRef3.Text & "', AlternateMtrl = '" & txtAltMtrl.Text & "' " & _
             " Where Enq_Detail_code = 	" & txtenqdetailintcode.Text & " and RFQ_Int_code = " & txtRFQIntcode.Text & " and Enq_Reg_NO = " & txtRegNo.Text & ""

        End If

        cnSQL.Open()
        cnSQL2.Open()
        cmSQL = New SqlCommand(strsql, cnSQL)

        If cmSQL.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot save  RFQ details " & strsql, MsgBoxStyle.Exclamation, "Error!")
            Exit Sub

        End If


        Dim strsql1 As String



        If ComboBoxPriceStatus.Text = "Accepted-Line Closed" And datagridEnquiryPending.CurrentRow.Cells(32).Value.ToString = "Project" And datagridEnquiryPending.CurrentRow.Cells(33).Value.ToString = "Forward to Apl. Dept" Then


            strsql1 = "update ENQ_Details set ItemStatus = 'J' Where Enq_Detail_code = " & txtenqdetailintcode.Text & ""
            cmSQL = New SqlCommand(strsql1, cnSQL)

        ElseIf ComboBoxPriceStatus.Text = "Accepted-Line Closed" And datagridEnquiryPending.CurrentRow.Cells(32).Value.ToString <> "Project" And datagridEnquiryPending.CurrentRow.Cells(33).Value.ToString = "Forward to Apl. Dept" Then

            strsql1 = "update ENQ_Details set ItemStatus = 'U' Where Enq_Detail_code = " & txtenqdetailintcode.Text & ""
            cmSQL = New SqlCommand(strsql1, cnSQL)
        ElseIf ComboBoxPriceStatus.Text = "Accepted-Line Closed" And datagridEnquiryPending.CurrentRow.Cells(32).Value.ToString <> "Project" Then
            'Or datagridEnquiryPending.CurrentRow.Cells(33).Value.ToString <> "Forward to Apl. Dept" Or datagridEnquiryPending.CurrentRow.Cells(33).Value.ToString <> "Released to Customer Sup" Then

            strsql1 = "update ENQ_Details set ItemStatus = 'U' Where Enq_Detail_code = " & txtenqdetailintcode.Text & ""
            cmSQL = New SqlCommand(strsql1, cnSQL)

        ElseIf ComboBoxPriceStatus.Text = "Accepted-Line Open" Then

            strsql1 = "update ENQ_Details set ItemStatus = 'H' Where Enq_Detail_code = " & txtenqdetailintcode.Text & ""
            cmSQL = New SqlCommand(strsql1, cnSQL)

        ElseIf ComboBoxPriceStatus.Text = "Rejected" Then

            strsql1 = "update ENQ_Details set ItemStatus = 'R' Where Enq_Detail_code = " & txtenqdetailintcode.Text & ""
            cmSQL = New SqlCommand(strsql1, cnSQL)

        ElseIf ComboBoxPriceStatus.Text = "Released to Customer Sup" Then

            strsql1 = "update ENQ_Details set ItemStatus = 'C' Where Enq_Detail_code = " & txtenqdetailintcode.Text & ""
            cmSQL = New SqlCommand(strsql1, cnSQL)

        ElseIf ComboBoxPriceStatus.Text = "Forwarded to App Dept" Then

            strsql1 = "update ENQ_Details set ItemStatus = 'F' Where Enq_Detail_code = " & txtenqdetailintcode.Text & ""
            cmSQL = New SqlCommand(strsql1, cnSQL)

        End If

        If cmSQL.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot save price details " & strsql, MsgBoxStyle.Exclamation, "Error!")
            Application.Exit()
        Else

            MsgBox("Price detail saved.", vbInformation)
            txtRFQIntcode.Text = ""
            PendingRFQs()

            'check

            Dim m As Integer
            Dim n As Integer
            Dim drSQL1 As SqlDataReader
            Dim drSQL2 As SqlDataReader


            strsql1 = "SELECT count(Enq_Int_code) FROM ENQ_Details WHERE (Req = 'Both' OR Req = 'Price')  and  Enq_Int_code  = " & txtintcode.Text & ""


            'and (Req = "Both" or Req = "Price")"
            cmSQL = New SqlCommand(strsql1, cnSQL)
            drSQL1 = cmSQL.ExecuteReader()
            If drSQL1.Read() Then
                m = drSQL1.Item(0)
            Else
                m = 0

            End If

            strsql1 = "SELECT  COUNT(Enq_Reg_NO)  FROM ENQ_RFQ_PriceDetails WHERE  Status = 'Accepted-Line Closed' AND Enq_Reg_NO  = " & txtRegNo.Text & ""
            cmSQL2 = New SqlCommand(strsql1, cnSQL2)
            drSQL2 = cmSQL2.ExecuteReader()
            If drSQL2.Read() Then
                n = drSQL2.Item(0)
            Else
                n = 0

            End If

            If m = n Then
                sendmail()
                ' MsgBox("Mail sent to ISR,CSR and Segment Heads")
                MsgBox("Mail sent to ISR", vbInformation)
                Exit Sub

            End If



            'checking and sending mail

            '         txtRFQIntcode.Text = ""

            Exit Sub
        End If

    End Sub
    Public Sub sendmail()
        Dim chartrange As Excel.Range

        'excel file
        'Make Connection ' Ammar
        ' Dim cnn As DataAccess = New DataAccess(CONNECTION_STRING)
        Dim cnn As SqlConnection = New SqlConnection(ConnectionStringNew)
        ' Variable ' Ammar
        Dim i, j As Integer
        'Excel WorkBook object ' Ammar
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        xlApp = New Microsoft.Office.Interop.Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        ' Sheet Name or Number ' Ammar
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        '  xlWorkBook.Sheets.Select("A1:A2")

        Dim sql As String

        sql = "SELECT Enq_Reg_NO, Enq_Reg_date, CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                    "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty  " & _
                    " FROM  TSS_RFQPriceViewQuickNew a " & _
                   "WHERE  a.Enq_Reg_NO = " & txtRegNo.Text & "  ORDER BY a.Enq_Reg_NO, a.Sl_no"



        '" & txtRegNo.Text & "

        ' SqlAdapter
        Dim dscmd As New SqlDataAdapter(sql, cnn.ConnectionString)
        ' DataSet
        Dim ds As New DataSet
        dscmd.Fill(ds)
        'COLUMN NAME ADD IN EXCEL SHEET OR HEADING 


        xlWorkSheet.Cells(1, 1).Value = "RegNo"
        xlWorkSheet.Cells(2, 1).Value = "Reg Dt"
        xlWorkSheet.Cells(3, 1).Value = "Customer"
        xlWorkSheet.Cells(4, 1).Value = "CSR"

        xlWorkSheet.Cells(1, 2) = _
               ds.Tables(0).Rows(i).Item(0)
        xlWorkSheet.Cells(2, 2) = _
              ds.Tables(0).Rows(i).Item(1)
        xlWorkSheet.Cells(3, 2) = _
              ds.Tables(0).Rows(i).Item(2)
        xlWorkSheet.Cells(4, 2) = _
              ds.Tables(0).Rows(i).Item(3)


        chartrange = xlWorkSheet.Range("A1", "D4")
        chartrange.Font.Size = 12
        chartrange.Font.Bold = True
        chartrange.Font.Color = 240

        chartrange.BorderAround(Excel.XlLineStyle.xlContinuous, _
        Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex. _
        xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)


        chartrange = xlWorkSheet.Range("A1", "A4")
        chartrange.BorderAround(Excel.XlLineStyle.xlContinuous, _
        Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex. _
        xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)

        chartrange.ColumnWidth = 14




        xlWorkSheet.Cells(6, 1).Value = "Slno"
        xlWorkSheet.Cells(6, 2).Value = "PartNumber"
        xlWorkSheet.Cells(6, 3).Value = "Part Description"
        chartrange = xlWorkSheet.Range("C6", "C6")
        chartrange.ColumnWidth = 18


        xlWorkSheet.Cells(6, 4).Value = "Qty"

        xlWorkSheet.Cells(6, 5).Value = "Qty Type"
        chartrange = xlWorkSheet.Range("E6", "E6")
        chartrange.ColumnWidth = 16


        xlWorkSheet.Cells(6, 6).Value = "Price/Each"
        xlWorkSheet.Cells(6, 7).Value = "MOQ"
        xlWorkSheet.Cells(6, 8).Value = "SPU"

        xlWorkSheet.Cells(6, 9).Value = "Leadtime"
        xlWorkSheet.Cells(6, 10).Value = "Type"
        xlWorkSheet.Cells(6, 11).Value = "Stock Avble"
        chartrange = xlWorkSheet.Range("K6", "K6")
        chartrange.ColumnWidth = 13

        xlWorkSheet.Cells(6, 12).Value = "Special Inst" 'K
        chartrange = xlWorkSheet.Range("L6", "L6")
        chartrange.ColumnWidth = 45


        xlWorkSheet.Cells(6, 13).Value = "Tool Req"
        xlWorkSheet.Cells(6, 14).Value = "Proto Total Cost"
        xlWorkSheet.Cells(6, 15).Value = "Proto Customer Cost"
        xlWorkSheet.Cells(6, 16).Value = "Proto Lead Time"
        xlWorkSheet.Cells(6, 17).Value = "Proto Qty"
        xlWorkSheet.Cells(6, 18).Value = "Proto Life of Tool"
        xlWorkSheet.Cells(6, 19).Value = "Prod Total Cost"
        xlWorkSheet.Cells(6, 20).Value = "Prod Customer Cost"
        xlWorkSheet.Cells(6, 21).Value = "Prod Lead Time"
        xlWorkSheet.Cells(6, 22).Value = "Prod Qty"

        chartrange = xlWorkSheet.Range("N6", "V6")
        chartrange.ColumnWidth = 14


        ' SQL Table Transfer to Excel
        For i = 0 To ds.Tables(0).Rows.Count - 1
            'Column
            For j = 4 To ds.Tables(0).Columns.Count - 1
                ' this i change to header line cells >>>
                xlWorkSheet.Cells(i + 7, j - 3) = _
                ds.Tables(0).Rows(i).Item(j)
            Next
        Next
        'HardCode in Excel sheet
        ' this i change to footer line cells  >>>
        '''''xlWorkSheet.Cells(i + 3, 7) = "Total"
        '''''xlWorkSheet.Cells.Item(i + 3, 8) = "=SUM(H2:H18)"
        ' Save as path of excel sheet


        'HEADING LINE

        chartrange = xlWorkSheet.Range("A6", "V6")
        chartrange.BorderAround(Excel.XlLineStyle.xlContinuous, _
        Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex. _
        xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)

        'END OF HEADING LINE

        Dim a As Integer = ds.Tables(0).Rows.Count
        a = a + 6
        chartrange = xlWorkSheet.Range("A6", "V" & a)
        chartrange.Font.Size = 10
        chartrange.Font.Bold = True

        'chartrange.Font.Color = 240

        chartrange.BorderAround(Excel.XlLineStyle.xlContinuous, _
        Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex. _
        xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)


        chartrange = xlWorkSheet.Range("A6", "V" & a)
        chartrange.RowHeight = 20


        xlApp.Range("A7", "V" & a).Select()
        xlApp.Selection.Interior.ColorIndex = 34    ' this is ok light blue


        xlApp.Range("A6", "V6").Select()
        xlApp.Selection.Interior.ColorIndex = 27   'this is ok yellow colour

        xlApp.Range("A1", "D4").Select()
        xlApp.Selection.Interior.ColorIndex = 8  ' aqua colour


        '30 - MAROON
        '4 light and bright green


        chartrange.BorderAround(Excel.XlLineStyle.xlContinuous, _
        Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex. _
        xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)

        Dim REG As String

        REG = txtRegNo.Text


        xlWorkSheet.SaveAs("C:\FOCUSMAIL\" & REG & ".xlsx")

        ' xlWorkSheet.SaveAs("C:\FOCUSMAIL\RFQ.xlsx")

        xlWorkBook.Close()
        xlApp.Quit()

        ' Create an Outlook application.
        Dim oApp As outlook._Application
        oApp = New outlook.Application()

        ' Create a new MailItem.
        Dim oMsg As outlook._MailItem
        oMsg = oApp.CreateItem(outlook.OlItemType.olMailItem)
        oMsg.Subject = "FOCUS SOFTWARE :  AUTOMATED MAIL. Reg No." & REG & " PRICE RECEIVED"
        ' oMsg.Body = "Dear,  Price Details are enclosed herewith." & vbCr & "Quoted prices are LANDED COST.You need to add required margin before quoting to customers" & vbCr & "For further details please refer focus software. " & vbCr & "Thanks and Regards : TSSI Purchase Team "

        ' TODO: Replace with a valid e-mail address.
        ' oMsg.To = "indira.shetty@trelleborg.com"

        'select  email ids

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        'Dim strSQL1 As String
        Dim t As String
        Dim cc As String
        ' Dim name As String

        If Len(datagridEnquiryPending.CurrentRow.Cells(8).Value.ToString) > 3 Then

            sql = "Select ISRMAILID, CSRMAILID, SEGHEADMAILID FROM TSS_CUSTOMERID_ISR where CustomerID = '" & txtcustcode.Text & "' "

        Else
            sql = "SELECT  ISRMAILID, CSRMAILID, SEGHEADMAILID FROM  TSS_CUSTOMER_ISR_NEWCUSTOMER WHERE RegNo  = '" & txtRegNo.Text & "' "
        End If



        cnSQL1.Open()
        cmSQL1 = New SqlCommand(sql, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then

            Else

                t = drSQL1.Item(0)
                ' If IsDBNull(drSQL1(1)) Then
                'cc = "@trelleborg.com"
                'Else
                'cc = drSQL1(1)
                'End If




                cc = "Libin.George@trelleborg.com" & ";" & "indira.shetty@trelleborg.com"
                't = "indira.shetty@trelleborg.com"
                'cc = "indira.shetty@trelleborg.com"

                Dim name As String = t
                name = name.Substring(0, name.Length - 15)



                '  DotCom = Email.Substring(5, 4)



                '(15, name.Length - 15)
                'name = name.Length - 15

                'Dim blah As String = str.Substring(0, str.IndexOf("\"))

                oMsg.Body = "Dear " & name & "," & vbCrLf & vbCrLf & "Price Details are enclosed herewith, for the above said enquiry." & vbCr & "Quoted price/s are LANDED COST." & vbCr & "You need to add required MARGIN before quoting to customers." & vbCr & "Cerificate charges are not included here (where applicable)." & vbCr & "For further details please refer Focus Software. " & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Purchase Team "




                oMsg.To = t
                oMsg.CC = cc

                ' oMsg.To = "'" & drSQL1.Item(1) & "'"
                ' oMsg.CC = "'" & drSQL1.Item(0) & "'"
                'oMsg.CC = "'" & drSQL1.Item(2) & "'"
            End If

        End If





        'end of select email ids


        ' Add an attachment
        ' TODO: Replace with a valid attachment path.
        'Dim sSource As String = "C:\MYDATA\RFQ.xlsx"
        Dim sSource As String = ("C:\FOCUSMAIL\" & REG & ".xlsx")

        ' TODO: Replace with attachment name
        Dim sDisplayName As String = "RFQ.xlsx"

        Dim sBodyLen As String = oMsg.Body.Length
        Dim oAttachs As outlook.Attachments = oMsg.Attachments
        Dim oAttach As outlook.Attachment
        oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)

        ' Send
        oMsg.Send()

        ' Clean up
        oApp = Nothing
        oMsg = Nothing
        oAttach = Nothing
        oAttachs = Nothing
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub RFQcodegen()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(RFQ_Int_code)from ENQ_RFQ_PriceDetails"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtRFQIntcode.Text = 1
            Else
                txtRFQIntcode.Text = drSQL1.Item(0) + 1
            End If

        End If


    End Sub
    Public Sub fillqty()

        DataGridQty.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        '  Dim stockDCQ As DataSet = New DataSet
        strSql = ""

        If multiple = "YES" Then

            'not used this as on 12 april 2013
            strSql = "SELECT  Price, Factor, FinalPrice,Qty,Qty_Type as Type FROM ENQ_RFQ_Qty_PriceDetails " & _
                                "WHERE  Enq_Detail_code  = " & txtenqdetailintcode.Text & " and Enq_Reg_NO = " & txtRegNo.Text & " and RFQ_Int_code = " & txtRFQIntcode.Text & " "


        ElseIf multiple = "NO" Then



            If (txtitemstatus.Text = "H" Or txtitemstatus.Text = "U" Or txtitemstatus.Text = "C" Or txtitemstatus.Text = "F") And rfqmode = "" Then

                strSql = "SELECT  Price, Factor, FinalPrice, Qty,Qty_Type as Type FROM ENQ_RFQ_Qty_PriceDetails " & _
                        "WHERE RFQ_Int_code = '" & txtRFQIntcode.Text & "' " & _
                            "ORDER BY Qty"

            ElseIf txtitemstatus.Text = "P" And rfqmode = "" Then

                strSql = "SELECT  0.00 as Price,0.00 as Fact, 0.00 as FPrice,Qty,Qty_Type as Type,Enq_Qty_IntCode as IntCode FROM ENQ_Qty_Details " & _
                                 "WHERE  Enq_Detail_code = '" & txtenqdetailintcode.Text & "' " & _
                                    "ORDER BY Qty"


            ElseIf (txtitemstatus.Text = "P" Or txtitemstatus.Text = "H") And rfqmode = "addprice" Then

                strSql = "SELECT  0.00 as Price,0.00 as Fact, 0.00 as FPrice,Qty,Qty_Type as Type,Enq_Qty_IntCode as IntCode FROM ENQ_Qty_Details " & _
                                             "WHERE  Enq_Detail_code = '" & txtenqdetailintcode.Text & "' " & _
                                                "ORDER BY Qty"


            ElseIf (txtitemstatus.Text = "R") Then

                strSql = "SELECT  0.00 as Price,0.00 as Fact, 0.00 as FPrice,Qty,Qty_Type as Type,Enq_Qty_IntCode as IntCode FROM ENQ_Qty_Details " & _
                                             "WHERE  Enq_Detail_code = '" & txtenqdetailintcode.Text & "' " & _
                                                "ORDER BY Qty"



            End If
        End If

        Dim stockDCQ As DataSet = New DataSet
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDQC As SqlDataAdapter = New SqlDataAdapter

        stockDQC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDQC.TableMappings.Add("Table", "Part")
        'get data
        stockDQC.Fill(stockDCQ)



        DataGridQty.DataSource = stockDCQ.Tables(0)
        sqlCon.Close()
        DataGridQty.Expand(-1)

        countqty = stockDCQ.Tables(0).Rows.Count



    End Sub

    Private Sub DataGridQty_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridQty.CurrentCellChanged
        'If e.ColumnIndex = 3 And Also e.RowIndex - 1 Then
        'e.Value = Convert.ToDouble(DataGridQty(1, .RowIndex).Value) * Convert.ToDouble(DataGridQty(2, e.RowIndex).Value)
        'End If
        'b = datagridEnquiryPending.CurrentCell.ColumnNumber()
        'Dim a As Double

        'If DataGridQty.CurrentCell.ColumnNumber = 1 Then 'AndAlso DataGridQty.CurrentRowIndex - 1 Then
        'a = Convert.ToDouble(DataGridQty(1, DataGridQty.CurrentRowIndex).value) * Convert.ToDouble(DataGridQty(1, DataGridQty.CurrentRowIndeex).value)

        'End If

    End Sub

    ' Private Sub DataGridQty_CellBeginEdit(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) Handles DataGridQty.CellBeginEdit
    '    If e.ColumnIndex = 3 Then
    '       e.Cancel = True
    '  End If
    ' End Sub

    'Private Sub DataGridQty_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles DataGridQty.CellFormatting
    '   If e.ColumnIndex = 3 AndAlso e.RowIndex - 1 Then
    '      e.Value = Convert.ToDouble(DataGridQty(1, .RowIndex).Value) * Convert.ToDouble(DataGridQty(2, e.RowIndex).Value)
    ' End If
    ' End Sub
    'Private Sub DataGridView1_CellValidated(ByVal sender As Object,
    'ByVal e As DataGridViewCellEventArgs) Handles
    '       DataGridView1.CellValidated()
    '      DataGridView1.Refresh()
    ' End Sub


    Private Sub DataGridQty_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridQty.Navigate

    End Sub
    Private Sub Saveqtydetails()

        Dim st As String
        Dim strsql2 As String


        'Dim cmSQL As SqlCommand

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cn As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim cm As SqlCommand

        Dim dr As SqlDataReader
        'Dim drSQL2 As SqlDataReader

        Dim exist As Boolean
        exist = False

        Dim i As Integer
        Dim k As Integer
        Dim m As Integer



        Dim qty As Double
        Dim price As Double
        Dim qtype As String
        Dim fact As Double
        Dim fprice As Double

        Dim counta As Integer
        'strsql3 = ""
        strsql2 = ""
        st = ""
        counta = 0

        i = 0
        k = 0
        m = 0


        curdate = System.DateTime.Now()


        txtitemstatus.Text = Trim(txtitemstatus.Text)


        Do While i < countqty
            price = DataGridQty.Item(i, k)
            fact = DataGridQty.Item(i, k + 1)

            If Val(DataGridQty.Item(i, k + 2)) = 0 Then
                fprice = price * fact

            Else
                fprice = DataGridQty.Item(i, k + 2)
            End If

            'fprice = price * fact
            'fprice = DataGridQty.Item(i, k + 2)

            qty = DataGridQty.Item(i, k + 3)
            qtype = DataGridQty.Item(i, k + 4)
            counta = counta + 1
            i = i + 1
            qtype = Trim(qtype)

            'check that qty is already saved for that detailcode

            st = "SELECT Enq_Detail_code, Qty FROM ENQ_RFQ_Qty_PriceDetails where  Enq_Detail_code = " & txtenqdetailintcode.Text & " and  Qty = " & qty & " and  Enq_Reg_NO = " & txtRegNo.Text & ""

            '            cnSQL1.Open()  'NOW COMMENTED
            'If cn.Open = True Then
            'cn.Close()
            'End If
            cn.Close()
            cn.Open()



            cm = New SqlCommand(st, cn)
            dr = cm.ExecuteReader()

            If dr.Read() Then

                If IsDBNull(dr.Item(0)) Then
                    exist = False
                Else
                    exist = True
                End If

            End If


            If txtitemstatus.Text = "P" And rfqmode = "" And exist = False Then
                strsql2 = "insert ENQ_RFQ_Qty_PriceDetails values(" & txtRegNo.Text & "," & txtenqdetailintcode.Text & "," & txtRFQIntcode.Text & "," & _
                "" & qty & "," & price & ",'" & curdate & "','" & curdate & "', '" & username & "','" & qtype & "'," & fact & ", " & fprice & ")"


            ElseIf txtitemstatus.Text = "P" And rfqmode = "" And exist = True Then

                strsql2 = "update ENQ_RFQ_Qty_PriceDetails set Price = " & price & ", Date_Modify = '" & curdate & "', UserID = '" & username & "',  Factor = " & fact & " , FinalPrice = " & fprice & " " & _
           " where Enq_Reg_NO = " & txtRegNo.Text & " and Enq_Detail_code	= " & txtenqdetailintcode.Text & " and RFQ_Int_code =" & txtRFQIntcode.Text & " and Qty_Type = '" & qtype & "'"



            ElseIf (txtitemstatus.Text = "H" Or txtitemstatus.Text = "F" Or txtitemstatus.Text = "C") And rfqmode = "" Then

                '             strsql = "update ENQ_RFQ_Qty_PriceDetails set Price = " & price & ", Date_Modify = '" & curdate & "', UserID = '" & username & "', Qty_Type = '" & qtype & "', Factor = " & fact & " , FinalPrice = " & fprice & " " & _
                '            " where Enq_Reg_NO = " & txtRegNo.Text & " and Enq_Detail_code	= " & txtenqdetailintcode.Text & " and RFQ_Int_code =" & txtRFQIntcode.Text & " "

                strsql2 = "update ENQ_RFQ_Qty_PriceDetails set Price = " & price & ", Date_Modify = '" & curdate & "', UserID = '" & username & "',  Factor = " & fact & " , FinalPrice = " & fprice & " " & _
               " where Enq_Reg_NO = " & txtRegNo.Text & " and Enq_Detail_code	= " & txtenqdetailintcode.Text & " and RFQ_Int_code =" & txtRFQIntcode.Text & " and Qty_Type = '" & qtype & "'"


                'ElseIf (txtitemstatus.Text = "H" Or txtitemstatus.Text = "P") And rfqmode = "addprice" Then

                '   strsql = "insert ENQ_RFQ_Qty_PriceDetails values(" & txtRegNo.Text & "," & txtenqdetailintcode.Text & "," & txtRFQIntcode.Text & "," & _
                '   "" & qty & "," & price & ",'" & curdate & "','" & curdate & "', '" & username & "','" & qtype & "'," & fact & ", " & fprice & ")"

            End If
            cnSQL1.Close()
            cnSQL1.Open()
            cmSQL1 = New SqlCommand(strsql2, cnSQL1)

            If cmSQL1.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot save the price ! " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If

        Loop


    End Sub


    Private Sub BtnAddPrice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddPrice.Click
        rfqmode = "addprice"

        clearpricedetails()
        fillqty()
        RFQcodegen()


    End Sub

    Private Sub clearpricedetails()


        txtQuoteRef1.Text = ""
        txtQuoteRef2.Text = ""
        txtQuoteRef3.Text = ""

        ComboBoxPriceStatus.Text = ""
        txtLeadTime.Text = ""
        txtMOQ.Text = ""
        txtMov.Text = ""
        txtSPU.Text = ""
        comboboxstocktype.Text = ""
        txtstockavble.Text = ""

        txtdetailremarks.Text = ""
        txtVendorName.Text = ""
        txtvendorcontact.Text = ""
        txtvendorref.Text = ""
        txtprogressdetails.Text = ""
        txtvendquote.Text = ""
        txtreasonforwarding.Text = ""
        txtSpecialInst.Text = ""

        ProtoLeadTime.Text = 0
        ProtoTotal.Text = 0
        ProtoLifeofTool.Text = 0
        ProtoCustShare.Text = 0
        ProtoQty.Text = 0

        ProdLeadTime.Text = 0
        ProdLifeofTool.Text = 0
        ProdCustShare.Text = 0
        ProdQty.Text = 0
        ProdTotalCost.Text = 0



    End Sub

    Public Sub callrfqdetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String
        Dim strsql As String
        Dim pricecount As Integer

        strsql = ""
        strSQL1 = ""

        If multiple = "NO" Then
            strsql = "select count(*) from ENQ_RFQ_PriceDetails where Enq_Detail_code  = " & txtenqdetailintcode.Text & "  "


            cnSQL1.Open()
            cmSQL1 = New SqlCommand(strsql, cnSQL1)
            drSQL1 = cmSQL1.ExecuteReader()

            If drSQL1.Read() Then

                If IsDBNull(drSQL1.Item(0)) Then

                Else
                    pricecount = drSQL1(0)
                End If

            End If

            drSQL1.Close()

            If pricecount > 1 Then
                callnext()
                Exit Sub


                'btnnext.Visible = True
                'btnnext.Enabled = True
            End If


            strSQL1 = "select * from ENQ_RFQ_PriceDetails where Enq_Detail_code  = " & txtenqdetailintcode.Text & "  "

        ElseIf multiple = "YES" Then

            strSQL1 = "select * from ENQ_RFQ_PriceDetails where Enq_Detail_code  = " & txtenqdetailintcode.Text & " and Enq_Reg_NO = " & txtRegNo.Text & " and RFQ_Int_code = " & txtRFQIntcode.Text & " "
            cnSQL1.Open()

        End If

        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                Exit Sub
            Else

                If Trim((drSQL1.Item(4))) = "Group" Then

                    RadioButtonGroup.Checked = True
                    RadioButton3P.Checked = False
                    RadioButtonFactory.Checked = False

                ElseIf Trim((drSQL1.Item(4))) = "3rdParty" Then

                    RadioButtonGroup.Checked = False
                    RadioButton3P.Checked = True
                    RadioButtonFactory.Checked = False

                ElseIf Trim((drSQL1.Item(4))) = "Factory" Then

                    RadioButtonFactory.Checked = True
                    RadioButton3P.Checked = False
                    RadioButtonGroup.Checked = False


                End If

                If Len(drSQL1.Item(12)) > 3 Then
                    RadioButtonExisting.Checked = True
                    RadioButtonNew.Checked = False
                Else
                    RadioButtonNew.Checked = True
                    RadioButtonExisting.Checked = False

                End If

                txtVendorName.Text = drSQL1.Item(33)


                txtRFQIntcode.Text = (drSQL1.Item(1))
                ComboBoxPriceStatus.Text = (drSQL1.Item(3))
                If (drSQL1.Item(3)) = "Forwarded to App Dept" Then
                    lblForwardApl.Visible = True
                    txtreasonforwarding.Visible = True
                    txtreasonforwarding.Enabled = True

                End If


                txtLeadTime.Text = (drSQL1.Item(7))
                txtMOQ.Text = (drSQL1.Item(5))
                txtSPU.Text = (drSQL1.Item(6))

                'comboboxstocktype.DropDownStyle = ComboBoxStyle.Simple
                'comboboxcurrency.DropDownStyle = ComboBoxStyle.Simple

                'comboboxstocktype.Text = (drSQL1.Item(8))
                'comboboxcurrency.Text = (drSQL1.Item(9))

                comboboxstocktype.DropDownStyle = ComboBoxStyle.DropDownList
                comboboxcurrency.DropDownStyle = ComboBoxStyle.DropDownList

                comboboxstocktype.Text = Trim((drSQL1.Item(8)))
                comboboxcurrency.Text = Trim((drSQL1.Item(9)))


                txtstockavble.Text = (drSQL1.Item(10))
                'txttooling.Text = (drSQL1.Item(10))
                txtvendorcontact.Text = (drSQL1.Item(13))
                txtvendorref.Text = (drSQL1.Item(12))
                txtvendquote.Text = (drSQL1.Item(14))

                txtdetailremarks.Text = (drSQL1.Item(11))
                txtSpecialInst.Text = (drSQL1.Item(15))

                ' txtprogressdetails.Text = (drSQL1.Item(16))

                LatestAction.Text = (drSQL1.Item(18))


                txtMov.Text = (drSQL1.Item(19))
                If Trim((drSQL1.Item(20))) = "YES" Then
                    RBToolYes.Checked = True
                Else
                    RBToolNo.Checked = False

                End If
                ProtoTotal.Text = (drSQL1.Item(21))
                ProtoCustShare.Text = (drSQL1.Item(22))
                ProtoLeadTime.Text = (drSQL1.Item(23))
                ProtoQty.Text = (drSQL1.Item(24))
                ProtoLifeofTool.Text = (drSQL1.Item(25))
                ProdTotalCost.Text = (drSQL1.Item(26))
                ProdCustShare.Text = (drSQL1.Item(27))
                ProdLeadTime.Text = (drSQL1.Item(28))
                ProdQty.Text = (drSQL1.Item(29))
                ProdLifeofTool.Text = (drSQL1.Item(30))

                txtprogressdetails.Text = (drSQL1.Item(31))
                txtreasonforwarding.Text = (drSQL1.Item(32))

                'QuoteRef1, QuoteRef2,
                'QuoteRef3()



                If IsDBNull(drSQL1.Item(34)) Then
                    txtQuoteRef1.Text = "_"
                Else
                    txtQuoteRef1.Text = (drSQL1.Item(34))
                End If

                If IsDBNull(drSQL1.Item(35)) Then
                    txtQuoteRef2.Text = "_"
                Else
                    txtQuoteRef2.Text = (drSQL1.Item(35))
                End If

                If IsDBNull(drSQL1.Item(36)) Then
                    txtQuoteRef3.Text = "_"
                Else
                    txtQuoteRef3.Text = (drSQL1.Item(36))
                End If


                If IsDBNull(drSQL1.Item(37)) Then
                    txtAltMtrl.Text = "_"
                Else
                    txtAltMtrl.Text = (drSQL1.Item(37))
                End If

            End If

        End If

    End Sub

    Private Sub callnext()
        'callrfqdetails()
        'fillqty()
        multiple = "YES"

        DatagridMultiprices.Visible = True

        DatagridMultiprices.Height = 120
        DatagridMultiprices.Width = 1050

        fillmultipleprices()
    End Sub

    Public Sub fillmultipleprices()

        DatagridMultiprices.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet

        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader


        strSql = "select RFQ_Int_code,Status,Source_Mtrl,MOQ,SPU,LeadTime,Type,Currency,Tooling_Cost,Stock_Avble,Remarks,Vendor_Ref,Name,Vendor_Quote,Special_Remarks," & _
        "Enq_Detail_code,Enq_Reg_NO from ENQ_RFQ_PriceDetails where Enq_Detail_code  = '" & txtenqdetailintcode.Text & "'"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDQC As SqlDataAdapter = New SqlDataAdapter

        stockDQC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDQC.TableMappings.Add("Table", "Part")
        'get data
        stockDQC.Fill(stockDCQ)


        DatagridMultiprices.DataSource = stockDCQ.Tables(0)
        sqlCon.Close()
        DatagridMultiprices.Expand(-1)

    End Sub


    Private Sub DatagridMultiprices_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub

    Private Sub DatagridMultiprices_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        clearpricedetails()


        Dim b As Integer
        'Dim custid As String
        b = datagridEnquiryPending1.CurrentCell.ColumnNumber()

        If b = 0 Then


            txtRFQIntcode.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell)

            If DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 2) = "Group" Then
                RadioButtonGroup.Checked = True
                RadioButton3P.Checked = False
                RadioButtonFactory.Checked = False
            ElseIf DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 2) = "3rdParty" Then
                RadioButtonGroup.Checked = False
                RadioButtonFactory.Checked = False
                RadioButton3P.Checked = True

            ElseIf DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 2) = "Factory" Then
                RadioButtonGroup.Checked = False
                RadioButtonFactory.Checked = True
                RadioButton3P.Checked = False


            End If
            '            strSql = "select RFQ_Int_code,Status,Source_Mtrl,MOQ,SPU,LeadTime,Type,Currency,Tooling_Cost,Stock_Avble,Remarks,
            'Vendor_Ref(, Name, Vendor_Quote, Special_Remarks, " & _")
            '            "Enq_Detail_code,Enq_Reg_NO from ENQ_RFQ_PriceDetails where Enq_Detail_code  = '" & txtenqdetailintcode.Text & "'"



            ComboBoxPriceStatus.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 1)
            txtLeadTime.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 5)
            txtMOQ.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 3)
            txtSPU.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 4)
            comboboxstocktype.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 6)
            comboboxcurrency.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 7)
            txtstockavble.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 9)
            'txttooling.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 8)
            txtdetailremarks.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 10)
            txtvendorcontact.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 11)
            txtvendorref.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 12)
            txtprogressdetails.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 14)
            txtvendquote.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 13)


            txtRegNo.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 15)
            txtenqdetailintcode.Text = datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 16)

            fillqty()
            DatagridMultiprices.Visible = False
        Else
            MsgBox("Click on firlst column", vbInformation)
            Exit Sub

        End If



    End Sub

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click

    End Sub

    Private Sub txtLeadTime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLeadTime.KeyPress
        'Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        'If allowedChars.IndexOf(e.KeyChar) = -1 Then
        '' Invalid Character
        'e.Handled = True
        'End If

        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtstockavble_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtstockavble.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txttooling_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtdetailremarks_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdetailremarks.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtvendorref_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtvendorref.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtvendorcontact_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtvendorcontact.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtvendquote_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtvendquote.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtspecial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtprogressdetails.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Public Sub PendingRFQs()

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet


        '       strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, UserId, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
        '          "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1,Enq_Type, Enq_Forward from TSS_Enquiry_Pending_Price order by RegNo,SlNo"

        strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, UserId, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
          "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1,Enq_Type, Enq_Forward,Status,Enq_Int_code " & _
           "from TSS_Enquiry_Pending_Price order by RegNo,SlNo"






        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)



        datagridEnquiryPending1.DataSource = stockDC.Tables(0)
        cnSQL.Close()
        datagridEnquiryPending1.Expand(-1)


    End Sub

    Private Sub DatagridMultiprices_Navigate_1(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DatagridMultiprices.Navigate

    End Sub

    Private Sub DatagridMultiprices_CurrentCellChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DatagridMultiprices.CurrentCellChanged
        Dim b As Integer

        b = DatagridMultiprices.CurrentCell.ColumnNumber()

        If b = 0 Then
            clearpricedetails()

            txtRFQIntcode.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell)

            txtRegNo.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 16)
            txtenqdetailintcode.Text = DatagridMultiprices.Item(DatagridMultiprices.CurrentCell.RowNumber, 15)

            callrfqdetails()

            fillqty()

        Else
            MsgBox("Click on first column", vbInformation)
            Exit Sub


        End If

    End Sub

    Private Sub ToolDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolDetails.Enter

    End Sub

    Private Sub RBToolYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBToolYes.CheckedChanged
        If RBToolYes.Checked = True Then
            '   ToolDetails.Visible = True
            '  ToolDetails.Width = 646
            ' ToolDetails.Height = 115


        End If
    End Sub

    Private Sub LatestAction_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtMov_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMov.KeyDown

    End Sub

    Private Sub txtMov_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMov.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMov.TextChanged

    End Sub

    Private Sub txtPurSpecial_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPurSpecial.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtPurSpecial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPurSpecial.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPurSpecial.TextChanged

    End Sub

    Private Sub txtreasonforwarding_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtreasonforwarding.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtreasonforwarding_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtreasonforwarding.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtreasonforwarding_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtreasonforwarding.TextChanged

    End Sub

    Private Sub LabelToolClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelToolClose.Click
        ToolDetails.Visible = False
    End Sub

    Private Sub ToolFrameOpen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolFrameOpen.CheckedChanged

        If RBToolYes.Checked = True Then

            If ToolFrameOpen.Checked = True Then
                ToolDetails.Visible = True
                'ToolDetails.BringToFront()
                'ToolDetails.Top = 624
                'ToolDetails.Left = 99

                'ToolDetails.Location.X = 624
                'ToolDetails.Location.Y = 99

                ToolDetails.Height = 115
                ToolDetails.Width = 646
            End If
        End If

    End Sub
    Private Sub Savecertificatedetails()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim i As Integer
        Dim k As Integer
        Dim m As Integer

        Dim protoprice As Double
        Dim prodPrice As Double
        Dim certificate As String


        Dim counta As Integer
        strsql = ""
        counta = 0

        i = 0
        k = 0
        m = 0


        curdate = System.DateTime.Now()
        cnSQL1.Open()

        'txtitemstatus.Text = Trim(txtitemstatus.Text)

        Do While i < countcerqty
            protoprice = DataGridCertificateCharges.Item(i, k)
            prodPrice = DataGridCertificateCharges.Item(i, k + 1)
            certificate = DataGridCertificateCharges.Item(i, k + 2)
            counta = counta + 1
            i = i + 1

            If txtitemstatus.Text = "P" And rfqmode = "" Then

                strsql = "insert ENQ_EnqWise_Certificates_Charges values(" & txtRegNo.Text & "," & txtenqdetailintcode.Text & "," & txtRFQIntcode.Text & "," & _
                    "'" & certificate & "'," & protoprice & "," & prodPrice & ",'" & curdate & "','" & curdate & "', '" & username & "')"


            ElseIf (txtitemstatus.Text = "H" Or txtitemstatus.Text = "F" Or txtitemstatus.Text = "C") And rfqmode = "" Then

                strsql = "update ENQ_EnqWise_Certificates_Charges  set Proto_Price = " & protoprice & ", Prod_Price = " & prodPrice & "," & _
                        " Date_Modify = '" & curdate & "', UserId = '" & username & "'" & _
                         " where Enq_Reg_NO = " & txtRegNo.Text & " and Enq_Detail_code	= " & txtenqdetailintcode.Text & " and RFQ_Int_code =" & txtRFQIntcode.Text & " and Certificates =  '" & certificate & "' "

            End If


            cmSQL = New SqlCommand(strsql, cnSQL1)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot save the price ! " & strsql, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If

        Loop


    End Sub

    Public Sub fillcertificate()

        ' If (txtitemstatus.Text = "H" Or txtitemstatus.Text = "C" Or txtitemstatus.Text = "F") And rfqmode = "" Then

        'strSql = "SELECT  Price, Factor, FinalPrice, Qty,Qty_Type as Type FROM ENQ_RFQ_Qty_PriceDetails " & _
        '       "WHERE RFQ_Int_code = '" & txtRFQIntcode.Text & "' " & _
        '          "ORDER BY Qty"

        ' ElseIf txtitemstatus.Text = "P" And rfqmode = "" Then


        DataGridCertificateCharges.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet
        strSql = ""
        If txtitemstatus.Text = "P" And rfqmode = "" Then


            strSql = "SELECT  0.00 as Protoprice,0.00 as ProdPrice,Certificates from ENQ_EnqWise_Certificates " & _
                                         "WHERE  Enq_Detail_code = " & txtenqdetailintcode.Text & " and Enq_Reg_NO = " & txtRegNo.Text & " "
        ElseIf (txtitemstatus.Text = "H" Or txtitemstatus.Text = "U" Or txtitemstatus.Text = "C" Or txtitemstatus.Text = "F") And rfqmode = "" Then

            strSql = "Select Proto_Price, Prod_Price,Certificates from ENQ_EnqWise_Certificates_Charges " & _
             "WHERE  Enq_Detail_code = " & txtenqdetailintcode.Text & " and Enq_Reg_NO = " & txtRegNo.Text & " "


        ElseIf txtitemstatus.Text = "R" Then


            strSql = "SELECT  0.00 as Protoprice,0.00 as ProdPrice,Certificates from ENQ_EnqWise_Certificates " & _
                                         "WHERE  Enq_Detail_code = " & txtenqdetailintcode.Text & " and Enq_Reg_NO = " & txtRegNo.Text & " "


        End If

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDQC As SqlDataAdapter = New SqlDataAdapter

        stockDQC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDQC.TableMappings.Add("Table", "Part")
        'get data
        stockDQC.Fill(stockDCQ)


        DataGridCertificateCharges.DataSource = stockDCQ.Tables(0)
        sqlCon.Close()
        DataGridCertificateCharges.Expand(-1)

        countcerqty = stockDCQ.Tables(0).Rows.Count



    End Sub



    Private Sub txtdocdetails_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtdocdetails.TextChanged

    End Sub

    Private Sub ProtoTotal_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProtoTotal.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProtoTotal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProtoTotal.TextChanged

    End Sub

    Private Sub ProtoCustShare_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProtoCustShare.GotFocus
        If Val(ProtoCustShare.Text) = 0 Then
            ProtoCustShare.Text = ProtoTotal.Text
        End If
    End Sub

    Private Sub ProtoTSSShare_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProtoCustShare.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProtoTSSShare_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProtoCustShare.TextChanged

    End Sub

    Private Sub ProtoLeadTime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProtoLeadTime.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProtoLeadTime_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProtoLeadTime.TextChanged

    End Sub

    Private Sub ProtoQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProtoQty.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProtoQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProtoQty.TextChanged

    End Sub

    Private Sub ProtoLifeofTool_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProtoLifeofTool.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProtoLifeofTool_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProtoLifeofTool.TextChanged

    End Sub

    Private Sub ProdTotalCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProdTotalCost.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProdTotalCost_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProdTotalCost.TextChanged

    End Sub

    Private Sub ProdCustShare_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProdCustShare.GotFocus
        If Val(ProdCustShare.Text) = 0 Then
            ProdCustShare.Text = ProdTotalCost.Text
        End If
    End Sub

    Private Sub ProdTSSSh_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProdCustShare.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProdTSSSh_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProdCustShare.TextChanged

    End Sub

    Private Sub ProdLeadTime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProdLeadTime.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProdLeadTime_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProdLeadTime.TextChanged

    End Sub

    Private Sub ProdQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProdQty.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProdQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProdQty.TextChanged

    End Sub

    Private Sub ProdLifeofTool_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ProdLifeofTool.KeyPress
        Dim allowedChars As String = "0123456789." & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ProdLifeofTool_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProdLifeofTool.TextChanged

    End Sub

    Private Sub DataGridCertificateCharges_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridCertificateCharges.Navigate

    End Sub

    Public Sub clearqty()

        DataGridQty.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet

        'Dim cmSQL As SqlCommand
        ' Dim drSQL As SqlDataReader

        ' If lblMode1.Text = "Add" Then
        strSql = "SELECT Qty,Qty_Type, Enq_Qty_IntCode FROM ENQ_Qty_Details " & _
                 "WHERE  Enq_Int_code=  '000000000000000'"


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDQC As SqlDataAdapter = New SqlDataAdapter

        stockDQC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDQC.TableMappings.Add("Table", "Part")
        'get data
        stockDQC.Fill(stockDCQ)


        DataGridQty.DataSource = stockDCQ.Tables(0)
        sqlCon.Close()
        DataGridQty.Expand(-1)


    End Sub

    Private Sub ClearCertificate()

        DataGridCertificateCharges.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet

        'Dim cmSQL As SqlCommand
        ' Dim drSQL As SqlDataReader

        ' If lblMode1.Text = "Add" Then
        strSql = "SELECT  0.00 as Protoprice,0.00 as ProdPrice,Certificates from ENQ_EnqWise_Certificates " & _
                                 "WHERE  Enq_Detail_code = '000000000000000000'"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDQC As SqlDataAdapter = New SqlDataAdapter

        stockDQC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDQC.TableMappings.Add("Table", "Part")
        'get data
        stockDQC.Fill(stockDCQ)


        DataGridCertificateCharges.DataSource = stockDCQ.Tables(0)
        sqlCon.Close()
        DataGridCertificateCharges.Expand(-1)

    End Sub
    Sub fillvendorlist()

        DataGridVendor.Show()

        Dim sqlcon As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        txtvendorref.Text = txtvendorref.Text & "%"

        txtVendorName.Text = txtVendorName.Text & "%"

        If Len(txtvendorref.Text) > 1 Then

            strSql = "SELECT VendorID, VendorName, VendorCity FROM FSDBBR.dbo.FS_Vendor " & _
                 "WHERE VendorID LIKE '" & txtvendorref.Text & "' " & _
                    "ORDER BY VendorID"
        Else

            strSql = "SELECT VendorID, VendorName, VendorCity FROM FSDBBR.dbo.FS_Vendor " & _
                  "WHERE VendorName LIKE '" & txtVendorName.Text & "' " & _
                     "ORDER BY VendorName"

        End If


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlcon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        sqlcon.Open()

        stockDAC.TableMappings.Add("Table", "Customer")
        'get data
        stockDAC.Fill(stockDC)

        DataGridVendor.Width = 650 '1150
        DataGridVendor.Height = 320 '800

        DataGridVendor.DataSource = stockDC.Tables(0)
        sqlcon.Close()
        DataGridVendor.Expand(-1)

    End Sub

    Private Sub DataGridVendor_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridVendor.CurrentCellChanged
        Dim a As Integer
        'Dim custid As String

        a = DataGridVendor.CurrentCell.ColumnNumber()

        If a = 0 Then
            txtvendorref.Text = DataGridVendor.Item(DataGridVendor.CurrentCell)

            txtVendorName.Text = DataGridVendor.Item(DataGridVendor.CurrentCell.RowNumber, 1)

        Else
            MsgBox("Click on VendorID to select the Vendor", vbInformation)
            Exit Sub
        End If

        DataGridVendor.Hide()

    End Sub

    Private Sub DataGridVendor_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridVendor.DoubleClick

    End Sub

    Private Sub DataGridVendor_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridVendor.Navigate

    End Sub

    Private Sub datagridEnquiryPending_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridEnquiryPending.CellContentClick

    End Sub

    Private Sub datagridEnquiryPending_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles datagridEnquiryPending.RowHeaderMouseClick

        'CODING TO BE WRITTEN
        'txtRegNo.Text = DataGridViewProjectPending.CurrentRow.Cells(0).Value.ToString

        txtRFQIntcode.Text = ""
        RadioButtonGroup.Checked = True
        comboboxcurrency.Text = "EUR"
        RBToolNo.Checked = True

        If DatagridMultiprices.Visible = True Then
            DatagridMultiprices.Visible = False
        End If

        rfqmode = ""
        multiple = "NO"

        'Dim b As Integer
        'Dim custid As String
        'b = datagridEnquiryPending.CurrentCell.ColumnNumber()

        '  If b = 0 Then
        clearpricedetails()

        txtenqdetailintcode.Text = datagridEnquiryPending.CurrentRow.Cells(0).Value


        txtRegNo.Text = datagridEnquiryPending.CurrentRow.Cells(1).Value
        DtpEnqRegDt.Value = datagridEnquiryPending.CurrentRow.Cells(2).Value

        txtCustomer.Text = datagridEnquiryPending.CurrentRow.Cells(9).Value.ToString
        txtcustcode.Text = datagridEnquiryPending.CurrentRow.Cells(8).Value.ToString
        txtCity.Text = datagridEnquiryPending.CurrentRow.Cells(30).Value.ToString
        TXTCL3.Text = datagridEnquiryPending.CurrentRow.Cells(11).Value.ToString
        TXTCL1.Text = datagridEnquiryPending.CurrentRow.Cells(31).Value.ToString
        txtCSR.Text = datagridEnquiryPending.CurrentRow.Cells(12).Value.ToString
        txtISR.Text = datagridEnquiryPending.CurrentRow.Cells(13).Value.ToString
        txtTSSISeg.Text = datagridEnquiryPending.CurrentRow.Cells(14).Value.ToString
        txtTSSSeg.Text = datagridEnquiryPending.CurrentRow.Cells(15).Value.ToString
        If datagridEnquiryPending.CurrentRow.Cells(16).ToString = "YES" Then
            RadioButtonExisting.Checked = True
            RadioButtonNew.Checked = False
        Else
            RadioButtonNew.Checked = True
            RadioButtonExisting.Checked = False
        End If
        'If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 10)) Then

        If IsDBNull(datagridEnquiryPending.CurrentRow.Cells(10).Value.ToString) Then
            RadioButtonDomestic.Checked = True
            RadioButtonExport.Checked = False
        ElseIf Trim(datagridEnquiryPending.CurrentRow.Cells(10).Value.ToString) = "Domestic" Then
            RadioButtonDomestic.Checked = True
            RadioButtonExport.Checked = False
        ElseIf Trim(datagridEnquiryPending.CurrentRow.Cells(10).Value.ToString) = "Export" Then
            RadioButtonDomestic.Checked = False
            RadioButtonExport.Checked = True
        End If
        If Trim(datagridEnquiryPending.CurrentRow.Cells(17).Value.ToString) = "YES" Then
            RadioButtondocyes.Checked = True
            RadioButtondocno.Checked = False
        Else
            RadioButtondocyes.Checked = False
            RadioButtondocno.Checked = True
        End If

        If IsDBNull(Trim(datagridEnquiryPending.CurrentRow.Cells(18).Value.ToString) = "Export") Then
            txtdocdetails.Text = ""
        Else

            txtdocdetails.Text = Trim(datagridEnquiryPending.CurrentRow.Cells(18).Value.ToString) = "Export"
        End If

        If IsDBNull(Trim(datagridEnquiryPending.CurrentRow.Cells(19).Value.ToString) = "Export") Then
            txtSpecialInst.Text = ""
        Else

            txtSpecialInst.Text = Trim(datagridEnquiryPending.CurrentRow.Cells(19).Value.ToString)
        End If


        txtslno.Text = datagridEnquiryPending.CurrentRow.Cells(3).Value
        ComboBoxFSYesNo.Text = datagridEnquiryPending.CurrentRow.Cells(20).Value.ToString

        ComboBoxItemSource.Text = datagridEnquiryPending.CurrentRow.Cells(21).Value.ToString
        txtpart.Text = datagridEnquiryPending.CurrentRow.Cells(4).Value.ToString
        txtPartDesc.Text = datagridEnquiryPending.CurrentRow.Cells(5).Value.ToString
        txtCustPart.Text = datagridEnquiryPending.CurrentRow.Cells(6).Value.ToString
        txtCustDesc.Text = datagridEnquiryPending.CurrentRow.Cells(22).Value.ToString
        ComboBoxuom.Text = datagridEnquiryPending.CurrentRow.Cells(7).Value.ToString

        If IsDBNull(datagridEnquiryPending.CurrentRow.Cells(23).Value.ToString) Then
            txtRecVend.Text = ""
        Else

            txtRecVend.Text = datagridEnquiryPending.CurrentRow.Cells(23).Value.ToString
        End If


        If IsDBNull(datagridEnquiryPending.CurrentRow.Cells(24).Value.ToString) Then
            txtDimension.Text = ""
        Else


            txtDimension.Text = datagridEnquiryPending.CurrentRow.Cells(24).Value.ToString

        End If

        If IsDBNull(datagridEnquiryPending.CurrentRow.Cells(25).Value.ToString) Then
            txtMaterial.Text = ""
        Else

            txtMaterial.Text = datagridEnquiryPending.CurrentRow.Cells(25).Value.ToString

        End If

        If IsDBNull(datagridEnquiryPending.CurrentRow.Cells(26).Value.ToString) Then
            txtDetailSpecial.Text = ""

        Else

            txtDetailSpecial.Text = datagridEnquiryPending.CurrentRow.Cells(26).Value.ToString
        End If

        If datagridEnquiryPending.CurrentRow.Cells(27).ToString = "01-01-1900" Then
            dtpenqduedt.Checked = False

            dtpenqduedt.Value = "01-01-1900"
        Else
            dtpenqduedt.Checked = True

            dtpenqduedt.Value = datagridEnquiryPending.CurrentRow.Cells(27).Value

        End If


        txtitemstatus.Text = datagridEnquiryPending.CurrentRow.Cells(28).Value.ToString

        If Trim(txtitemstatus.Text) = "R" Then

            ComboBoxPriceStatus.Text = "Rejected"

        End If



        txtitemstatus.Text = Trim(txtitemstatus.Text)

        If screentype <> "COMP" Then
            txtintcode.Text = datagridEnquiryPending.CurrentRow.Cells(35).Value.ToString
        End If

        'EditCertDetails()

        If txtitemstatus.Text = "H" Or txtitemstatus.Text = "U" Or txtitemstatus.Text = "C" Or txtitemstatus.Text = "F" Then
            callrfqdetails()

        End If

        ' Else

        'MsgBox("Click on Detailcode ", vbInformation)
        'Exit Sub
        'End If

        If multiple = "YES" Then
            Exit Sub
        Else
            clearqty()
            fillqty()

            ClearCertificate()
            fillcertificate()


        End If




    End Sub

    Private Sub BtnMail_Click(sender As Object, e As EventArgs) Handles BtnMail.Click

        Dim OutlookMessage As outlook.MailItem
        Dim AppOutlook As New outlook.Application
        Try
            OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
            Dim Recipents As outlook.Recipients = OutlookMessage.Recipients
            Recipents.Add("indira.shetty@trelleborg.com")
            OutlookMessage.Subject = "Automated Mail form Focus Sofware - Price Received"
            OutlookMessage.Body = "Pl refere focus software regarding Reg.No.     Thanks and Regards. PURCHASE TEAM"
            OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
            OutlookMessage.Send()
        Catch ex As Exception
            MessageBox.Show("Mail could not be sent") 'if you dont want this message, simply delete this line  
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try




    End Sub

    Private Sub BtnHistory_Click(sender As Object, e As EventArgs) Handles BtnHistory.Click
        ' GroupBoxPartNumberHistory.Visible = True
        'GroupBoxPartNumberHistory.BringToFront()

        'GroupBoxPartNumberHistory.Width = 731
        'GroupBoxPartNumberHistory.Height = 629

        'DataGridViewHistory.Height = 604
        'DataGridViewHistory.Width = 687
        If Len(txtpart.Text) > 4 Then

            RFQHistory.Location = New System.Drawing.Point(575, 24)
            RFQHistory.Width = 1180
            RFQHistory.Height = 656
            parthistory = Trim(txtpart.Text)
            RFQHistory.Show()



        End If

    End Sub

    Private Sub txtQuoteRef1_KeyDown(sender As Object, e As KeyEventArgs) Handles txtQuoteRef1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtQuoteRef1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuoteRef1.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtQuoteRef1_TextChanged(sender As Object, e As EventArgs) Handles txtQuoteRef1.TextChanged

    End Sub

    Private Sub txtQuoteRef2_KeyDown(sender As Object, e As KeyEventArgs) Handles txtQuoteRef2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtQuoteRef2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuoteRef2.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtQuoteRef2_TextChanged(sender As Object, e As EventArgs) Handles txtQuoteRef2.TextChanged

    End Sub

    Private Sub txtQuoteRef3_KeyDown(sender As Object, e As KeyEventArgs) Handles txtQuoteRef3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtQuoteRef3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuoteRef3.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtQuoteRef3_TextChanged(sender As Object, e As EventArgs) Handles txtQuoteRef3.TextChanged

    End Sub

    Private Sub BtnQuoteRefSave_Click(sender As Object, e As EventArgs) Handles BtnQuoteRefSave.Click
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim sql As String


        sql = "update  ENQ_RFQ_PriceDetails set QuoteRef1 = '" & txtQuoteRef1.Text & "', QuoteRef2 = '" & txtQuoteRef2.Text & "', QuoteRef3 = '" & txtQuoteRef3.Text & "' " & _
             " Where Enq_Detail_code = 	" & txtenqdetailintcode.Text & " and RFQ_Int_code = " & txtRFQIntcode.Text & " and Enq_Reg_NO = " & txtRegNo.Text & ""


        cnSQL.Open()
        cmSQL = New SqlCommand(sql, cnSQL)

        If cmSQL.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot save  RFQ details " & sql, MsgBoxStyle.Exclamation, "Error!")
            Exit Sub

        End If



    End Sub

    Private Sub Btnsearch_Click(sender As Object, e As EventArgs) Handles Btnsearch.Click




        datagridEnquiryPending.Enabled = True

        RBToolNo.Checked = True
        RadioButtonGroup.Checked = True
        RadioButtonVendorYes.Checked = True


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        If Len(Trim(txtReg.Text)) = 0 Then
            txtReg.Text = ""
            txtReg.Text = "'%'"
        Else
            'txtReg.Text = txtReg.Text & "%"

        End If

        If Len(Trim(txtCID.Text)) = 0 Then
            txtCID.Text = ""
            txtCID.Text = "%"
        Else
            txtCID.Text = txtCID.Text & "%"
        End If


        If Len(Trim(txtCustN.Text)) = 0 Then
            txtCustN.Text = ""
            txtCustN.Text = "%"
        Else
            txtCustN.Text = txtCustN.Text & "%"
        End If


        If Len(Trim(txtPartN.Text)) = 0 Then
            txtPartN.Text = ""
            txtPartN.Text = "%"
        Else
            txtPartN.Text = txtPartN.Text & "%"

        End If



        If screentype = "PEND" Then

            GroupBox2.Text = "Request for Quotation"


            btnRFQSave.Visible = True



            strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, UserId, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                  "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1,Enq_Type, Enq_Forward,Status,Enq_Int_code " & _
                   "from TSS_Enquiry_Pending_Price where RegNo like " & txtReg.Text & " and  CustomerID like '" & txtCID.Text & "' and CustomerName like '" & txtCustN.Text & "' and PartNumber like '" & txtPartN.Text & "' order by RegNo,SlNo"



        ElseIf screentype = "PENDP" Then

            GroupBox2.Text = "Request for Quotation"
            btnRFQSave.Visible = True

            strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, UserId, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                     "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1,Enq_Type, Enq_Forward,Enq_Int_code " & _
                     "from TSS_Enquiry_Pending_Price where Enq_Type in ('Project','Project-Budgetary') and Enq_Forward = 'Forward to Apl. Dept' AND [Reg.Date] >= '11-01-2013' and RegNo like " & txtReg.Text & " and  CustomerID like '" & txtCID.Text & "' and CustomerName like '" & txtCustN.Text & "' and PartNumber like '" & txtPartN.Text & "' order by RegNo,SlNo"


        ElseIf screentype = "COMP" Then

            GroupBox2.Text = "RFQ Completed"
            btnRFQSave.Visible = False
            BtnDelete.Visible = False
            BtnQuoteRefSave.Visible = True

            'btnRFQSave.Visible = True
            'btnRFQSave.Name = "QRef Save"

            BtnDelete.Visible = False
            '            strSQL = "Select * from TSS_Enquiry_Price_Completed order by RegNo,SlNo"

            strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                     "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1 from TSS_Enquiry_Price_Completed where RegNo like " & txtReg.Text & " and  CustomerID like '" & txtCID.Text & "' and CustomerName like '" & txtCustN.Text & "' and PartNumber like '" & txtPartN.Text & "' order by RegNo,SlNo"

        ElseIf screentype = "COMPP" Then

            GroupBox2.Text = "RFQ Completed"

            btnRFQSave.Visible = False
            BtnDelete.Visible = False
            'btnRFQSave.Text = "QRefSave"

            strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
           "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1 from TSS_Enquiry_Price_Completed " & _
           "where Enq_Type in ('Project','Project-Budgetary') and Enq_Forward = 'Forward to Apl. Dept' AND [Reg.Date] >= '11-01-2013' and RegNo like '" & txtReg.Text & "' and  CustomerID like '" & txtCID.Text & "' and CustomerName like '" & txtCustN.Text & "' and PartNumber like '" & txtPartN.Text & "' order by RegNo,SlNo"



        End If


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)



        datagridEnquiryPending.DataSource = stockDC.Tables(0)
        'datagridEnquiryPending.Expand(-1)


        'colouring unattended enquiries

        'datagridEnquiryPending.Rows()
        ' Dim a As String


        If screentype = "PEND" Then

            For i As Integer = 0 To datagridEnquiryPending.RowCount - 2
                '    a = ""
                ' a = datagridEnquiryPending.CurrentRow.Cells("Status").Value.ToString

                'a = datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString
                'If IsDBNull(datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) Or Len(a) < 2 Then

                '   If IsDBNull(datagridEnquiryPending.Rows(i).Cells("Status").ToString) Or Len(datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) < 2 Then


                If (datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) <> "-" Then

                    'datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Red
                    datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Black

                Else
                    'datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Black
                    datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Red

                End If


            Next

        End If



        'end of colouring






        cnSQL.Close()
        '        datagridEnquiryPending.Expand(-1)





1:


        listloadCertificate()








    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        txtReg.Text = ""
        txtCustN.Text = ""
        txtCID.Text = ""
        txtPartN.Text = ""

    End Sub

    Private Sub LOADING()

        ' If screentype = "PEND" Or screentype = "PENDP" Then

        datagridEnquiryPending.Enabled = True

        RBToolNo.Checked = True
        RadioButtonGroup.Checked = True
        RadioButtonVendorYes.Checked = True


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        If screentype = "PEND" Then

            GroupBox2.Text = "Request for Quotation"


            btnRFQSave.Visible = True



            strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, UserId, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                  "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1,Enq_Type, Enq_Forward,Status,Enq_Int_code " & _
                   "from TSS_Enquiry_Pending_Price order by RegNo,SlNo"




        ElseIf screentype = "PENDP" Then

            GroupBox2.Text = "Request for Quotation"
            btnRFQSave.Visible = True

            strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, UserId, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                     "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1,Enq_Type, Enq_Forward,Enq_Int_code " & _
                     "from TSS_Enquiry_Pending_Price where Enq_Type in ('Project','Project-Budgetary') and Enq_Forward = 'Forward to Apl. Dept' AND [Reg.Date] >= '11-01-2013' order by RegNo,SlNo"


        ElseIf screentype = "COMP" Then

            GroupBox2.Text = "RFQ Completed"
            btnRFQSave.Visible = False
            BtnDelete.Visible = False
            BtnQuoteRefSave.Visible = True

            'btnRFQSave.Visible = True
            'btnRFQSave.Name = "QRef Save"

            BtnDelete.Visible = False
            '            strSQL = "Select * from TSS_Enquiry_Price_Completed order by RegNo,SlNo"

            strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
                     "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1 from TSS_Enquiry_Price_Completed order by RegNo,SlNo"

        ElseIf screentype = "COMPP" Then

            GroupBox2.Text = "RFQ Completed"

            btnRFQSave.Visible = False
            BtnDelete.Visible = False
            'btnRFQSave.Text = "QRefSave"

            strSQL = "SELECT Enq_Detail_code, RegNo, [Reg.Date],SlNo, PartNumber, PartDescription,CustPartNumber,uom,CustomerID, CustomerName,MarketType, Class, CSR, ISR, TSSISeg, TSSSeg, Cust_Exist_New, Doc_upload, Doc_Details, " & _
           "Special_instructions, FS_Yes_NO, Part_Source,  CustPartDescription,  RecomVendor, Dimension, Material, Special, Enq_Due_date, ItemStatus, Req,City as CustomerCity,Class1 from TSS_Enquiry_Price_Completed " & _
           "where Enq_Type in ('Project','Project-Budgetary') and Enq_Forward = 'Forward to Apl. Dept' AND [Reg.Date] >= '11-01-2013' order by RegNo,SlNo"



        End If


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)



        datagridEnquiryPending.DataSource = stockDC.Tables(0)
        'datagridEnquiryPending.Expand(-1)


        'colouring unattended enquiries

        'datagridEnquiryPending.Rows()
        ' Dim a As String


        If screentype = "PEND" Then

            For i As Integer = 0 To datagridEnquiryPending.RowCount - 2
                '    a = ""
                ' a = datagridEnquiryPending.CurrentRow.Cells("Status").Value.ToString

                'a = datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString
                'If IsDBNull(datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) Or Len(a) < 2 Then

                '   If IsDBNull(datagridEnquiryPending.Rows(i).Cells("Status").ToString) Or Len(datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) < 2 Then


                If (datagridEnquiryPending.Rows(i).Cells("Status").Value.ToString) <> "-" Then

                    'datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Red
                    datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Black

                ElseIf (datagridEnquiryPending.Rows(i).Cells("Enq_Type").Value.ToString) <> "Internal RFQ" Then
                    ' datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Black
                    datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Red

                ElseIf (datagridEnquiryPending.Rows(i).Cells("Enq_Type").Value.ToString) = "Internal RFQ" Then
                    ' datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Black
                    datagridEnquiryPending.Rows(i).Cells("Enq_Detail_code").Style.ForeColor = Color.Blue


                End If


            Next

        End If



        'end of colouring






        cnSQL.Close()
        '        datagridEnquiryPending.Expand(-1)





1:


        listloadCertificate()

        '  End If

    End Sub

    Private Sub txtAltMtrl_KeyDown(sender As Object, e As KeyEventArgs) Handles txtAltMtrl.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtAltMtrl_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAltMtrl.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtAltMtrl_MultilineChanged(sender As Object, e As EventArgs) Handles txtAltMtrl.MultilineChanged

    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles txtAltMtrl.TextChanged

    End Sub

    Private Sub txtVendorName_TextChanged(sender As Object, e As EventArgs) Handles txtVendorName.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LOADING()

    End Sub

    Private Sub txtRFQIntcode_TextChanged(sender As Object, e As EventArgs) Handles txtRFQIntcode.TextChanged

    End Sub
End Class
