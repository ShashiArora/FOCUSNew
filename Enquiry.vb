'*************************************************************************************************************************************************************
'Software               : Enquiry Reg software
'Data                   : 
'Views                  : 
'Tabes                  : 
'Software by            : Indira, IT dept. TSS
'Sofware completed date : 
'Sofware in use         : 
'Software modified on   : 
'Modification           : 
'**************************************************************************************************************************************************************
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports outlook = Microsoft.Office.Interop.Outlook
Imports Excel = Microsoft.Office.Interop.Excel


Imports System.Windows.Forms

Imports System.Math
Imports System.Data.OleDb
Imports System.IO.StreamWriter

Imports System.Text
Imports Microsoft.VisualBasic
Imports System.Net.WebRequest
Imports System.Net.WebClient
Imports System.Net
Imports System.IO

Imports System.Collections.Generic
Imports System.Drawing








Public Class Enquiry
    Inherits System.Windows.Forms.Form

    Private ConnectionString As String

    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"
               
           

    Public purcheck As Integer
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonId As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonName As System.Windows.Forms.RadioButton
    Friend WithEvents Label56 As System.Windows.Forms.Label
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents txtBothNot As System.Windows.Forms.TextBox
    Friend WithEvents txtDunsno As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonDomestic As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonExport As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents RBCustomerExisting As System.Windows.Forms.RadioButton
    Friend WithEvents RBCustomerNew As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBoxPartCreation As System.Windows.Forms.GroupBox
    Friend WithEvents lblmb As System.Windows.Forms.Label
    Friend WithEvents lblpartdesc As System.Windows.Forms.Label
    Friend WithEvents lblPart As System.Windows.Forms.Label
    Friend WithEvents txtPartnum As System.Windows.Forms.TextBox
    Friend WithEvents lblChildItem As System.Windows.Forms.Label
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents lblPlanner As System.Windows.Forms.Label
    Friend WithEvents txtfix As System.Windows.Forms.TextBox
    Friend WithEvents txtinsp As System.Windows.Forms.TextBox
    Friend WithEvents txtrun As System.Windows.Forms.TextBox
    Friend WithEvents ComboBoxBuyer As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxPlanner As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxItemType As System.Windows.Forms.ComboBox
    Friend WithEvents txtpartDescription As System.Windows.Forms.TextBox
    Friend WithEvents sp1 As System.Windows.Forms.Label
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtchilditemDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxprodline As System.Windows.Forms.ComboBox
    Friend WithEvents lblLeadTimeType As System.Windows.Forms.Label
    Friend WithEvents txtsp2 As System.Windows.Forms.TextBox
    Friend WithEvents txtsp1 As System.Windows.Forms.TextBox
    Friend WithEvents LblClose As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents btnQtyAdd As System.Windows.Forms.Button
    Friend WithEvents QtyEdit As System.Windows.Forms.Button
    Friend WithEvents lblqty As System.Windows.Forms.Label
    Friend WithEvents txtqtyintcode As System.Windows.Forms.TextBox
    Friend WithEvents btnItemsave As System.Windows.Forms.Button
    Friend WithEvents DataGridQty As System.Windows.Forms.DataGrid
    Friend WithEvents txtqty As System.Windows.Forms.TextBox
    Friend WithEvents ComboBoxReqType As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBoxPartCreation As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxTax As System.Windows.Forms.ComboBox
    Friend WithEvents Label61 As System.Windows.Forms.Label
    Friend WithEvents Label63 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxClass3 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxInvAc As System.Windows.Forms.TextBox
    Friend WithEvents ComboBoxCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents ComboboxISR As System.Windows.Forms.ComboBox
    Friend WithEvents txtISR As System.Windows.Forms.TextBox
    Friend WithEvents PanelProjectDetails As System.Windows.Forms.Panel
    Friend WithEvents txtfax1 As System.Windows.Forms.TextBox
    Friend WithEvents Label69 As System.Windows.Forms.Label
    Friend WithEvents txtemail1 As System.Windows.Forms.TextBox
    Friend WithEvents Label70 As System.Windows.Forms.Label
    Friend WithEvents txtph1 As System.Windows.Forms.TextBox
    Friend WithEvents Label71 As System.Windows.Forms.Label
    Friend WithEvents txtMob1 As System.Windows.Forms.TextBox
    Friend WithEvents Label72 As System.Windows.Forms.Label
    Friend WithEvents TxtBuyerName1 As System.Windows.Forms.TextBox
    Friend WithEvents Label66 As System.Windows.Forms.Label
    Friend WithEvents txtDept1 As System.Windows.Forms.TextBox
    Friend WithEvents txtDesig1 As System.Windows.Forms.TextBox
    Friend WithEvents Label67 As System.Windows.Forms.Label
    Friend WithEvents Label68 As System.Windows.Forms.Label
    Friend WithEvents txtBusinessPotential As System.Windows.Forms.TextBox
    Friend WithEvents Label74 As System.Windows.Forms.Label
    Friend WithEvents dtpProjectStDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label73 As System.Windows.Forms.Label
    Friend WithEvents txt As System.Windows.Forms.Button
    Friend WithEvents Label75 As System.Windows.Forms.Label
    Friend WithEvents CheckedListBoxComp As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label76 As System.Windows.Forms.Label
    Friend WithEvents GroupBoxPriceAvble As System.Windows.Forms.GroupBox
    Friend WithEvents LabelClose As System.Windows.Forms.Label
    Friend WithEvents datagridPriceAvble As System.Windows.Forms.DataGridView
    Friend WithEvents btnItemHistory As System.Windows.Forms.Button

    Public stockDA As SqlDataAdapter = New SqlDataAdapter

    'Public invno As String
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DataGridCustomer As System.Windows.Forms.DataGrid
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents txtslno As System.Windows.Forms.TextBox
    Friend WithEvents btnsave As System.Windows.Forms.Button
    Friend WithEvents DataUpdation As System.Windows.Forms.GroupBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxSource As System.Windows.Forms.ComboBox
    Friend WithEvents DTPEnqRecd As System.Windows.Forms.DateTimePicker
    Friend WithEvents ComboBoxClarity As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxEnquiryType As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxCategory As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxRejectionReasons As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxEnquiryStatus As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxPriceStatus As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxStatusRemarks As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxForward As System.Windows.Forms.ComboBox
    Friend WithEvents txtEnqRef As System.Windows.Forms.TextBox
    Friend WithEvents dtpEnqDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents RBTenderNo As System.Windows.Forms.RadioButton
    Friend WithEvents RBTenderYes As System.Windows.Forms.RadioButton
    Friend WithEvents dtpTenderDueDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtSpecial As System.Windows.Forms.TextBox
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents txtCustID As System.Windows.Forms.TextBox
    Friend WithEvents txtVat As System.Windows.Forms.TextBox
    Friend WithEvents txtCst As System.Windows.Forms.TextBox
    Friend WithEvents txtEcc As System.Windows.Forms.TextBox
    Friend WithEvents txtFax As System.Windows.Forms.TextBox
    Friend WithEvents txtemail As System.Windows.Forms.TextBox
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtMobile As System.Windows.Forms.TextBox
    Friend WithEvents txtDesignation As System.Windows.Forms.TextBox
    Friend WithEvents txtCustCountry As System.Windows.Forms.TextBox
    Friend WithEvents txtContact As System.Windows.Forms.TextBox
    Friend WithEvents txtCustPin As System.Windows.Forms.TextBox
    Friend WithEvents txtCustState As System.Windows.Forms.TextBox
    Friend WithEvents txtCustcity As System.Windows.Forms.TextBox
    Friend WithEvents txtCustAdr3 As System.Windows.Forms.TextBox
    Friend WithEvents txtCustAd1 As System.Windows.Forms.TextBox
    Friend WithEvents txtCustAdr2 As System.Windows.Forms.TextBox
    Friend WithEvents txtRejected As System.Windows.Forms.TextBox
    Friend WithEvents txtPartNot As System.Windows.Forms.TextBox
    Friend WithEvents txtPriceNot As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalItems As System.Windows.Forms.TextBox
    Friend WithEvents btnUpload As System.Windows.Forms.Button
    Friend WithEvents txtDocDetails As System.Windows.Forms.TextBox
    Friend WithEvents btnHeaderSave As System.Windows.Forms.Button
    Friend WithEvents DTPStatusDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtPartYesPriceYes As System.Windows.Forms.TextBox
    Friend WithEvents txtenqintcode As System.Windows.Forms.TextBox
    Friend WithEvents GroupYesNo As System.Windows.Forms.GroupBox
    Friend WithEvents rbdocyes As System.Windows.Forms.RadioButton
    Friend WithEvents ComboboxClass As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxCSR As System.Windows.Forms.ComboBox
    Friend WithEvents ComboboxTSSISeg As System.Windows.Forms.ComboBox
    Friend WithEvents ComboboxSegment As System.Windows.Forms.ComboBox
    Friend WithEvents txtcustintcode As System.Windows.Forms.TextBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxItemSource As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxFSYesNo As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxuom As System.Windows.Forms.ComboBox
    Friend WithEvents txtpart As System.Windows.Forms.TextBox
    Friend WithEvents txtCustPart As System.Windows.Forms.TextBox
    Friend WithEvents txtPartDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterial As System.Windows.Forms.TextBox
    Friend WithEvents txtDetailSpecial As System.Windows.Forms.TextBox
    Friend WithEvents txtRecVend As System.Windows.Forms.TextBox
    Friend WithEvents txtDimension As System.Windows.Forms.TextBox
    Friend WithEvents txtCustDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents txtdetailintcode As System.Windows.Forms.TextBox
    Friend WithEvents datagridDetail As System.Windows.Forms.DataGrid
    Friend WithEvents DataGridPartNumbers As System.Windows.Forms.DataGrid
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Protected WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnItemDelete As System.Windows.Forms.Button
    Friend WithEvents lblEnqAdd As System.Windows.Forms.Label
    Friend WithEvents lblmode As System.Windows.Forms.Label
    Friend WithEvents lblMode1 As System.Windows.Forms.Label
    Friend WithEvents DataGridEnquiryEdit As System.Windows.Forms.DataGrid
    Friend WithEvents rbDocNo As System.Windows.Forms.RadioButton
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents CheckedListBoxCertificate As System.Windows.Forms.CheckedListBox
    Friend WithEvents txtitemkey As System.Windows.Forms.TextBox
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents txtDept As System.Windows.Forms.TextBox
    Protected WithEvents txtRegNo As System.Windows.Forms.TextBox
    Protected WithEvents DTPRegDt As System.Windows.Forms.DateTimePicker
    Friend WithEvents ComboboxReq As System.Windows.Forms.ComboBox
    Friend WithEvents Label57 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.datagridDetail = New System.Windows.Forms.DataGrid()
        Me.GroupBoxPriceAvble = New System.Windows.Forms.GroupBox()
        Me.datagridPriceAvble = New System.Windows.Forms.DataGridView()
        Me.LabelClose = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtCustomer = New System.Windows.Forms.TextBox()
        Me.txtRegNo = New System.Windows.Forms.TextBox()
        Me.DataGridCustomer = New System.Windows.Forms.DataGrid()
        Me.txtEnqRef = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.PanelProjectDetails = New System.Windows.Forms.Panel()
        Me.Label76 = New System.Windows.Forms.Label()
        Me.txt = New System.Windows.Forms.Button()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.CheckedListBoxComp = New System.Windows.Forms.CheckedListBox()
        Me.txtBusinessPotential = New System.Windows.Forms.TextBox()
        Me.Label74 = New System.Windows.Forms.Label()
        Me.dtpProjectStDate = New System.Windows.Forms.DateTimePicker()
        Me.Label73 = New System.Windows.Forms.Label()
        Me.txtfax1 = New System.Windows.Forms.TextBox()
        Me.Label69 = New System.Windows.Forms.Label()
        Me.txtemail1 = New System.Windows.Forms.TextBox()
        Me.Label70 = New System.Windows.Forms.Label()
        Me.txtph1 = New System.Windows.Forms.TextBox()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.txtMob1 = New System.Windows.Forms.TextBox()
        Me.Label72 = New System.Windows.Forms.Label()
        Me.TxtBuyerName1 = New System.Windows.Forms.TextBox()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.txtDept1 = New System.Windows.Forms.TextBox()
        Me.txtDesig1 = New System.Windows.Forms.TextBox()
        Me.Label67 = New System.Windows.Forms.Label()
        Me.Label68 = New System.Windows.Forms.Label()
        Me.txtISR = New System.Windows.Forms.TextBox()
        Me.Label61 = New System.Windows.Forms.Label()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.RBCustomerExisting = New System.Windows.Forms.RadioButton()
        Me.RBCustomerNew = New System.Windows.Forms.RadioButton()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonDomestic = New System.Windows.Forms.RadioButton()
        Me.RadioButtonExport = New System.Windows.Forms.RadioButton()
        Me.DataGridEnquiryEdit = New System.Windows.Forms.DataGrid()
        Me.DataGridPartNumbers = New System.Windows.Forms.DataGrid()
        Me.txtDunsno = New System.Windows.Forms.TextBox()
        Me.Label56 = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonId = New System.Windows.Forms.RadioButton()
        Me.RadioButtonName = New System.Windows.Forms.RadioButton()
        Me.Label55 = New System.Windows.Forms.Label()
        Me.txtDept = New System.Windows.Forms.TextBox()
        Me.txtcustintcode = New System.Windows.Forms.TextBox()
        Me.ComboBoxStatusRemarks = New System.Windows.Forms.ComboBox()
        Me.txtSpecial = New System.Windows.Forms.TextBox()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ComboBoxRejectionReasons = New System.Windows.Forms.ComboBox()
        Me.Label50 = New System.Windows.Forms.Label()
        Me.txtRejected = New System.Windows.Forms.TextBox()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.txtPartNot = New System.Windows.Forms.TextBox()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtPriceNot = New System.Windows.Forms.TextBox()
        Me.txtPartYesPriceYes = New System.Windows.Forms.TextBox()
        Me.txtTotalItems = New System.Windows.Forms.TextBox()
        Me.btnUpload = New System.Windows.Forms.Button()
        Me.txtDocDetails = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.btnHeaderSave = New System.Windows.Forms.Button()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.ComboBoxEnquiryStatus = New System.Windows.Forms.ComboBox()
        Me.ComboBoxClarity = New System.Windows.Forms.ComboBox()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.ComboBoxEnquiryType = New System.Windows.Forms.ComboBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.ComboBoxCategory = New System.Windows.Forms.ComboBox()
        Me.ComboBoxPriceStatus = New System.Windows.Forms.ComboBox()
        Me.ComboBoxForward = New System.Windows.Forms.ComboBox()
        Me.txtRemarks = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.txtCustID = New System.Windows.Forms.TextBox()
        Me.txtVat = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtCst = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.txtEcc = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.txtFax = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.txtemail = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.txtMobile = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txtDesignation = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtCustCountry = New System.Windows.Forms.TextBox()
        Me.txtContact = New System.Windows.Forms.TextBox()
        Me.txtCustPin = New System.Windows.Forms.TextBox()
        Me.txtCustState = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtCustcity = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtCustAdr3 = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtCustAd1 = New System.Windows.Forms.TextBox()
        Me.txtCustAdr2 = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DTPStatusDt = New System.Windows.Forms.DateTimePicker()
        Me.GroupYesNo = New System.Windows.Forms.GroupBox()
        Me.rbdocyes = New System.Windows.Forms.RadioButton()
        Me.rbDocNo = New System.Windows.Forms.RadioButton()
        Me.txtBothNot = New System.Windows.Forms.TextBox()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.ComboBoxTax = New System.Windows.Forms.ComboBox()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.ComboboxTSSISeg = New System.Windows.Forms.ComboBox()
        Me.ComboboxISR = New System.Windows.Forms.ComboBox()
        Me.ComboBoxCurrency = New System.Windows.Forms.ComboBox()
        Me.ComboBoxClass3 = New System.Windows.Forms.ComboBox()
        Me.ComboboxClass = New System.Windows.Forms.ComboBox()
        Me.ComboBoxCSR = New System.Windows.Forms.ComboBox()
        Me.ComboboxSegment = New System.Windows.Forms.ComboBox()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.CheckedListBoxCertificate = New System.Windows.Forms.CheckedListBox()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DTPRegDt = New System.Windows.Forms.DateTimePicker()
        Me.dtpEnqDt = New System.Windows.Forms.DateTimePicker()
        Me.DTPEnqRecd = New System.Windows.Forms.DateTimePicker()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.txtpart = New System.Windows.Forms.TextBox()
        Me.txtCustPart = New System.Windows.Forms.TextBox()
        Me.txtPartDesc = New System.Windows.Forms.TextBox()
        Me.txtslno = New System.Windows.Forms.TextBox()
        Me.btnsave = New System.Windows.Forms.Button()
        Me.txtMaterial = New System.Windows.Forms.TextBox()
        Me.txtDetailSpecial = New System.Windows.Forms.TextBox()
        Me.DataUpdation = New System.Windows.Forms.GroupBox()
        Me.btnItemHistory = New System.Windows.Forms.Button()
        Me.GroupBoxPartCreation = New System.Windows.Forms.GroupBox()
        Me.ComboBoxInvAc = New System.Windows.Forms.TextBox()
        Me.LblClose = New System.Windows.Forms.Label()
        Me.ComboBoxprodline = New System.Windows.Forms.ComboBox()
        Me.lblLeadTimeType = New System.Windows.Forms.Label()
        Me.txtsp2 = New System.Windows.Forms.TextBox()
        Me.txtsp1 = New System.Windows.Forms.TextBox()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.txtchilditemDesc = New System.Windows.Forms.TextBox()
        Me.txtfix = New System.Windows.Forms.TextBox()
        Me.txtinsp = New System.Windows.Forms.TextBox()
        Me.txtrun = New System.Windows.Forms.TextBox()
        Me.ComboBoxBuyer = New System.Windows.Forms.ComboBox()
        Me.ComboBoxPlanner = New System.Windows.Forms.ComboBox()
        Me.ComboBoxItemType = New System.Windows.Forms.ComboBox()
        Me.txtpartDescription = New System.Windows.Forms.TextBox()
        Me.sp1 = New System.Windows.Forms.Label()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.lblPlanner = New System.Windows.Forms.Label()
        Me.lblChildItem = New System.Windows.Forms.Label()
        Me.lblmb = New System.Windows.Forms.Label()
        Me.lblpartdesc = New System.Windows.Forms.Label()
        Me.lblPart = New System.Windows.Forms.Label()
        Me.txtPartnum = New System.Windows.Forms.TextBox()
        Me.ComboboxReq = New System.Windows.Forms.ComboBox()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.txtitemkey = New System.Windows.Forms.TextBox()
        Me.btnItemDelete = New System.Windows.Forms.Button()
        Me.txtdetailintcode = New System.Windows.Forms.TextBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.ComboBoxuom = New System.Windows.Forms.ComboBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnQtyAdd = New System.Windows.Forms.Button()
        Me.QtyEdit = New System.Windows.Forms.Button()
        Me.lblqty = New System.Windows.Forms.Label()
        Me.txtqtyintcode = New System.Windows.Forms.TextBox()
        Me.btnItemsave = New System.Windows.Forms.Button()
        Me.DataGridQty = New System.Windows.Forms.DataGrid()
        Me.txtqty = New System.Windows.Forms.TextBox()
        Me.ComboBoxReqType = New System.Windows.Forms.ComboBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.ComboBoxItemSource = New System.Windows.Forms.ComboBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.txtRecVend = New System.Windows.Forms.TextBox()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtDimension = New System.Windows.Forms.TextBox()
        Me.txtCustDesc = New System.Windows.Forms.TextBox()
        Me.ComboBoxFSYesNo = New System.Windows.Forms.ComboBox()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.CheckBoxPartCreation = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblMode1 = New System.Windows.Forms.Label()
        Me.txtenqintcode = New System.Windows.Forms.TextBox()
        Me.RBTenderNo = New System.Windows.Forms.RadioButton()
        Me.RBTenderYes = New System.Windows.Forms.RadioButton()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.ComboBoxSource = New System.Windows.Forms.ComboBox()
        Me.dtpTenderDueDt = New System.Windows.Forms.DateTimePicker()
        Me.lblmode = New System.Windows.Forms.Label()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.lblEnqAdd = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        CType(Me.datagridDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxPriceAvble.SuspendLayout()
        CType(Me.datagridPriceAvble, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.PanelProjectDetails.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.DataGridEnquiryEdit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridPartNumbers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.GroupYesNo.SuspendLayout()
        Me.DataUpdation.SuspendLayout()
        Me.GroupBoxPartCreation.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DataGridQty, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.datagridDetail)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(16, 413)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1248, 160)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Enquiry Item Details"
        '
        'datagridDetail
        '
        Me.datagridDetail.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.datagridDetail.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridDetail.CaptionVisible = False
        Me.datagridDetail.DataMember = ""
        Me.datagridDetail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridDetail.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridDetail.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.datagridDetail.Location = New System.Drawing.Point(5, 17)
        Me.datagridDetail.Name = "datagridDetail"
        Me.datagridDetail.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.datagridDetail.ParentRowsVisible = False
        Me.datagridDetail.PreferredColumnWidth = 85
        Me.datagridDetail.ReadOnly = True
        Me.datagridDetail.RowHeadersVisible = False
        Me.datagridDetail.RowHeaderWidth = 55
        Me.datagridDetail.Size = New System.Drawing.Size(1230, 136)
        Me.datagridDetail.TabIndex = 0
        '
        'GroupBoxPriceAvble
        '
        Me.GroupBoxPriceAvble.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBoxPriceAvble.Controls.Add(Me.datagridPriceAvble)
        Me.GroupBoxPriceAvble.Controls.Add(Me.LabelClose)
        Me.GroupBoxPriceAvble.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBoxPriceAvble.ForeColor = System.Drawing.Color.Red
        Me.GroupBoxPriceAvble.Location = New System.Drawing.Point(504, 72)
        Me.GroupBoxPriceAvble.Name = "GroupBoxPriceAvble"
        Me.GroupBoxPriceAvble.Size = New System.Drawing.Size(706, 256)
        Me.GroupBoxPriceAvble.TabIndex = 179
        Me.GroupBoxPriceAvble.TabStop = False
        Me.GroupBoxPriceAvble.Text = "Price already available "
        Me.GroupBoxPriceAvble.Visible = False
        '
        'datagridPriceAvble
        '
        Me.datagridPriceAvble.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.datagridPriceAvble.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.datagridPriceAvble.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Red
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.Blue
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.datagridPriceAvble.DefaultCellStyle = DataGridViewCellStyle2
        Me.datagridPriceAvble.Location = New System.Drawing.Point(6, 33)
        Me.datagridPriceAvble.Name = "datagridPriceAvble"
        Me.datagridPriceAvble.Size = New System.Drawing.Size(684, 210)
        Me.datagridPriceAvble.TabIndex = 19
        '
        'LabelClose
        '
        Me.LabelClose.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelClose.ForeColor = System.Drawing.Color.Red
        Me.LabelClose.Location = New System.Drawing.Point(664, 11)
        Me.LabelClose.Name = "LabelClose"
        Me.LabelClose.Size = New System.Drawing.Size(18, 23)
        Me.LabelClose.TabIndex = 145
        Me.LabelClose.Text = "X"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(24, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 24)
        Me.Label2.TabIndex = 63
        Me.Label2.Text = "Reg. No."
        '
        'txtCustomer
        '
        Me.txtCustomer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustomer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomer.Location = New System.Drawing.Point(88, 48)
        Me.txtCustomer.MaxLength = 60
        Me.txtCustomer.Name = "txtCustomer"
        Me.txtCustomer.Size = New System.Drawing.Size(328, 20)
        Me.txtCustomer.TabIndex = 14
        '
        'txtRegNo
        '
        Me.txtRegNo.BackColor = System.Drawing.Color.Bisque
        Me.txtRegNo.Font = New System.Drawing.Font("Arial", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegNo.ForeColor = System.Drawing.Color.Red
        Me.txtRegNo.Location = New System.Drawing.Point(104, 24)
        Me.txtRegNo.Name = "txtRegNo"
        Me.txtRegNo.Size = New System.Drawing.Size(72, 24)
        Me.txtRegNo.TabIndex = 1
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
        Me.DataGridCustomer.Location = New System.Drawing.Point(573, 12)
        Me.DataGridCustomer.Name = "DataGridCustomer"
        Me.DataGridCustomer.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridCustomer.ParentRowsVisible = False
        Me.DataGridCustomer.PreferredColumnWidth = 85
        Me.DataGridCustomer.ReadOnly = True
        Me.DataGridCustomer.RowHeadersVisible = False
        Me.DataGridCustomer.Size = New System.Drawing.Size(8, 24)
        Me.DataGridCustomer.TabIndex = 13
        Me.DataGridCustomer.Visible = False
        '
        'txtEnqRef
        '
        Me.txtEnqRef.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEnqRef.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEnqRef.Location = New System.Drawing.Point(416, 24)
        Me.txtEnqRef.MaxLength = 50
        Me.txtEnqRef.Name = "txtEnqRef"
        Me.txtEnqRef.Size = New System.Drawing.Size(152, 20)
        Me.txtEnqRef.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(296, 24)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(120, 16)
        Me.Label9.TabIndex = 84
        Me.Label9.Text = "Enquiry Ref. No. && Dt."
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox3.Controls.Add(Me.GroupBoxPriceAvble)
        Me.GroupBox3.Controls.Add(Me.PanelProjectDetails)
        Me.GroupBox3.Controls.Add(Me.txtISR)
        Me.GroupBox3.Controls.Add(Me.Label61)
        Me.GroupBox3.Controls.Add(Me.GroupBox7)
        Me.GroupBox3.Controls.Add(Me.GroupBox6)
        Me.GroupBox3.Controls.Add(Me.DataGridEnquiryEdit)
        Me.GroupBox3.Controls.Add(Me.DataGridPartNumbers)
        Me.GroupBox3.Controls.Add(Me.DataGridCustomer)
        Me.GroupBox3.Controls.Add(Me.txtDunsno)
        Me.GroupBox3.Controls.Add(Me.Label56)
        Me.GroupBox3.Controls.Add(Me.GroupBox5)
        Me.GroupBox3.Controls.Add(Me.Label55)
        Me.GroupBox3.Controls.Add(Me.txtDept)
        Me.GroupBox3.Controls.Add(Me.txtcustintcode)
        Me.GroupBox3.Controls.Add(Me.ComboBoxStatusRemarks)
        Me.GroupBox3.Controls.Add(Me.txtSpecial)
        Me.GroupBox3.Controls.Add(Me.Label52)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.ComboBoxRejectionReasons)
        Me.GroupBox3.Controls.Add(Me.Label50)
        Me.GroupBox3.Controls.Add(Me.txtRejected)
        Me.GroupBox3.Controls.Add(Me.Label49)
        Me.GroupBox3.Controls.Add(Me.txtPartNot)
        Me.GroupBox3.Controls.Add(Me.Label48)
        Me.GroupBox3.Controls.Add(Me.Label47)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.txtPriceNot)
        Me.GroupBox3.Controls.Add(Me.txtPartYesPriceYes)
        Me.GroupBox3.Controls.Add(Me.txtTotalItems)
        Me.GroupBox3.Controls.Add(Me.btnUpload)
        Me.GroupBox3.Controls.Add(Me.txtDocDetails)
        Me.GroupBox3.Controls.Add(Me.Label34)
        Me.GroupBox3.Controls.Add(Me.btnHeaderSave)
        Me.GroupBox3.Controls.Add(Me.Label33)
        Me.GroupBox3.Controls.Add(Me.ComboBoxEnquiryStatus)
        Me.GroupBox3.Controls.Add(Me.ComboBoxClarity)
        Me.GroupBox3.Controls.Add(Me.Label32)
        Me.GroupBox3.Controls.Add(Me.Label31)
        Me.GroupBox3.Controls.Add(Me.ComboBoxEnquiryType)
        Me.GroupBox3.Controls.Add(Me.Label30)
        Me.GroupBox3.Controls.Add(Me.Label29)
        Me.GroupBox3.Controls.Add(Me.Label28)
        Me.GroupBox3.Controls.Add(Me.ComboBoxCategory)
        Me.GroupBox3.Controls.Add(Me.ComboBoxPriceStatus)
        Me.GroupBox3.Controls.Add(Me.ComboBoxForward)
        Me.GroupBox3.Controls.Add(Me.txtRemarks)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.Label27)
        Me.GroupBox3.Controls.Add(Me.txtCustID)
        Me.GroupBox3.Controls.Add(Me.txtVat)
        Me.GroupBox3.Controls.Add(Me.Label24)
        Me.GroupBox3.Controls.Add(Me.txtCst)
        Me.GroupBox3.Controls.Add(Me.Label25)
        Me.GroupBox3.Controls.Add(Me.txtEcc)
        Me.GroupBox3.Controls.Add(Me.Label26)
        Me.GroupBox3.Controls.Add(Me.txtFax)
        Me.GroupBox3.Controls.Add(Me.Label23)
        Me.GroupBox3.Controls.Add(Me.txtemail)
        Me.GroupBox3.Controls.Add(Me.Label22)
        Me.GroupBox3.Controls.Add(Me.txtPhone)
        Me.GroupBox3.Controls.Add(Me.Label21)
        Me.GroupBox3.Controls.Add(Me.txtMobile)
        Me.GroupBox3.Controls.Add(Me.Label20)
        Me.GroupBox3.Controls.Add(Me.txtDesignation)
        Me.GroupBox3.Controls.Add(Me.Label19)
        Me.GroupBox3.Controls.Add(Me.Label18)
        Me.GroupBox3.Controls.Add(Me.txtCustCountry)
        Me.GroupBox3.Controls.Add(Me.txtContact)
        Me.GroupBox3.Controls.Add(Me.txtCustPin)
        Me.GroupBox3.Controls.Add(Me.txtCustState)
        Me.GroupBox3.Controls.Add(Me.Label17)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Controls.Add(Me.txtCustcity)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.txtCustAdr3)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Controls.Add(Me.txtCustAd1)
        Me.GroupBox3.Controls.Add(Me.txtCustAdr2)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.Label11)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.txtCustomer)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.DTPStatusDt)
        Me.GroupBox3.Controls.Add(Me.GroupYesNo)
        Me.GroupBox3.Controls.Add(Me.txtBothNot)
        Me.GroupBox3.Controls.Add(Me.Label58)
        Me.GroupBox3.Controls.Add(Me.ComboBoxTax)
        Me.GroupBox3.Controls.Add(Me.Label62)
        Me.GroupBox3.Controls.Add(Me.ComboboxTSSISeg)
        Me.GroupBox3.Controls.Add(Me.ComboboxISR)
        Me.GroupBox3.Controls.Add(Me.ComboBoxCurrency)
        Me.GroupBox3.Controls.Add(Me.ComboBoxClass3)
        Me.GroupBox3.Controls.Add(Me.ComboboxClass)
        Me.GroupBox3.Controls.Add(Me.ComboBoxCSR)
        Me.GroupBox3.Controls.Add(Me.ComboboxSegment)
        Me.GroupBox3.Controls.Add(Me.Label65)
        Me.GroupBox3.Controls.Add(Me.Label63)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox3.Location = New System.Drawing.Point(16, 67)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1248, 341)
        Me.GroupBox3.TabIndex = 86
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Enquiry Customer Details"
        '
        'PanelProjectDetails
        '
        Me.PanelProjectDetails.BackColor = System.Drawing.Color.LightCyan
        Me.PanelProjectDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelProjectDetails.Controls.Add(Me.Label76)
        Me.PanelProjectDetails.Controls.Add(Me.txt)
        Me.PanelProjectDetails.Controls.Add(Me.Label75)
        Me.PanelProjectDetails.Controls.Add(Me.CheckedListBoxComp)
        Me.PanelProjectDetails.Controls.Add(Me.txtBusinessPotential)
        Me.PanelProjectDetails.Controls.Add(Me.Label74)
        Me.PanelProjectDetails.Controls.Add(Me.dtpProjectStDate)
        Me.PanelProjectDetails.Controls.Add(Me.Label73)
        Me.PanelProjectDetails.Controls.Add(Me.txtfax1)
        Me.PanelProjectDetails.Controls.Add(Me.Label69)
        Me.PanelProjectDetails.Controls.Add(Me.txtemail1)
        Me.PanelProjectDetails.Controls.Add(Me.Label70)
        Me.PanelProjectDetails.Controls.Add(Me.txtph1)
        Me.PanelProjectDetails.Controls.Add(Me.Label71)
        Me.PanelProjectDetails.Controls.Add(Me.txtMob1)
        Me.PanelProjectDetails.Controls.Add(Me.Label72)
        Me.PanelProjectDetails.Controls.Add(Me.TxtBuyerName1)
        Me.PanelProjectDetails.Controls.Add(Me.Label66)
        Me.PanelProjectDetails.Controls.Add(Me.txtDept1)
        Me.PanelProjectDetails.Controls.Add(Me.txtDesig1)
        Me.PanelProjectDetails.Controls.Add(Me.Label67)
        Me.PanelProjectDetails.Controls.Add(Me.Label68)
        Me.PanelProjectDetails.Location = New System.Drawing.Point(413, 9)
        Me.PanelProjectDetails.Name = "PanelProjectDetails"
        Me.PanelProjectDetails.Size = New System.Drawing.Size(56, 23)
        Me.PanelProjectDetails.TabIndex = 178
        Me.PanelProjectDetails.Visible = False
        '
        'Label76
        '
        Me.Label76.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label76.ForeColor = System.Drawing.Color.Red
        Me.Label76.Location = New System.Drawing.Point(13, 3)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(55, 34)
        Me.Label76.TabIndex = 182
        Me.Label76.Text = "More Info on Project"
        '
        'txt
        '
        Me.txt.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.txt.Location = New System.Drawing.Point(784, 4)
        Me.txt.Name = "txt"
        Me.txt.Size = New System.Drawing.Size(21, 23)
        Me.txt.TabIndex = 181
        Me.txt.Text = "X"
        Me.txt.UseVisualStyleBackColor = True
        '
        'Label75
        '
        Me.Label75.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label75.ForeColor = System.Drawing.Color.Black
        Me.Label75.Location = New System.Drawing.Point(4, 99)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(80, 24)
        Me.Label75.TabIndex = 180
        Me.Label75.Text = "Comp. Name"
        '
        'CheckedListBoxComp
        '
        Me.CheckedListBoxComp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.CheckedListBoxComp.Location = New System.Drawing.Point(119, 101)
        Me.CheckedListBoxComp.Name = "CheckedListBoxComp"
        Me.CheckedListBoxComp.Size = New System.Drawing.Size(233, 49)
        Me.CheckedListBoxComp.TabIndex = 179
        '
        'txtBusinessPotential
        '
        Me.txtBusinessPotential.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBusinessPotential.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBusinessPotential.Location = New System.Drawing.Point(119, 72)
        Me.txtBusinessPotential.Name = "txtBusinessPotential"
        Me.txtBusinessPotential.Size = New System.Drawing.Size(233, 20)
        Me.txtBusinessPotential.TabIndex = 178
        '
        'Label74
        '
        Me.Label74.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label74.ForeColor = System.Drawing.Color.Black
        Me.Label74.Location = New System.Drawing.Point(5, 73)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(132, 24)
        Me.Label74.TabIndex = 177
        Me.Label74.Text = "Business Potential"
        '
        'dtpProjectStDate
        '
        Me.dtpProjectStDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpProjectStDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpProjectStDate.Location = New System.Drawing.Point(119, 45)
        Me.dtpProjectStDate.Name = "dtpProjectStDate"
        Me.dtpProjectStDate.Size = New System.Drawing.Size(174, 20)
        Me.dtpProjectStDate.TabIndex = 176
        '
        'Label73
        '
        Me.Label73.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label73.ForeColor = System.Drawing.Color.Black
        Me.Label73.Location = New System.Drawing.Point(5, 47)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(115, 24)
        Me.Label73.TabIndex = 175
        Me.Label73.Text = "Project Start Dt"
        '
        'txtfax1
        '
        Me.txtfax1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtfax1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtfax1.Location = New System.Drawing.Point(436, 172)
        Me.txtfax1.MaxLength = 50
        Me.txtfax1.Name = "txtfax1"
        Me.txtfax1.Size = New System.Drawing.Size(344, 20)
        Me.txtfax1.TabIndex = 169
        '
        'Label69
        '
        Me.Label69.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label69.ForeColor = System.Drawing.Color.Black
        Me.Label69.Location = New System.Drawing.Point(359, 173)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(56, 24)
        Me.Label69.TabIndex = 174
        Me.Label69.Text = "Fax"
        '
        'txtemail1
        '
        Me.txtemail1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtemail1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtemail1.Location = New System.Drawing.Point(436, 196)
        Me.txtemail1.MaxLength = 70
        Me.txtemail1.Name = "txtemail1"
        Me.txtemail1.Size = New System.Drawing.Size(344, 20)
        Me.txtemail1.TabIndex = 170
        '
        'Label70
        '
        Me.Label70.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label70.ForeColor = System.Drawing.Color.Black
        Me.Label70.Location = New System.Drawing.Point(359, 195)
        Me.Label70.Name = "Label70"
        Me.Label70.Size = New System.Drawing.Size(64, 21)
        Me.Label70.TabIndex = 173
        Me.Label70.Text = "Email"
        '
        'txtph1
        '
        Me.txtph1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtph1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtph1.Location = New System.Drawing.Point(436, 148)
        Me.txtph1.MaxLength = 50
        Me.txtph1.Name = "txtph1"
        Me.txtph1.Size = New System.Drawing.Size(344, 20)
        Me.txtph1.TabIndex = 168
        '
        'Label71
        '
        Me.Label71.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label71.ForeColor = System.Drawing.Color.Black
        Me.Label71.Location = New System.Drawing.Point(357, 145)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(56, 24)
        Me.Label71.TabIndex = 172
        Me.Label71.Text = "Phone"
        '
        'txtMob1
        '
        Me.txtMob1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMob1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMob1.Location = New System.Drawing.Point(-147, 71)
        Me.txtMob1.MaxLength = 50
        Me.txtMob1.Name = "txtMob1"
        Me.txtMob1.Size = New System.Drawing.Size(344, 20)
        Me.txtMob1.TabIndex = 167
        '
        'Label72
        '
        Me.Label72.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label72.ForeColor = System.Drawing.Color.Black
        Me.Label72.Location = New System.Drawing.Point(359, 121)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(48, 16)
        Me.Label72.TabIndex = 171
        Me.Label72.Text = "Mobile"
        '
        'TxtBuyerName1
        '
        Me.TxtBuyerName1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtBuyerName1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtBuyerName1.Location = New System.Drawing.Point(436, 45)
        Me.TxtBuyerName1.Name = "TxtBuyerName1"
        Me.TxtBuyerName1.Size = New System.Drawing.Size(344, 20)
        Me.TxtBuyerName1.TabIndex = 166
        '
        'Label66
        '
        Me.Label66.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label66.ForeColor = System.Drawing.Color.Black
        Me.Label66.Location = New System.Drawing.Point(359, 93)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(32, 24)
        Me.Label66.TabIndex = 165
        Me.Label66.Text = "Dept"
        '
        'txtDept1
        '
        Me.txtDept1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDept1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDept1.Location = New System.Drawing.Point(436, 97)
        Me.txtDept1.MaxLength = 60
        Me.txtDept1.Name = "txtDept1"
        Me.txtDept1.Size = New System.Drawing.Size(344, 20)
        Me.txtDept1.TabIndex = 162
        '
        'txtDesig1
        '
        Me.txtDesig1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDesig1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesig1.Location = New System.Drawing.Point(436, 72)
        Me.txtDesig1.MaxLength = 60
        Me.txtDesig1.Name = "txtDesig1"
        Me.txtDesig1.Size = New System.Drawing.Size(344, 20)
        Me.txtDesig1.TabIndex = 161
        '
        'Label67
        '
        Me.Label67.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label67.ForeColor = System.Drawing.Color.Black
        Me.Label67.Location = New System.Drawing.Point(359, 69)
        Me.Label67.Name = "Label67"
        Me.Label67.Size = New System.Drawing.Size(80, 24)
        Me.Label67.TabIndex = 164
        Me.Label67.Text = "Designation"
        '
        'Label68
        '
        Me.Label68.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label68.ForeColor = System.Drawing.Color.Black
        Me.Label68.Location = New System.Drawing.Point(359, 45)
        Me.Label68.Name = "Label68"
        Me.Label68.Size = New System.Drawing.Size(80, 24)
        Me.Label68.TabIndex = 163
        Me.Label68.Text = "Buyer Name"
        '
        'txtISR
        '
        Me.txtISR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtISR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtISR.Location = New System.Drawing.Point(617, 321)
        Me.txtISR.MaxLength = 4
        Me.txtISR.Name = "txtISR"
        Me.txtISR.Size = New System.Drawing.Size(50, 20)
        Me.txtISR.TabIndex = 177
        Me.txtISR.Visible = False
        '
        'Label61
        '
        Me.Label61.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label61.ForeColor = System.Drawing.Color.Black
        Me.Label61.Location = New System.Drawing.Point(421, 263)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(60, 24)
        Me.Label61.TabIndex = 169
        Me.Label61.Text = "Tax"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.RBCustomerExisting)
        Me.GroupBox7.Controls.Add(Me.RBCustomerNew)
        Me.GroupBox7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox7.ForeColor = System.Drawing.Color.Red
        Me.GroupBox7.Location = New System.Drawing.Point(11, 16)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(126, 32)
        Me.GroupBox7.TabIndex = 168
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Customer"
        '
        'RBCustomerExisting
        '
        Me.RBCustomerExisting.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBCustomerExisting.ForeColor = System.Drawing.Color.Red
        Me.RBCustomerExisting.Location = New System.Drawing.Point(57, 12)
        Me.RBCustomerExisting.Name = "RBCustomerExisting"
        Me.RBCustomerExisting.Size = New System.Drawing.Size(63, 17)
        Me.RBCustomerExisting.TabIndex = 13
        Me.RBCustomerExisting.Text = "Existing"
        '
        'RBCustomerNew
        '
        Me.RBCustomerNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBCustomerNew.ForeColor = System.Drawing.Color.Red
        Me.RBCustomerNew.Location = New System.Drawing.Point(5, 13)
        Me.RBCustomerNew.Name = "RBCustomerNew"
        Me.RBCustomerNew.Size = New System.Drawing.Size(56, 16)
        Me.RBCustomerNew.TabIndex = 12
        Me.RBCustomerNew.Text = "New"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.RadioButtonDomestic)
        Me.GroupBox6.Controls.Add(Me.RadioButtonExport)
        Me.GroupBox6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox6.ForeColor = System.Drawing.Color.Red
        Me.GroupBox6.Location = New System.Drawing.Point(156, 9)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(140, 36)
        Me.GroupBox6.TabIndex = 167
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Market Type"
        '
        'RadioButtonDomestic
        '
        Me.RadioButtonDomestic.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonDomestic.ForeColor = System.Drawing.Color.Red
        Me.RadioButtonDomestic.Location = New System.Drawing.Point(5, 12)
        Me.RadioButtonDomestic.Name = "RadioButtonDomestic"
        Me.RadioButtonDomestic.Size = New System.Drawing.Size(72, 20)
        Me.RadioButtonDomestic.TabIndex = 52
        Me.RadioButtonDomestic.Text = "Domestic"
        '
        'RadioButtonExport
        '
        Me.RadioButtonExport.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonExport.ForeColor = System.Drawing.Color.Red
        Me.RadioButtonExport.Location = New System.Drawing.Point(79, 11)
        Me.RadioButtonExport.Name = "RadioButtonExport"
        Me.RadioButtonExport.Size = New System.Drawing.Size(57, 22)
        Me.RadioButtonExport.TabIndex = 53
        Me.RadioButtonExport.Text = "Export"
        '
        'DataGridEnquiryEdit
        '
        Me.DataGridEnquiryEdit.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridEnquiryEdit.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridEnquiryEdit.CaptionVisible = False
        Me.DataGridEnquiryEdit.DataMember = ""
        Me.DataGridEnquiryEdit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridEnquiryEdit.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridEnquiryEdit.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridEnquiryEdit.Location = New System.Drawing.Point(601, 12)
        Me.DataGridEnquiryEdit.Name = "DataGridEnquiryEdit"
        Me.DataGridEnquiryEdit.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridEnquiryEdit.ParentRowsVisible = False
        Me.DataGridEnquiryEdit.PreferredColumnWidth = 85
        Me.DataGridEnquiryEdit.ReadOnly = True
        Me.DataGridEnquiryEdit.RowHeadersVisible = False
        Me.DataGridEnquiryEdit.Size = New System.Drawing.Size(11, 24)
        Me.DataGridEnquiryEdit.TabIndex = 155
        Me.DataGridEnquiryEdit.Visible = False
        '
        'DataGridPartNumbers
        '
        Me.DataGridPartNumbers.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridPartNumbers.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridPartNumbers.CaptionVisible = False
        Me.DataGridPartNumbers.DataMember = ""
        Me.DataGridPartNumbers.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridPartNumbers.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridPartNumbers.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridPartNumbers.Location = New System.Drawing.Point(588, 12)
        Me.DataGridPartNumbers.Name = "DataGridPartNumbers"
        Me.DataGridPartNumbers.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridPartNumbers.ParentRowsVisible = False
        Me.DataGridPartNumbers.PreferredColumnWidth = 85
        Me.DataGridPartNumbers.ReadOnly = True
        Me.DataGridPartNumbers.RowHeadersVisible = False
        Me.DataGridPartNumbers.Size = New System.Drawing.Size(8, 24)
        Me.DataGridPartNumbers.TabIndex = 154
        Me.DataGridPartNumbers.Visible = False
        '
        'txtDunsno
        '
        Me.txtDunsno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDunsno.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.txtDunsno.Location = New System.Drawing.Point(488, 216)
        Me.txtDunsno.MaxLength = 50
        Me.txtDunsno.Name = "txtDunsno"
        Me.txtDunsno.Size = New System.Drawing.Size(296, 20)
        Me.txtDunsno.TabIndex = 31
        '
        'Label56
        '
        Me.Label56.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label56.ForeColor = System.Drawing.Color.Black
        Me.Label56.Location = New System.Drawing.Point(423, 216)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(65, 24)
        Me.Label56.TabIndex = 163
        Me.Label56.Text = "Duns No."
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.RadioButtonId)
        Me.GroupBox5.Controls.Add(Me.RadioButtonName)
        Me.GroupBox5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBox5.ForeColor = System.Drawing.Color.Red
        Me.GroupBox5.Location = New System.Drawing.Point(299, 12)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(126, 33)
        Me.GroupBox5.TabIndex = 162
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Customer Search By"
        '
        'RadioButtonId
        '
        Me.RadioButtonId.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonId.ForeColor = System.Drawing.Color.Red
        Me.RadioButtonId.Location = New System.Drawing.Point(13, 13)
        Me.RadioButtonId.Name = "RadioButtonId"
        Me.RadioButtonId.Size = New System.Drawing.Size(34, 16)
        Me.RadioButtonId.TabIndex = 52
        Me.RadioButtonId.Text = "Id"
        '
        'RadioButtonName
        '
        Me.RadioButtonName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonName.ForeColor = System.Drawing.Color.Red
        Me.RadioButtonName.Location = New System.Drawing.Point(53, 13)
        Me.RadioButtonName.Name = "RadioButtonName"
        Me.RadioButtonName.Size = New System.Drawing.Size(54, 16)
        Me.RadioButtonName.TabIndex = 53
        Me.RadioButtonName.Text = "Name"
        '
        'Label55
        '
        Me.Label55.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label55.ForeColor = System.Drawing.Color.Black
        Me.Label55.Location = New System.Drawing.Point(248, 240)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(32, 24)
        Me.Label55.TabIndex = 160
        Me.Label55.Text = "Dept"
        '
        'txtDept
        '
        Me.txtDept.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDept.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDept.Location = New System.Drawing.Point(280, 240)
        Me.txtDept.MaxLength = 60
        Me.txtDept.Name = "txtDept"
        Me.txtDept.Size = New System.Drawing.Size(136, 20)
        Me.txtDept.TabIndex = 24
        '
        'txtcustintcode
        '
        Me.txtcustintcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcustintcode.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcustintcode.Location = New System.Drawing.Point(617, 9)
        Me.txtcustintcode.Name = "txtcustintcode"
        Me.txtcustintcode.Size = New System.Drawing.Size(27, 22)
        Me.txtcustintcode.TabIndex = 153
        Me.txtcustintcode.Visible = False
        '
        'ComboBoxStatusRemarks
        '
        Me.ComboBoxStatusRemarks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxStatusRemarks.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxStatusRemarks.Location = New System.Drawing.Point(800, 176)
        Me.ComboBoxStatusRemarks.Name = "ComboBoxStatusRemarks"
        Me.ComboBoxStatusRemarks.Size = New System.Drawing.Size(144, 22)
        Me.ComboBoxStatusRemarks.TabIndex = 45
        '
        'txtSpecial
        '
        Me.txtSpecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSpecial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpecial.Location = New System.Drawing.Point(101, 272)
        Me.txtSpecial.MaxLength = 500
        Me.txtSpecial.Multiline = True
        Me.txtSpecial.Name = "txtSpecial"
        Me.txtSpecial.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSpecial.Size = New System.Drawing.Size(275, 56)
        Me.txtSpecial.TabIndex = 37
        '
        'Label52
        '
        Me.Label52.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label52.ForeColor = System.Drawing.Color.Black
        Me.Label52.Location = New System.Drawing.Point(8, 280)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(96, 40)
        Me.Label52.TabIndex = 151
        Me.Label52.Text = "Special Instruction for Pur/App.Dept."
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(800, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(160, 16)
        Me.Label6.TabIndex = 150
        Me.Label6.Text = "Reason for Rejection"
        '
        'ComboBoxRejectionReasons
        '
        Me.ComboBoxRejectionReasons.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxRejectionReasons.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxRejectionReasons.Location = New System.Drawing.Point(800, 224)
        Me.ComboBoxRejectionReasons.Name = "ComboBoxRejectionReasons"
        Me.ComboBoxRejectionReasons.Size = New System.Drawing.Size(144, 22)
        Me.ComboBoxRejectionReasons.TabIndex = 46
        '
        'Label50
        '
        Me.Label50.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label50.ForeColor = System.Drawing.Color.Black
        Me.Label50.Location = New System.Drawing.Point(1108, 181)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(77, 16)
        Me.Label50.TabIndex = 148
        Me.Label50.Text = "Rejected Items"
        '
        'txtRejected
        '
        Me.txtRejected.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRejected.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRejected.Location = New System.Drawing.Point(1192, 179)
        Me.txtRejected.MaxLength = 4
        Me.txtRejected.Name = "txtRejected"
        Me.txtRejected.Size = New System.Drawing.Size(40, 20)
        Me.txtRejected.TabIndex = 52
        '
        'Label49
        '
        Me.Label49.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label49.ForeColor = System.Drawing.Color.Black
        Me.Label49.Location = New System.Drawing.Point(1096, 121)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(96, 24)
        Me.Label49.TabIndex = 146
        Me.Label49.Text = "Part No. not avlbe"
        '
        'txtPartNot
        '
        Me.txtPartNot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartNot.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartNot.Location = New System.Drawing.Point(1192, 120)
        Me.txtPartNot.MaxLength = 4
        Me.txtPartNot.Name = "txtPartNot"
        Me.txtPartNot.Size = New System.Drawing.Size(40, 20)
        Me.txtPartNot.TabIndex = 50
        '
        'Label48
        '
        Me.Label48.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label48.ForeColor = System.Drawing.Color.Black
        Me.Label48.Location = New System.Drawing.Point(1101, 92)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(83, 16)
        Me.Label48.TabIndex = 144
        Me.Label48.Text = " Price not avble"
        '
        'Label47
        '
        Me.Label47.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label47.ForeColor = System.Drawing.Color.Black
        Me.Label47.Location = New System.Drawing.Point(1083, 57)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(109, 23)
        Me.Label47.TabIndex = 143
        Me.Label47.Text = "Part and Price avble"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(1096, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 16)
        Me.Label7.TabIndex = 142
        Me.Label7.Text = "Total No. of Items"
        '
        'txtPriceNot
        '
        Me.txtPriceNot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPriceNot.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPriceNot.Location = New System.Drawing.Point(1192, 88)
        Me.txtPriceNot.MaxLength = 4
        Me.txtPriceNot.Name = "txtPriceNot"
        Me.txtPriceNot.Size = New System.Drawing.Size(40, 20)
        Me.txtPriceNot.TabIndex = 49
        '
        'txtPartYesPriceYes
        '
        Me.txtPartYesPriceYes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartYesPriceYes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartYesPriceYes.Location = New System.Drawing.Point(1192, 56)
        Me.txtPartYesPriceYes.MaxLength = 4
        Me.txtPartYesPriceYes.Name = "txtPartYesPriceYes"
        Me.txtPartYesPriceYes.Size = New System.Drawing.Size(40, 20)
        Me.txtPartYesPriceYes.TabIndex = 48
        '
        'txtTotalItems
        '
        Me.txtTotalItems.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotalItems.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalItems.Location = New System.Drawing.Point(1192, 24)
        Me.txtTotalItems.MaxLength = 4
        Me.txtTotalItems.Name = "txtTotalItems"
        Me.txtTotalItems.Size = New System.Drawing.Size(40, 20)
        Me.txtTotalItems.TabIndex = 47
        '
        'btnUpload
        '
        Me.btnUpload.BackColor = System.Drawing.Color.LightGray
        Me.btnUpload.Enabled = False
        Me.btnUpload.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpload.ForeColor = System.Drawing.Color.Blue
        Me.btnUpload.Location = New System.Drawing.Point(1144, 304)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(80, 24)
        Me.btnUpload.TabIndex = 56
        Me.btnUpload.Text = "Upload"
        Me.btnUpload.UseVisualStyleBackColor = False
        '
        'txtDocDetails
        '
        Me.txtDocDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDocDetails.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDocDetails.Location = New System.Drawing.Point(800, 271)
        Me.txtDocDetails.MaxLength = 500
        Me.txtDocDetails.Multiline = True
        Me.txtDocDetails.Name = "txtDocDetails"
        Me.txtDocDetails.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDocDetails.Size = New System.Drawing.Size(336, 56)
        Me.txtDocDetails.TabIndex = 54
        '
        'Label34
        '
        Me.Label34.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.ForeColor = System.Drawing.Color.Black
        Me.Label34.Location = New System.Drawing.Point(800, 256)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(112, 24)
        Me.Label34.TabIndex = 133
        Me.Label34.Text = "Document Details"
        '
        'btnHeaderSave
        '
        Me.btnHeaderSave.BackColor = System.Drawing.Color.LightGray
        Me.btnHeaderSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHeaderSave.ForeColor = System.Drawing.Color.Blue
        Me.btnHeaderSave.Location = New System.Drawing.Point(1144, 272)
        Me.btnHeaderSave.Name = "btnHeaderSave"
        Me.btnHeaderSave.Size = New System.Drawing.Size(80, 24)
        Me.btnHeaderSave.TabIndex = 55
        Me.btnHeaderSave.Text = "Save"
        Me.btnHeaderSave.UseVisualStyleBackColor = False
        '
        'Label33
        '
        Me.Label33.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label33.ForeColor = System.Drawing.Color.Black
        Me.Label33.Location = New System.Drawing.Point(952, 112)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(96, 16)
        Me.Label33.TabIndex = 131
        Me.Label33.Text = "Enquiry Status"
        '
        'ComboBoxEnquiryStatus
        '
        Me.ComboBoxEnquiryStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxEnquiryStatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxEnquiryStatus.Location = New System.Drawing.Point(952, 128)
        Me.ComboBoxEnquiryStatus.Name = "ComboBoxEnquiryStatus"
        Me.ComboBoxEnquiryStatus.Size = New System.Drawing.Size(136, 22)
        Me.ComboBoxEnquiryStatus.TabIndex = 43
        '
        'ComboBoxClarity
        '
        Me.ComboBoxClarity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxClarity.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxClarity.Location = New System.Drawing.Point(800, 80)
        Me.ComboBoxClarity.Name = "ComboBoxClarity"
        Me.ComboBoxClarity.Size = New System.Drawing.Size(144, 22)
        Me.ComboBoxClarity.TabIndex = 40
        '
        'Label32
        '
        Me.Label32.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.ForeColor = System.Drawing.Color.Black
        Me.Label32.Location = New System.Drawing.Point(800, 64)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(96, 24)
        Me.Label32.TabIndex = 128
        Me.Label32.Text = "Enquiry Clarity"
        '
        'Label31
        '
        Me.Label31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.ForeColor = System.Drawing.Color.Black
        Me.Label31.Location = New System.Drawing.Point(952, 16)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(88, 16)
        Me.Label31.TabIndex = 127
        Me.Label31.Text = "Enquiry Type"
        '
        'ComboBoxEnquiryType
        '
        Me.ComboBoxEnquiryType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxEnquiryType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxEnquiryType.Location = New System.Drawing.Point(952, 32)
        Me.ComboBoxEnquiryType.Name = "ComboBoxEnquiryType"
        Me.ComboBoxEnquiryType.Size = New System.Drawing.Size(136, 22)
        Me.ComboBoxEnquiryType.TabIndex = 39
        '
        'Label30
        '
        Me.Label30.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.ForeColor = System.Drawing.Color.Black
        Me.Label30.Location = New System.Drawing.Point(800, 112)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(88, 16)
        Me.Label30.TabIndex = 125
        Me.Label30.Text = "Price Status"
        '
        'Label29
        '
        Me.Label29.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.Color.Black
        Me.Label29.Location = New System.Drawing.Point(952, 64)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(88, 16)
        Me.Label29.TabIndex = 124
        Me.Label29.Text = "Quote  Forward"
        '
        'Label28
        '
        Me.Label28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.Color.Black
        Me.Label28.Location = New System.Drawing.Point(800, 16)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(112, 16)
        Me.Label28.TabIndex = 123
        Me.Label28.Text = "Enquiry Category"
        '
        'ComboBoxCategory
        '
        Me.ComboBoxCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxCategory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxCategory.Location = New System.Drawing.Point(800, 32)
        Me.ComboBoxCategory.Name = "ComboBoxCategory"
        Me.ComboBoxCategory.Size = New System.Drawing.Size(144, 22)
        Me.ComboBoxCategory.TabIndex = 38
        '
        'ComboBoxPriceStatus
        '
        Me.ComboBoxPriceStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPriceStatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxPriceStatus.Location = New System.Drawing.Point(800, 128)
        Me.ComboBoxPriceStatus.Name = "ComboBoxPriceStatus"
        Me.ComboBoxPriceStatus.Size = New System.Drawing.Size(144, 22)
        Me.ComboBoxPriceStatus.TabIndex = 42
        '
        'ComboBoxForward
        '
        Me.ComboBoxForward.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxForward.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxForward.Location = New System.Drawing.Point(952, 80)
        Me.ComboBoxForward.Name = "ComboBoxForward"
        Me.ComboBoxForward.Size = New System.Drawing.Size(136, 22)
        Me.ComboBoxForward.TabIndex = 41
        '
        'txtRemarks
        '
        Me.txtRemarks.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRemarks.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRemarks.Location = New System.Drawing.Point(488, 240)
        Me.txtRemarks.MaxLength = 150
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(296, 20)
        Me.txtRemarks.TabIndex = 32
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(421, 240)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 24)
        Me.Label8.TabIndex = 118
        Me.Label8.Text = "PayTerms"
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.Black
        Me.Label27.Location = New System.Drawing.Point(800, 160)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(88, 16)
        Me.Label27.TabIndex = 116
        Me.Label27.Text = "Status Remarks"
        '
        'txtCustID
        '
        Me.txtCustID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustID.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustID.Location = New System.Drawing.Point(432, 16)
        Me.txtCustID.MaxLength = 30
        Me.txtCustID.Name = "txtCustID"
        Me.txtCustID.Size = New System.Drawing.Size(136, 20)
        Me.txtCustID.TabIndex = 12
        '
        'txtVat
        '
        Me.txtVat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtVat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVat.Location = New System.Drawing.Point(488, 168)
        Me.txtVat.MaxLength = 50
        Me.txtVat.Name = "txtVat"
        Me.txtVat.Size = New System.Drawing.Size(296, 20)
        Me.txtVat.TabIndex = 29
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.Black
        Me.Label24.Location = New System.Drawing.Point(424, 168)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(56, 24)
        Me.Label24.TabIndex = 111
        Me.Label24.Text = "Vat No."
        '
        'txtCst
        '
        Me.txtCst.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCst.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCst.Location = New System.Drawing.Point(488, 192)
        Me.txtCst.MaxLength = 50
        Me.txtCst.Name = "txtCst"
        Me.txtCst.Size = New System.Drawing.Size(296, 20)
        Me.txtCst.TabIndex = 30
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.Black
        Me.Label25.Location = New System.Drawing.Point(424, 192)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(56, 24)
        Me.Label25.TabIndex = 109
        Me.Label25.Text = "Cst No."
        '
        'txtEcc
        '
        Me.txtEcc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEcc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEcc.Location = New System.Drawing.Point(488, 144)
        Me.txtEcc.MaxLength = 50
        Me.txtEcc.Name = "txtEcc"
        Me.txtEcc.Size = New System.Drawing.Size(296, 20)
        Me.txtEcc.TabIndex = 28
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Black
        Me.Label26.Location = New System.Drawing.Point(424, 144)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(48, 24)
        Me.Label26.TabIndex = 107
        Me.Label26.Text = "Ecc No."
        '
        'txtFax
        '
        Me.txtFax.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFax.Location = New System.Drawing.Point(488, 96)
        Me.txtFax.MaxLength = 50
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Size = New System.Drawing.Size(296, 20)
        Me.txtFax.TabIndex = 26
        '
        'Label23
        '
        Me.Label23.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.Black
        Me.Label23.Location = New System.Drawing.Point(424, 96)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(56, 24)
        Me.Label23.TabIndex = 101
        Me.Label23.Text = "Fax"
        '
        'txtemail
        '
        Me.txtemail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtemail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtemail.Location = New System.Drawing.Point(488, 120)
        Me.txtemail.MaxLength = 70
        Me.txtemail.Name = "txtemail"
        Me.txtemail.Size = New System.Drawing.Size(296, 20)
        Me.txtemail.TabIndex = 27
        '
        'Label22
        '
        Me.Label22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Black
        Me.Label22.Location = New System.Drawing.Point(424, 123)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(64, 21)
        Me.Label22.TabIndex = 99
        Me.Label22.Text = "Email"
        '
        'txtPhone
        '
        Me.txtPhone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPhone.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhone.Location = New System.Drawing.Point(488, 72)
        Me.txtPhone.MaxLength = 50
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(296, 20)
        Me.txtPhone.TabIndex = 25
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Black
        Me.Label21.Location = New System.Drawing.Point(424, 72)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(56, 24)
        Me.Label21.TabIndex = 97
        Me.Label21.Text = "Phone"
        '
        'txtMobile
        '
        Me.txtMobile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMobile.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMobile.Location = New System.Drawing.Point(488, 48)
        Me.txtMobile.MaxLength = 50
        Me.txtMobile.Name = "txtMobile"
        Me.txtMobile.Size = New System.Drawing.Size(296, 20)
        Me.txtMobile.TabIndex = 24
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Black
        Me.Label20.Location = New System.Drawing.Point(421, 49)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(48, 16)
        Me.Label20.TabIndex = 95
        Me.Label20.Text = "Mobile"
        '
        'txtDesignation
        '
        Me.txtDesignation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDesignation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesignation.Location = New System.Drawing.Point(88, 240)
        Me.txtDesignation.MaxLength = 60
        Me.txtDesignation.Name = "txtDesignation"
        Me.txtDesignation.Size = New System.Drawing.Size(152, 20)
        Me.txtDesignation.TabIndex = 23
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Black
        Me.Label19.Location = New System.Drawing.Point(8, 240)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(80, 24)
        Me.Label19.TabIndex = 93
        Me.Label19.Text = "Designation"
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Black
        Me.Label18.Location = New System.Drawing.Point(8, 192)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(64, 24)
        Me.Label18.TabIndex = 92
        Me.Label18.Text = "Country"
        '
        'txtCustCountry
        '
        Me.txtCustCountry.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustCountry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustCountry.Location = New System.Drawing.Point(88, 192)
        Me.txtCustCountry.MaxLength = 60
        Me.txtCustCountry.Name = "txtCustCountry"
        Me.txtCustCountry.Size = New System.Drawing.Size(328, 20)
        Me.txtCustCountry.TabIndex = 21
        '
        'txtContact
        '
        Me.txtContact.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtContact.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtContact.Location = New System.Drawing.Point(88, 216)
        Me.txtContact.Name = "txtContact"
        Me.txtContact.Size = New System.Drawing.Size(328, 20)
        Me.txtContact.TabIndex = 22
        '
        'txtCustPin
        '
        Me.txtCustPin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustPin.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustPin.Location = New System.Drawing.Point(312, 144)
        Me.txtCustPin.Name = "txtCustPin"
        Me.txtCustPin.Size = New System.Drawing.Size(104, 20)
        Me.txtCustPin.TabIndex = 19
        '
        'txtCustState
        '
        Me.txtCustState.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustState.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustState.Location = New System.Drawing.Point(88, 168)
        Me.txtCustState.MaxLength = 60
        Me.txtCustState.Name = "txtCustState"
        Me.txtCustState.Size = New System.Drawing.Size(328, 20)
        Me.txtCustState.TabIndex = 20
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Black
        Me.Label17.Location = New System.Drawing.Point(8, 168)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(64, 24)
        Me.Label17.TabIndex = 87
        Me.Label17.Text = "State"
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(288, 144)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(24, 24)
        Me.Label16.TabIndex = 86
        Me.Label16.Text = "ZIP"
        '
        'txtCustcity
        '
        Me.txtCustcity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustcity.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustcity.Location = New System.Drawing.Point(88, 144)
        Me.txtCustcity.MaxLength = 60
        Me.txtCustcity.Name = "txtCustcity"
        Me.txtCustcity.Size = New System.Drawing.Size(200, 20)
        Me.txtCustcity.TabIndex = 18
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(8, 144)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(64, 24)
        Me.Label15.TabIndex = 84
        Me.Label15.Text = "City"
        '
        'txtCustAdr3
        '
        Me.txtCustAdr3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustAdr3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustAdr3.Location = New System.Drawing.Point(88, 120)
        Me.txtCustAdr3.MaxLength = 60
        Me.txtCustAdr3.Name = "txtCustAdr3"
        Me.txtCustAdr3.Size = New System.Drawing.Size(328, 20)
        Me.txtCustAdr3.TabIndex = 17
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Black
        Me.Label14.Location = New System.Drawing.Point(8, 120)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(64, 24)
        Me.Label14.TabIndex = 82
        Me.Label14.Text = "Address3"
        '
        'txtCustAd1
        '
        Me.txtCustAd1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustAd1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustAd1.Location = New System.Drawing.Point(88, 72)
        Me.txtCustAd1.MaxLength = 60
        Me.txtCustAd1.Name = "txtCustAd1"
        Me.txtCustAd1.Size = New System.Drawing.Size(328, 20)
        Me.txtCustAd1.TabIndex = 15
        '
        'txtCustAdr2
        '
        Me.txtCustAdr2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustAdr2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustAdr2.Location = New System.Drawing.Point(88, 96)
        Me.txtCustAdr2.MaxLength = 60
        Me.txtCustAdr2.Name = "txtCustAdr2"
        Me.txtCustAdr2.Size = New System.Drawing.Size(328, 20)
        Me.txtCustAdr2.TabIndex = 16
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(8, 216)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(80, 24)
        Me.Label13.TabIndex = 79
        Me.Label13.Text = "Buyer Name"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(8, 96)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(64, 24)
        Me.Label12.TabIndex = 78
        Me.Label12.Text = "Address2"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(8, 72)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 16)
        Me.Label11.TabIndex = 77
        Me.Label11.Text = "Address1"
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(8, 52)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 20)
        Me.Label10.TabIndex = 76
        Me.Label10.Text = "Name"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(952, 160)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 16)
        Me.Label3.TabIndex = 87
        Me.Label3.Text = "Status Dt."
        '
        'DTPStatusDt
        '
        Me.DTPStatusDt.AllowDrop = True
        Me.DTPStatusDt.Checked = False
        Me.DTPStatusDt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTPStatusDt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTPStatusDt.Location = New System.Drawing.Point(952, 176)
        Me.DTPStatusDt.Name = "DTPStatusDt"
        Me.DTPStatusDt.ShowCheckBox = True
        Me.DTPStatusDt.Size = New System.Drawing.Size(136, 20)
        Me.DTPStatusDt.TabIndex = 44
        '
        'GroupYesNo
        '
        Me.GroupYesNo.Controls.Add(Me.rbdocyes)
        Me.GroupYesNo.Controls.Add(Me.rbDocNo)
        Me.GroupYesNo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupYesNo.ForeColor = System.Drawing.Color.Black
        Me.GroupYesNo.Location = New System.Drawing.Point(968, 216)
        Me.GroupYesNo.Name = "GroupYesNo"
        Me.GroupYesNo.Size = New System.Drawing.Size(136, 40)
        Me.GroupYesNo.TabIndex = 152
        Me.GroupYesNo.TabStop = False
        Me.GroupYesNo.Text = "Document Uploaded"
        '
        'rbdocyes
        '
        Me.rbdocyes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbdocyes.ForeColor = System.Drawing.Color.Red
        Me.rbdocyes.Location = New System.Drawing.Point(16, 16)
        Me.rbdocyes.Name = "rbdocyes"
        Me.rbdocyes.Size = New System.Drawing.Size(48, 16)
        Me.rbdocyes.TabIndex = 52
        Me.rbdocyes.Text = "Yes"
        '
        'rbDocNo
        '
        Me.rbDocNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbDocNo.ForeColor = System.Drawing.Color.Red
        Me.rbDocNo.Location = New System.Drawing.Point(80, 16)
        Me.rbDocNo.Name = "rbDocNo"
        Me.rbDocNo.Size = New System.Drawing.Size(40, 16)
        Me.rbDocNo.TabIndex = 53
        Me.rbDocNo.Text = "No"
        '
        'txtBothNot
        '
        Me.txtBothNot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBothNot.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBothNot.Location = New System.Drawing.Point(1192, 149)
        Me.txtBothNot.MaxLength = 4
        Me.txtBothNot.Name = "txtBothNot"
        Me.txtBothNot.Size = New System.Drawing.Size(40, 20)
        Me.txtBothNot.TabIndex = 51
        '
        'Label58
        '
        Me.Label58.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label58.ForeColor = System.Drawing.Color.Black
        Me.Label58.Location = New System.Drawing.Point(1105, 152)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(80, 17)
        Me.Label58.TabIndex = 166
        Me.Label58.Text = "Both Not Avble"
        '
        'ComboBoxTax
        '
        Me.ComboBoxTax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxTax.Location = New System.Drawing.Point(488, 265)
        Me.ComboBoxTax.Name = "ComboBoxTax"
        Me.ComboBoxTax.Size = New System.Drawing.Size(296, 22)
        Me.ComboBoxTax.TabIndex = 33
        '
        'Label62
        '
        Me.Label62.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label62.ForeColor = System.Drawing.Color.Black
        Me.Label62.Location = New System.Drawing.Point(381, 285)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(64, 16)
        Me.Label62.TabIndex = 174
        Me.Label62.Text = "Currency"
        '
        'ComboboxTSSISeg
        '
        Me.ComboboxTSSISeg.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboboxTSSISeg.Location = New System.Drawing.Point(675, 305)
        Me.ComboboxTSSISeg.Name = "ComboboxTSSISeg"
        Me.ComboboxTSSISeg.Size = New System.Drawing.Size(52, 22)
        Me.ComboboxTSSISeg.TabIndex = 38
        '
        'ComboboxISR
        '
        Me.ComboboxISR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboboxISR.Location = New System.Drawing.Point(617, 304)
        Me.ComboboxISR.Name = "ComboboxISR"
        Me.ComboboxISR.Size = New System.Drawing.Size(54, 22)
        Me.ComboboxISR.TabIndex = 176
        '
        'ComboBoxCurrency
        '
        Me.ComboBoxCurrency.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxCurrency.Items.AddRange(New Object() {"INR", "EUR", "USD", "SGD", "SEK", "JPY", "GBP", "DKK", "CAD", "CHF"})
        Me.ComboBoxCurrency.Location = New System.Drawing.Point(385, 304)
        Me.ComboBoxCurrency.Name = "ComboBoxCurrency"
        Me.ComboBoxCurrency.Size = New System.Drawing.Size(50, 22)
        Me.ComboBoxCurrency.TabIndex = 173
        '
        'ComboBoxClass3
        '
        Me.ComboBoxClass3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxClass3.Items.AddRange(New Object() {"K", "I", "Non KI", " "})
        Me.ComboBoxClass3.Location = New System.Drawing.Point(507, 305)
        Me.ComboBoxClass3.Name = "ComboBoxClass3"
        Me.ComboBoxClass3.Size = New System.Drawing.Size(50, 22)
        Me.ComboBoxClass3.TabIndex = 35
        '
        'ComboboxClass
        '
        Me.ComboboxClass.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboboxClass.Items.AddRange(New Object() {"A", "B", "C", "Z", "G", " "})
        Me.ComboboxClass.Location = New System.Drawing.Point(445, 305)
        Me.ComboboxClass.Name = "ComboboxClass"
        Me.ComboboxClass.Size = New System.Drawing.Size(56, 22)
        Me.ComboboxClass.TabIndex = 34
        '
        'ComboBoxCSR
        '
        Me.ComboBoxCSR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxCSR.Location = New System.Drawing.Point(563, 305)
        Me.ComboBoxCSR.Name = "ComboBoxCSR"
        Me.ComboBoxCSR.Size = New System.Drawing.Size(48, 22)
        Me.ComboBoxCSR.TabIndex = 36
        '
        'ComboboxSegment
        '
        Me.ComboboxSegment.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboboxSegment.Location = New System.Drawing.Point(733, 305)
        Me.ComboboxSegment.Name = "ComboboxSegment"
        Me.ComboboxSegment.Size = New System.Drawing.Size(60, 22)
        Me.ComboboxSegment.TabIndex = 39
        '
        'Label65
        '
        Me.Label65.AutoSize = True
        Me.Label65.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Label65.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label65.Location = New System.Drawing.Point(673, 292)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(106, 11)
        Me.Label65.TabIndex = 175
        Me.Label65.Text = "TSSISeg          TSSSeg"
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Label63.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label63.Location = New System.Drawing.Point(461, 292)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(203, 11)
        Me.Label63.TabIndex = 172
        Me.Label63.Text = " CL1          CL3             CSR                ISR  "
        '
        'CheckedListBoxCertificate
        '
        Me.CheckedListBoxCertificate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.CheckedListBoxCertificate.Location = New System.Drawing.Point(675, 69)
        Me.CheckedListBoxCertificate.Name = "CheckedListBoxCertificate"
        Me.CheckedListBoxCertificate.Size = New System.Drawing.Size(242, 4)
        Me.CheckedListBoxCertificate.TabIndex = 158
        '
        'Label44
        '
        Me.Label44.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label44.ForeColor = System.Drawing.Color.Black
        Me.Label44.Location = New System.Drawing.Point(672, 56)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(121, 13)
        Me.Label44.TabIndex = 157
        Me.Label44.Text = "Certificate Requirement"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(912, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 24)
        Me.Label4.TabIndex = 89
        Me.Label4.Text = "Enq. Recd Dt."
        '
        'DTPRegDt
        '
        Me.DTPRegDt.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DTPRegDt.Enabled = False
        Me.DTPRegDt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTPRegDt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTPRegDt.Location = New System.Drawing.Point(185, 24)
        Me.DTPRegDt.Name = "DTPRegDt"
        Me.DTPRegDt.Size = New System.Drawing.Size(104, 22)
        Me.DTPRegDt.TabIndex = 2
        '
        'dtpEnqDt
        '
        Me.dtpEnqDt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpEnqDt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEnqDt.Location = New System.Drawing.Point(576, 24)
        Me.dtpEnqDt.Name = "dtpEnqDt"
        Me.dtpEnqDt.Size = New System.Drawing.Size(88, 20)
        Me.dtpEnqDt.TabIndex = 4
        '
        'DTPEnqRecd
        '
        Me.DTPEnqRecd.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DTPEnqRecd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTPEnqRecd.Location = New System.Drawing.Point(992, 24)
        Me.DTPEnqRecd.Name = "DTPEnqRecd"
        Me.DTPEnqRecd.Size = New System.Drawing.Size(88, 20)
        Me.DTPEnqRecd.TabIndex = 6
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.Color.Transparent
        Me.Label38.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.ForeColor = System.Drawing.Color.Black
        Me.Label38.Location = New System.Drawing.Point(712, 16)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(184, 16)
        Me.Label38.TabIndex = 79
        Me.Label38.Text = "Cust Part Description"
        '
        'Label37
        '
        Me.Label37.BackColor = System.Drawing.Color.Transparent
        Me.Label37.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.ForeColor = System.Drawing.Color.Black
        Me.Label37.Location = New System.Drawing.Point(368, 16)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(128, 16)
        Me.Label37.TabIndex = 78
        Me.Label37.Text = "Description"
        '
        'Label36
        '
        Me.Label36.BackColor = System.Drawing.Color.Transparent
        Me.Label36.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.ForeColor = System.Drawing.Color.Black
        Me.Label36.Location = New System.Drawing.Point(200, 16)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(64, 16)
        Me.Label36.TabIndex = 77
        Me.Label36.Text = "Part No."
        '
        'txtpart
        '
        Me.txtpart.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpart.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtpart.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpart.Location = New System.Drawing.Point(200, 32)
        Me.txtpart.MaxLength = 50
        Me.txtpart.Name = "txtpart"
        Me.txtpart.Size = New System.Drawing.Size(144, 20)
        Me.txtpart.TabIndex = 60
        '
        'txtCustPart
        '
        Me.txtCustPart.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustPart.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustPart.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustPart.Location = New System.Drawing.Point(520, 32)
        Me.txtCustPart.MaxLength = 50
        Me.txtCustPart.Name = "txtCustPart"
        Me.txtCustPart.Size = New System.Drawing.Size(184, 20)
        Me.txtCustPart.TabIndex = 62
        '
        'txtPartDesc
        '
        Me.txtPartDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPartDesc.Location = New System.Drawing.Point(352, 32)
        Me.txtPartDesc.MaxLength = 50
        Me.txtPartDesc.Name = "txtPartDesc"
        Me.txtPartDesc.Size = New System.Drawing.Size(160, 20)
        Me.txtPartDesc.TabIndex = 61
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
        Me.txtslno.TabIndex = 57
        '
        'btnsave
        '
        Me.btnsave.BackColor = System.Drawing.Color.LightGray
        Me.btnsave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnsave.ForeColor = System.Drawing.Color.Blue
        Me.btnsave.Location = New System.Drawing.Point(977, 69)
        Me.btnsave.Name = "btnsave"
        Me.btnsave.Size = New System.Drawing.Size(48, 24)
        Me.btnsave.TabIndex = 70
        Me.btnsave.Text = "Save"
        Me.btnsave.UseVisualStyleBackColor = False
        '
        'txtMaterial
        '
        Me.txtMaterial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMaterial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterial.Location = New System.Drawing.Point(245, 72)
        Me.txtMaterial.MaxLength = 80
        Me.txtMaterial.Name = "txtMaterial"
        Me.txtMaterial.Size = New System.Drawing.Size(180, 20)
        Me.txtMaterial.TabIndex = 68
        '
        'txtDetailSpecial
        '
        Me.txtDetailSpecial.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDetailSpecial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDetailSpecial.Location = New System.Drawing.Point(432, 72)
        Me.txtDetailSpecial.MaxLength = 100
        Me.txtDetailSpecial.Name = "txtDetailSpecial"
        Me.txtDetailSpecial.Size = New System.Drawing.Size(240, 20)
        Me.txtDetailSpecial.TabIndex = 69
        '
        'DataUpdation
        '
        Me.DataUpdation.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataUpdation.Controls.Add(Me.btnItemHistory)
        Me.DataUpdation.Controls.Add(Me.GroupBoxPartCreation)
        Me.DataUpdation.Controls.Add(Me.ComboboxReq)
        Me.DataUpdation.Controls.Add(Me.Label57)
        Me.DataUpdation.Controls.Add(Me.txtitemkey)
        Me.DataUpdation.Controls.Add(Me.btnItemDelete)
        Me.DataUpdation.Controls.Add(Me.txtdetailintcode)
        Me.DataUpdation.Controls.Add(Me.btnAdd)
        Me.DataUpdation.Controls.Add(Me.ComboBoxuom)
        Me.DataUpdation.Controls.Add(Me.GroupBox4)
        Me.DataUpdation.Controls.Add(Me.Label35)
        Me.DataUpdation.Controls.Add(Me.Label54)
        Me.DataUpdation.Controls.Add(Me.CheckedListBoxCertificate)
        Me.DataUpdation.Controls.Add(Me.ComboBoxItemSource)
        Me.DataUpdation.Controls.Add(Me.Label44)
        Me.DataUpdation.Controls.Add(Me.Label46)
        Me.DataUpdation.Controls.Add(Me.txtRecVend)
        Me.DataUpdation.Controls.Add(Me.Label45)
        Me.DataUpdation.Controls.Add(Me.Label43)
        Me.DataUpdation.Controls.Add(Me.Label42)
        Me.DataUpdation.Controls.Add(Me.txtDimension)
        Me.DataUpdation.Controls.Add(Me.txtCustDesc)
        Me.DataUpdation.Controls.Add(Me.ComboBoxFSYesNo)
        Me.DataUpdation.Controls.Add(Me.Label41)
        Me.DataUpdation.Controls.Add(Me.Label40)
        Me.DataUpdation.Controls.Add(Me.Label39)
        Me.DataUpdation.Controls.Add(Me.Label38)
        Me.DataUpdation.Controls.Add(Me.Label37)
        Me.DataUpdation.Controls.Add(Me.Label36)
        Me.DataUpdation.Controls.Add(Me.txtpart)
        Me.DataUpdation.Controls.Add(Me.txtCustPart)
        Me.DataUpdation.Controls.Add(Me.txtPartDesc)
        Me.DataUpdation.Controls.Add(Me.txtslno)
        Me.DataUpdation.Controls.Add(Me.btnsave)
        Me.DataUpdation.Controls.Add(Me.txtMaterial)
        Me.DataUpdation.Controls.Add(Me.txtDetailSpecial)
        Me.DataUpdation.Controls.Add(Me.CheckBoxPartCreation)
        Me.DataUpdation.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataUpdation.ForeColor = System.Drawing.Color.Firebrick
        Me.DataUpdation.Location = New System.Drawing.Point(16, 575)
        Me.DataUpdation.Name = "DataUpdation"
        Me.DataUpdation.Size = New System.Drawing.Size(1248, 122)
        Me.DataUpdation.TabIndex = 61
        Me.DataUpdation.TabStop = False
        Me.DataUpdation.Text = "Details"
        '
        'btnItemHistory
        '
        Me.btnItemHistory.BackColor = System.Drawing.Color.LightGray
        Me.btnItemHistory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnItemHistory.ForeColor = System.Drawing.Color.Blue
        Me.btnItemHistory.Location = New System.Drawing.Point(836, 91)
        Me.btnItemHistory.Name = "btnItemHistory"
        Me.btnItemHistory.Size = New System.Drawing.Size(84, 24)
        Me.btnItemHistory.TabIndex = 161
        Me.btnItemHistory.Text = "Item History"
        Me.btnItemHistory.UseVisualStyleBackColor = False
        '
        'GroupBoxPartCreation
        '
        Me.GroupBoxPartCreation.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBoxPartCreation.Controls.Add(Me.ComboBoxInvAc)
        Me.GroupBoxPartCreation.Controls.Add(Me.LblClose)
        Me.GroupBoxPartCreation.Controls.Add(Me.ComboBoxprodline)
        Me.GroupBoxPartCreation.Controls.Add(Me.lblLeadTimeType)
        Me.GroupBoxPartCreation.Controls.Add(Me.txtsp2)
        Me.GroupBoxPartCreation.Controls.Add(Me.txtsp1)
        Me.GroupBoxPartCreation.Controls.Add(Me.Label64)
        Me.GroupBoxPartCreation.Controls.Add(Me.txtchilditemDesc)
        Me.GroupBoxPartCreation.Controls.Add(Me.txtfix)
        Me.GroupBoxPartCreation.Controls.Add(Me.txtinsp)
        Me.GroupBoxPartCreation.Controls.Add(Me.txtrun)
        Me.GroupBoxPartCreation.Controls.Add(Me.ComboBoxBuyer)
        Me.GroupBoxPartCreation.Controls.Add(Me.ComboBoxPlanner)
        Me.GroupBoxPartCreation.Controls.Add(Me.ComboBoxItemType)
        Me.GroupBoxPartCreation.Controls.Add(Me.txtpartDescription)
        Me.GroupBoxPartCreation.Controls.Add(Me.sp1)
        Me.GroupBoxPartCreation.Controls.Add(Me.Label59)
        Me.GroupBoxPartCreation.Controls.Add(Me.Label5)
        Me.GroupBoxPartCreation.Controls.Add(Me.Label60)
        Me.GroupBoxPartCreation.Controls.Add(Me.lblPlanner)
        Me.GroupBoxPartCreation.Controls.Add(Me.lblChildItem)
        Me.GroupBoxPartCreation.Controls.Add(Me.lblmb)
        Me.GroupBoxPartCreation.Controls.Add(Me.lblpartdesc)
        Me.GroupBoxPartCreation.Controls.Add(Me.lblPart)
        Me.GroupBoxPartCreation.Controls.Add(Me.txtPartnum)
        Me.GroupBoxPartCreation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.GroupBoxPartCreation.ForeColor = System.Drawing.Color.Black
        Me.GroupBoxPartCreation.Location = New System.Drawing.Point(161, 1)
        Me.GroupBoxPartCreation.Name = "GroupBoxPartCreation"
        Me.GroupBoxPartCreation.Size = New System.Drawing.Size(420, 23)
        Me.GroupBoxPartCreation.TabIndex = 159
        Me.GroupBoxPartCreation.TabStop = False
        Me.GroupBoxPartCreation.Text = "Part Creation"
        Me.GroupBoxPartCreation.Visible = False
        '
        'ComboBoxInvAc
        '
        Me.ComboBoxInvAc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ComboBoxInvAc.Enabled = False
        Me.ComboBoxInvAc.Location = New System.Drawing.Point(501, 84)
        Me.ComboBoxInvAc.Name = "ComboBoxInvAc"
        Me.ComboBoxInvAc.Size = New System.Drawing.Size(152, 20)
        Me.ComboBoxInvAc.TabIndex = 78
        '
        'LblClose
        '
        Me.LblClose.AutoSize = True
        Me.LblClose.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.LblClose.ForeColor = System.Drawing.Color.Red
        Me.LblClose.Location = New System.Drawing.Point(996, 5)
        Me.LblClose.Name = "LblClose"
        Me.LblClose.Size = New System.Drawing.Size(20, 19)
        Me.LblClose.TabIndex = 77
        Me.LblClose.Text = "X"
        '
        'ComboBoxprodline
        '
        Me.ComboBoxprodline.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxprodline.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxprodline.Items.AddRange(New Object() {"Warehouse", "Factory"})
        Me.ComboBoxprodline.Location = New System.Drawing.Point(504, 53)
        Me.ComboBoxprodline.Name = "ComboBoxprodline"
        Me.ComboBoxprodline.Size = New System.Drawing.Size(92, 22)
        Me.ComboBoxprodline.TabIndex = 10
        '
        'lblLeadTimeType
        '
        Me.lblLeadTimeType.AutoSize = True
        Me.lblLeadTimeType.Location = New System.Drawing.Point(495, 5)
        Me.lblLeadTimeType.Name = "lblLeadTimeType"
        Me.lblLeadTimeType.Size = New System.Drawing.Size(99, 14)
        Me.lblLeadTimeType.TabIndex = 74
        Me.lblLeadTimeType.Text = "Run      Fix       Insp"
        '
        'txtsp2
        '
        Me.txtsp2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtsp2.Location = New System.Drawing.Point(660, 67)
        Me.txtsp2.MaxLength = 200
        Me.txtsp2.Multiline = True
        Me.txtsp2.Name = "txtsp2"
        Me.txtsp2.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtsp2.Size = New System.Drawing.Size(329, 46)
        Me.txtsp2.TabIndex = 13
        '
        'txtsp1
        '
        Me.txtsp1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtsp1.Location = New System.Drawing.Point(660, 9)
        Me.txtsp1.MaxLength = 200
        Me.txtsp1.Multiline = True
        Me.txtsp1.Name = "txtsp1"
        Me.txtsp1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtsp1.Size = New System.Drawing.Size(329, 56)
        Me.txtsp1.TabIndex = 12
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Location = New System.Drawing.Point(600, 56)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(54, 14)
        Me.Label64.TabIndex = 67
        Me.Label64.Text = "Sp Note 2"
        '
        'txtchilditemDesc
        '
        Me.txtchilditemDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtchilditemDesc.Location = New System.Drawing.Point(261, 27)
        Me.txtchilditemDesc.MaxLength = 200
        Me.txtchilditemDesc.Multiline = True
        Me.txtchilditemDesc.Name = "txtchilditemDesc"
        Me.txtchilditemDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtchilditemDesc.Size = New System.Drawing.Size(179, 82)
        Me.txtchilditemDesc.TabIndex = 6
        '
        'txtfix
        '
        Me.txtfix.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtfix.Location = New System.Drawing.Point(533, 23)
        Me.txtfix.Name = "txtfix"
        Me.txtfix.Size = New System.Drawing.Size(32, 20)
        Me.txtfix.TabIndex = 8
        '
        'txtinsp
        '
        Me.txtinsp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtinsp.Location = New System.Drawing.Point(564, 23)
        Me.txtinsp.Name = "txtinsp"
        Me.txtinsp.Size = New System.Drawing.Size(32, 20)
        Me.txtinsp.TabIndex = 9
        '
        'txtrun
        '
        Me.txtrun.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtrun.Location = New System.Drawing.Point(501, 23)
        Me.txtrun.Name = "txtrun"
        Me.txtrun.Size = New System.Drawing.Size(44, 20)
        Me.txtrun.TabIndex = 7
        '
        'ComboBoxBuyer
        '
        Me.ComboBoxBuyer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxBuyer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxBuyer.Items.AddRange(New Object() {"WHB", "DHB"})
        Me.ComboBoxBuyer.Location = New System.Drawing.Point(149, 93)
        Me.ComboBoxBuyer.Name = "ComboBoxBuyer"
        Me.ComboBoxBuyer.Size = New System.Drawing.Size(56, 22)
        Me.ComboBoxBuyer.TabIndex = 5
        '
        'ComboBoxPlanner
        '
        Me.ComboBoxPlanner.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPlanner.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxPlanner.Items.AddRange(New Object() {"WHP", "DHP"})
        Me.ComboBoxPlanner.Location = New System.Drawing.Point(88, 93)
        Me.ComboBoxPlanner.Name = "ComboBoxPlanner"
        Me.ComboBoxPlanner.Size = New System.Drawing.Size(56, 22)
        Me.ComboBoxPlanner.TabIndex = 4
        '
        'ComboBoxItemType
        '
        Me.ComboBoxItemType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxItemType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxItemType.Items.AddRange(New Object() {"B", "M", "S", ""})
        Me.ComboBoxItemType.Location = New System.Drawing.Point(88, 67)
        Me.ComboBoxItemType.Name = "ComboBoxItemType"
        Me.ComboBoxItemType.Size = New System.Drawing.Size(56, 22)
        Me.ComboBoxItemType.TabIndex = 2
        '
        'txtpartDescription
        '
        Me.txtpartDescription.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpartDescription.Enabled = False
        Me.txtpartDescription.Location = New System.Drawing.Point(77, 45)
        Me.txtpartDescription.Name = "txtpartDescription"
        Me.txtpartDescription.Size = New System.Drawing.Size(176, 20)
        Me.txtpartDescription.TabIndex = 1
        '
        'sp1
        '
        Me.sp1.AutoSize = True
        Me.sp1.Location = New System.Drawing.Point(600, 9)
        Me.sp1.Name = "sp1"
        Me.sp1.Size = New System.Drawing.Size(54, 14)
        Me.sp1.TabIndex = 11
        Me.sp1.Text = "Sp Note 1"
        '
        'Label59
        '
        Me.Label59.AutoSize = True
        Me.Label59.Location = New System.Drawing.Point(445, 84)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(40, 14)
        Me.Label59.TabIndex = 10
        Me.Label59.Text = "Inv A/c"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(445, 56)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 14)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Prod LIne"
        '
        'Label60
        '
        Me.Label60.AutoSize = True
        Me.Label60.Location = New System.Drawing.Point(445, 25)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(56, 14)
        Me.Label60.TabIndex = 8
        Me.Label60.Text = "Lead Time"
        '
        'lblPlanner
        '
        Me.lblPlanner.AutoSize = True
        Me.lblPlanner.Location = New System.Drawing.Point(5, 96)
        Me.lblPlanner.Name = "lblPlanner"
        Me.lblPlanner.Size = New System.Drawing.Size(75, 14)
        Me.lblPlanner.TabIndex = 6
        Me.lblPlanner.Text = "Planner/Buyer"
        '
        'lblChildItem
        '
        Me.lblChildItem.AutoSize = True
        Me.lblChildItem.Location = New System.Drawing.Point(259, 9)
        Me.lblChildItem.Name = "lblChildItem"
        Me.lblChildItem.Size = New System.Drawing.Size(80, 14)
        Me.lblChildItem.TabIndex = 5
        Me.lblChildItem.Text = "Child Item Desc"
        '
        'lblmb
        '
        Me.lblmb.AutoSize = True
        Me.lblmb.Location = New System.Drawing.Point(8, 69)
        Me.lblmb.Name = "lblmb"
        Me.lblmb.Size = New System.Drawing.Size(52, 14)
        Me.lblmb.TabIndex = 4
        Me.lblmb.Text = "Item Type"
        '
        'lblpartdesc
        '
        Me.lblpartdesc.AutoSize = True
        Me.lblpartdesc.Location = New System.Drawing.Point(5, 40)
        Me.lblpartdesc.Name = "lblpartdesc"
        Me.lblpartdesc.Size = New System.Drawing.Size(54, 14)
        Me.lblpartdesc.TabIndex = 2
        Me.lblpartdesc.Text = "Part Desc"
        '
        'lblPart
        '
        Me.lblPart.AutoSize = True
        Me.lblPart.Location = New System.Drawing.Point(5, 21)
        Me.lblPart.Name = "lblPart"
        Me.lblPart.Size = New System.Drawing.Size(66, 14)
        Me.lblPart.TabIndex = 1
        Me.lblPart.Text = "Part Number"
        '
        'txtPartnum
        '
        Me.txtPartnum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPartnum.Enabled = False
        Me.txtPartnum.Location = New System.Drawing.Point(77, 19)
        Me.txtPartnum.Name = "txtPartnum"
        Me.txtPartnum.Size = New System.Drawing.Size(176, 20)
        Me.txtPartnum.TabIndex = 0
        '
        'ComboboxReq
        '
        Me.ComboboxReq.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboboxReq.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboboxReq.Items.AddRange(New Object() {"Price", "Part Creation", "Both"})
        Me.ComboboxReq.Location = New System.Drawing.Point(865, 29)
        Me.ComboboxReq.Name = "ComboboxReq"
        Me.ComboboxReq.Size = New System.Drawing.Size(88, 22)
        Me.ComboboxReq.TabIndex = 64
        '
        'Label57
        '
        Me.Label57.BackColor = System.Drawing.Color.Transparent
        Me.Label57.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label57.ForeColor = System.Drawing.Color.Black
        Me.Label57.Location = New System.Drawing.Point(863, 16)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(33, 16)
        Me.Label57.TabIndex = 132
        Me.Label57.Text = "Req"
        '
        'txtitemkey
        '
        Me.txtitemkey.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtitemkey.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtitemkey.Location = New System.Drawing.Point(328, 8)
        Me.txtitemkey.MaxLength = 50
        Me.txtitemkey.Name = "txtitemkey"
        Me.txtitemkey.Size = New System.Drawing.Size(32, 20)
        Me.txtitemkey.TabIndex = 128
        Me.txtitemkey.Visible = False
        '
        'btnItemDelete
        '
        Me.btnItemDelete.BackColor = System.Drawing.Color.LightGray
        Me.btnItemDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnItemDelete.ForeColor = System.Drawing.Color.Blue
        Me.btnItemDelete.Location = New System.Drawing.Point(924, 93)
        Me.btnItemDelete.Name = "btnItemDelete"
        Me.btnItemDelete.Size = New System.Drawing.Size(101, 24)
        Me.btnItemDelete.TabIndex = 127
        Me.btnItemDelete.Text = "Item Delete"
        Me.btnItemDelete.UseVisualStyleBackColor = False
        '
        'txtdetailintcode
        '
        Me.txtdetailintcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtdetailintcode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdetailintcode.Location = New System.Drawing.Point(264, 1)
        Me.txtdetailintcode.MaxLength = 50
        Me.txtdetailintcode.Name = "txtdetailintcode"
        Me.txtdetailintcode.Size = New System.Drawing.Size(32, 20)
        Me.txtdetailintcode.TabIndex = 126
        Me.txtdetailintcode.Visible = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.LightGray
        Me.btnAdd.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.ForeColor = System.Drawing.Color.Blue
        Me.btnAdd.Location = New System.Drawing.Point(924, 69)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(48, 24)
        Me.btnAdd.TabIndex = 125
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'ComboBoxuom
        '
        Me.ComboBoxuom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxuom.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxuom.Items.AddRange(New Object() {"EA", "Pcs", "Sets", "Mtrs", "Cms", "Ltrs", "Ft", "Sheet", "Length", "Sq Ft"})
        Me.ComboBoxuom.Location = New System.Drawing.Point(960, 32)
        Me.ComboBoxuom.Name = "ComboBoxuom"
        Me.ComboBoxuom.Size = New System.Drawing.Size(64, 22)
        Me.ComboBoxuom.TabIndex = 65
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnQtyAdd)
        Me.GroupBox4.Controls.Add(Me.QtyEdit)
        Me.GroupBox4.Controls.Add(Me.lblqty)
        Me.GroupBox4.Controls.Add(Me.txtqtyintcode)
        Me.GroupBox4.Controls.Add(Me.btnItemsave)
        Me.GroupBox4.Controls.Add(Me.DataGridQty)
        Me.GroupBox4.Controls.Add(Me.txtqty)
        Me.GroupBox4.Controls.Add(Me.ComboBoxReqType)
        Me.GroupBox4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(1032, 0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(216, 120)
        Me.GroupBox4.TabIndex = 123
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Qty Slabs.  S.No."
        '
        'btnQtyAdd
        '
        Me.btnQtyAdd.BackColor = System.Drawing.Color.LightGray
        Me.btnQtyAdd.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnQtyAdd.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnQtyAdd.Location = New System.Drawing.Point(160, 16)
        Me.btnQtyAdd.Name = "btnQtyAdd"
        Me.btnQtyAdd.Size = New System.Drawing.Size(48, 24)
        Me.btnQtyAdd.TabIndex = 70
        Me.btnQtyAdd.Text = "Add"
        Me.btnQtyAdd.UseVisualStyleBackColor = False
        Me.btnQtyAdd.Visible = False
        '
        'QtyEdit
        '
        Me.QtyEdit.BackColor = System.Drawing.Color.LightGray
        Me.QtyEdit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QtyEdit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.QtyEdit.Location = New System.Drawing.Point(160, 76)
        Me.QtyEdit.Name = "QtyEdit"
        Me.QtyEdit.Size = New System.Drawing.Size(56, 24)
        Me.QtyEdit.TabIndex = 132
        Me.QtyEdit.Text = "Qty Edit"
        Me.QtyEdit.UseVisualStyleBackColor = False
        '
        'lblqty
        '
        Me.lblqty.Location = New System.Drawing.Point(104, 0)
        Me.lblqty.Name = "lblqty"
        Me.lblqty.Size = New System.Drawing.Size(16, 16)
        Me.lblqty.TabIndex = 131
        '
        'txtqtyintcode
        '
        Me.txtqtyintcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtqtyintcode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtqtyintcode.Location = New System.Drawing.Point(120, 0)
        Me.txtqtyintcode.Name = "txtqtyintcode"
        Me.txtqtyintcode.Size = New System.Drawing.Size(16, 20)
        Me.txtqtyintcode.TabIndex = 130
        Me.txtqtyintcode.Visible = False
        '
        'btnItemsave
        '
        Me.btnItemsave.BackColor = System.Drawing.Color.LightGray
        Me.btnItemsave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnItemsave.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnItemsave.Location = New System.Drawing.Point(160, 48)
        Me.btnItemsave.Name = "btnItemsave"
        Me.btnItemsave.Size = New System.Drawing.Size(48, 24)
        Me.btnItemsave.TabIndex = 73
        Me.btnItemsave.Text = "Save"
        Me.btnItemsave.UseVisualStyleBackColor = False
        '
        'DataGridQty
        '
        Me.DataGridQty.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.DataGridQty.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQty.CaptionVisible = False
        Me.DataGridQty.DataMember = ""
        Me.DataGridQty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQty.HeaderFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridQty.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridQty.Location = New System.Drawing.Point(8, 16)
        Me.DataGridQty.Name = "DataGridQty"
        Me.DataGridQty.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridQty.ParentRowsVisible = False
        Me.DataGridQty.PreferredColumnWidth = 85
        Me.DataGridQty.ReadOnly = True
        Me.DataGridQty.RowHeadersVisible = False
        Me.DataGridQty.Size = New System.Drawing.Size(144, 72)
        Me.DataGridQty.TabIndex = 128
        '
        'txtqty
        '
        Me.txtqty.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtqty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtqty.Location = New System.Drawing.Point(8, 96)
        Me.txtqty.Name = "txtqty"
        Me.txtqty.Size = New System.Drawing.Size(56, 20)
        Me.txtqty.TabIndex = 71
        '
        'ComboBoxReqType
        '
        Me.ComboBoxReqType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxReqType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxReqType.Items.AddRange(New Object() {"Proto", "Monthly", "Anual", "Lot1 (one time)", "Lot2 (one time)", "Lot3 (one time)", "Lot1 (Monthly)", "Lot2 (Monthly)", "Lot3 (Monthly)", "Lot1 (Yearly)", "Lot2 (Yearly)", "Lot3 (Yearly)"})
        Me.ComboBoxReqType.Location = New System.Drawing.Point(69, 93)
        Me.ComboBoxReqType.Name = "ComboBoxReqType"
        Me.ComboBoxReqType.Size = New System.Drawing.Size(80, 22)
        Me.ComboBoxReqType.TabIndex = 72
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.Color.Transparent
        Me.Label35.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.ForeColor = System.Drawing.Color.Black
        Me.Label35.Location = New System.Drawing.Point(56, 16)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(56, 16)
        Me.Label35.TabIndex = 122
        Me.Label35.Text = "Avbl FS"
        '
        'Label54
        '
        Me.Label54.BackColor = System.Drawing.Color.Transparent
        Me.Label54.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label54.ForeColor = System.Drawing.Color.Black
        Me.Label54.Location = New System.Drawing.Point(120, 16)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(56, 16)
        Me.Label54.TabIndex = 121
        Me.Label54.Text = "Source"
        '
        'ComboBoxItemSource
        '
        Me.ComboBoxItemSource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxItemSource.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxItemSource.Items.AddRange(New Object() {"Procure", "Mfg in Blore"})
        Me.ComboBoxItemSource.Location = New System.Drawing.Point(112, 32)
        Me.ComboBoxItemSource.Name = "ComboBoxItemSource"
        Me.ComboBoxItemSource.Size = New System.Drawing.Size(80, 22)
        Me.ComboBoxItemSource.TabIndex = 59
        '
        'Label46
        '
        Me.Label46.BackColor = System.Drawing.Color.Transparent
        Me.Label46.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label46.ForeColor = System.Drawing.Color.Black
        Me.Label46.Location = New System.Drawing.Point(8, 56)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(112, 16)
        Me.Label46.TabIndex = 119
        Me.Label46.Text = "Recom. Vendor"
        '
        'txtRecVend
        '
        Me.txtRecVend.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRecVend.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRecVend.Location = New System.Drawing.Point(8, 72)
        Me.txtRecVend.MaxLength = 50
        Me.txtRecVend.Name = "txtRecVend"
        Me.txtRecVend.Size = New System.Drawing.Size(112, 20)
        Me.txtRecVend.TabIndex = 66
        '
        'Label45
        '
        Me.Label45.BackColor = System.Drawing.Color.Transparent
        Me.Label45.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label45.ForeColor = System.Drawing.Color.Black
        Me.Label45.Location = New System.Drawing.Point(981, 13)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(32, 16)
        Me.Label45.TabIndex = 116
        Me.Label45.Text = "Uom"
        '
        'Label43
        '
        Me.Label43.BackColor = System.Drawing.Color.Transparent
        Me.Label43.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label43.ForeColor = System.Drawing.Color.Black
        Me.Label43.Location = New System.Drawing.Point(432, 53)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(120, 16)
        Me.Label43.TabIndex = 114
        Me.Label43.Text = "Special Instructions"
        '
        'Label42
        '
        Me.Label42.BackColor = System.Drawing.Color.Transparent
        Me.Label42.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label42.ForeColor = System.Drawing.Color.Black
        Me.Label42.Location = New System.Drawing.Point(248, 55)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(64, 16)
        Me.Label42.TabIndex = 112
        Me.Label42.Text = "Material"
        '
        'txtDimension
        '
        Me.txtDimension.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDimension.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDimension.Location = New System.Drawing.Point(123, 72)
        Me.txtDimension.Name = "txtDimension"
        Me.txtDimension.Size = New System.Drawing.Size(117, 20)
        Me.txtDimension.TabIndex = 67
        '
        'txtCustDesc
        '
        Me.txtCustDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustDesc.Location = New System.Drawing.Point(712, 32)
        Me.txtCustDesc.MaxLength = 80
        Me.txtCustDesc.Name = "txtCustDesc"
        Me.txtCustDesc.Size = New System.Drawing.Size(152, 20)
        Me.txtCustDesc.TabIndex = 63
        '
        'ComboBoxFSYesNo
        '
        Me.ComboBoxFSYesNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxFSYesNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxFSYesNo.Items.AddRange(New Object() {"Yes", "No"})
        Me.ComboBoxFSYesNo.Location = New System.Drawing.Point(48, 32)
        Me.ComboBoxFSYesNo.Name = "ComboBoxFSYesNo"
        Me.ComboBoxFSYesNo.Size = New System.Drawing.Size(56, 22)
        Me.ComboBoxFSYesNo.TabIndex = 58
        '
        'Label41
        '
        Me.Label41.BackColor = System.Drawing.Color.Transparent
        Me.Label41.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label41.ForeColor = System.Drawing.Color.Black
        Me.Label41.Location = New System.Drawing.Point(8, 16)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(40, 16)
        Me.Label41.TabIndex = 82
        Me.Label41.Text = "Sl.No."
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(4, Byte), Integer))
        Me.Label40.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.ForeColor = System.Drawing.Color.Black
        Me.Label40.Location = New System.Drawing.Point(120, 57)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(64, 16)
        Me.Label40.TabIndex = 81
        Me.Label40.Text = "Dimension"
        '
        'Label39
        '
        Me.Label39.BackColor = System.Drawing.Color.Transparent
        Me.Label39.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.ForeColor = System.Drawing.Color.Black
        Me.Label39.Location = New System.Drawing.Point(517, 17)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(176, 16)
        Me.Label39.TabIndex = 80
        Me.Label39.Text = "Customer Part No."
        '
        'CheckBoxPartCreation
        '
        Me.CheckBoxPartCreation.AutoSize = True
        Me.CheckBoxPartCreation.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxPartCreation.Location = New System.Drawing.Point(901, 9)
        Me.CheckBoxPartCreation.Name = "CheckBoxPartCreation"
        Me.CheckBoxPartCreation.Size = New System.Drawing.Size(74, 15)
        Me.CheckBoxPartCreation.TabIndex = 160
        Me.CheckBoxPartCreation.Text = "Part Create"
        Me.CheckBoxPartCreation.UseVisualStyleBackColor = True
        Me.CheckBoxPartCreation.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.lblMode1)
        Me.GroupBox1.Controls.Add(Me.txtenqintcode)
        Me.GroupBox1.Controls.Add(Me.RBTenderNo)
        Me.GroupBox1.Controls.Add(Me.RBTenderYes)
        Me.GroupBox1.Controls.Add(Me.Label53)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label51)
        Me.GroupBox1.Controls.Add(Me.ComboBoxSource)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtRegNo)
        Me.GroupBox1.Controls.Add(Me.DTPRegDt)
        Me.GroupBox1.Controls.Add(Me.txtEnqRef)
        Me.GroupBox1.Controls.Add(Me.dtpEnqDt)
        Me.GroupBox1.Controls.Add(Me.DTPEnqRecd)
        Me.GroupBox1.Controls.Add(Me.dtpTenderDueDt)
        Me.GroupBox1.Controls.Add(Me.lblmode)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(16, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1248, 56)
        Me.GroupBox1.TabIndex = 95
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Enquiry Details"
        '
        'lblMode1
        '
        Me.lblMode1.ForeColor = System.Drawing.Color.Red
        Me.lblMode1.Location = New System.Drawing.Point(664, 0)
        Me.lblMode1.Name = "lblMode1"
        Me.lblMode1.Size = New System.Drawing.Size(80, 15)
        Me.lblMode1.TabIndex = 115
        Me.lblMode1.Text = "1"
        '
        'txtenqintcode
        '
        Me.txtenqintcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtenqintcode.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtenqintcode.Location = New System.Drawing.Point(296, 40)
        Me.txtenqintcode.Name = "txtenqintcode"
        Me.txtenqintcode.Size = New System.Drawing.Size(80, 22)
        Me.txtenqintcode.TabIndex = 113
        Me.txtenqintcode.Visible = False
        '
        'RBTenderNo
        '
        Me.RBTenderNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBTenderNo.ForeColor = System.Drawing.Color.Red
        Me.RBTenderNo.Location = New System.Drawing.Point(1188, 13)
        Me.RBTenderNo.Name = "RBTenderNo"
        Me.RBTenderNo.Size = New System.Drawing.Size(44, 16)
        Me.RBTenderNo.TabIndex = 8
        Me.RBTenderNo.Text = "No"
        '
        'RBTenderYes
        '
        Me.RBTenderYes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBTenderYes.ForeColor = System.Drawing.Color.Red
        Me.RBTenderYes.Location = New System.Drawing.Point(1136, 13)
        Me.RBTenderYes.Name = "RBTenderYes"
        Me.RBTenderYes.Size = New System.Drawing.Size(48, 16)
        Me.RBTenderYes.TabIndex = 7
        Me.RBTenderYes.Text = "Yes"
        '
        'Label53
        '
        Me.Label53.BackColor = System.Drawing.Color.Transparent
        Me.Label53.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label53.ForeColor = System.Drawing.Color.Black
        Me.Label53.Location = New System.Drawing.Point(1080, 13)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(56, 16)
        Me.Label53.TabIndex = 112
        Me.Label53.Text = "Tender:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(1088, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 16)
        Me.Label1.TabIndex = 110
        Me.Label1.Text = "Due Dt."
        '
        'Label51
        '
        Me.Label51.BackColor = System.Drawing.Color.Transparent
        Me.Label51.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label51.ForeColor = System.Drawing.Color.Black
        Me.Label51.Location = New System.Drawing.Point(672, 24)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(96, 16)
        Me.Label51.TabIndex = 109
        Me.Label51.Text = "Enquiry Source"
        '
        'ComboBoxSource
        '
        Me.ComboBoxSource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxSource.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxSource.Location = New System.Drawing.Point(768, 24)
        Me.ComboBoxSource.Name = "ComboBoxSource"
        Me.ComboBoxSource.Size = New System.Drawing.Size(136, 22)
        Me.ComboBoxSource.TabIndex = 5
        '
        'dtpTenderDueDt
        '
        Me.dtpTenderDueDt.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.dtpTenderDueDt.CalendarTitleBackColor = System.Drawing.Color.Lime
        Me.dtpTenderDueDt.Checked = False
        Me.dtpTenderDueDt.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTenderDueDt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTenderDueDt.Location = New System.Drawing.Point(1136, 32)
        Me.dtpTenderDueDt.Name = "dtpTenderDueDt"
        Me.dtpTenderDueDt.ShowCheckBox = True
        Me.dtpTenderDueDt.Size = New System.Drawing.Size(104, 21)
        Me.dtpTenderDueDt.TabIndex = 9
        '
        'lblmode
        '
        Me.lblmode.ForeColor = System.Drawing.Color.Red
        Me.lblmode.Location = New System.Drawing.Point(592, 0)
        Me.lblmode.Name = "lblmode"
        Me.lblmode.Size = New System.Drawing.Size(80, 16)
        Me.lblmode.TabIndex = 114
        Me.lblmode.Text = "Trans Type:"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.CreatePrompt = True
        Me.SaveFileDialog1.InitialDirectory = "\\TSSBLRDOM111\PUBLIC\RFQ"
        Me.SaveFileDialog1.RestoreDirectory = True
        Me.SaveFileDialog1.Title = "PASTE THE FILE"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DereferenceLinks = False
        Me.OpenFileDialog1.InitialDirectory = "C:\ENQUIRY"
        Me.OpenFileDialog1.ReadOnlyChecked = True
        Me.OpenFileDialog1.ShowHelp = True
        Me.OpenFileDialog1.ShowReadOnly = True
        Me.OpenFileDialog1.Title = "COPY FILE FROM THIS LOCATION"
        '
        'lblEnqAdd
        '
        Me.lblEnqAdd.Font = New System.Drawing.Font("Monotype Corsiva", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEnqAdd.Location = New System.Drawing.Point(112, 0)
        Me.lblEnqAdd.Name = "lblEnqAdd"
        Me.lblEnqAdd.Size = New System.Drawing.Size(24, 24)
        Me.lblEnqAdd.TabIndex = 114
        Me.lblEnqAdd.Text = "+"
        '
        'Enquiry
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 24)
        Me.BackColor = System.Drawing.Color.CornflowerBlue
        Me.ClientSize = New System.Drawing.Size(1818, 1049)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.lblEnqAdd)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.DataUpdation)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Monotype Corsiva", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Firebrick
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Location = New System.Drawing.Point(-10, 0)
        Me.Name = "Enquiry"
        Me.Text = "Enquiry "
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.datagridDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxPriceAvble.ResumeLayout(False)
        CType(Me.datagridPriceAvble, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.PanelProjectDetails.ResumeLayout(False)
        Me.PanelProjectDetails.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        CType(Me.DataGridEnquiryEdit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridPartNumbers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupYesNo.ResumeLayout(False)
        Me.DataUpdation.ResumeLayout(False)
        Me.DataUpdation.PerformLayout()
        Me.GroupBoxPartCreation.ResumeLayout(False)
        Me.GroupBoxPartCreation.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.DataGridQty, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region
    'eneble XP styles
    '   Shared Sub Main()
    '      'Application.EnableVisualStyles()
    '     Application.Run(New main)
    'End Sub
    'set-up start parameters
    Private Sub main_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim use As New UserAccess
        purcheck = 0
        txtISR.Visible = False

        dtpTenderDueDt.Checked = False
        DTPStatusDt.Checked = False
        btnsave.Enabled = False
        lblMode1.Text = mode

        RadioButtonDomestic.Checked = True
        RadioButtonExport.Checked = False

        btnHeaderSave.Enabled = True
        btnItemDelete.Enabled = False

        detailsectiondisable()




        use.Close()
        use.Dispose()


        ConnectionString = ConnectionStringNew

        RBCustomerExisting.Checked = True
        RBTenderYes.Checked = False
        RBTenderNo.Checked = True
        rbdocyes.Checked = True

        listloadCertificate()
        COMBOLOAD()

        If lblMode1.Text = "Edit" Then
            EnquiryEdit()
            lblEnqAdd.Visible = False
        Else : lblMode1.Text = "Add"

            lblEnqAdd.Visible = True
        End If

        RadioButtonName.Checked = True


    End Sub
    'change status of text boxes
    Sub ChangeTextBoxStatus(ByVal newStatus As Boolean)
        'txtPartNo.Enabled = newStatus
        'txtFromStk.Enabled = newStatus
        'txtFromBin.Enabled = newStatus
        'txtToBin.Enabled = newStatus
        'txtToStk.Enabled = newStatus
        'txtQty.Enabled = newStatus
        'txtUnitPrice.Enabled = newStatus
        'txtLotNo.Enabled = newStatus
    End Sub
    Sub fillcustomerlist()


        'Dim custid As New DataGridTextBoxColumn
        'custid.HeaderText = "CustomerID"
        'colDescription.NullText = ""
        'colDescription.MappingName = ds.Tables(Str).Columns(1).ToString()
        'custid.Width = 150



        DataGridCustomer.Show()



        Dim sqlcon As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        txtCustID.Text = txtCustID.Text & "%"

        If RadioButtonName.Checked = True Then
            strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.FS_Customer " & _
                 "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRELLEBORG%') AND (CustomerID NOT LIKE '0000%')AND CustomerName like '" & txtCustID.Text & "' " & _
                    "ORDER BY CustomerID"

        Else

            strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.FS_Customer " & _
                     "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRELLEBORG%') AND (CustomerID NOT LIKE '0000%')AND CustomerID like '" & txtCustID.Text & "' " & _
                        "ORDER BY CustomerID"

        End If


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlcon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        sqlcon.Open()

        stockDAC.TableMappings.Add("Table", "Customer")
        'get data
        stockDAC.Fill(stockDC)

        DataGridCustomer.Width = 650 '1150
        DataGridCustomer.Height = 320 '800



        stockDC.Tables(0).Columns(0).ColumnName = "CustomerID"
        stockDC.Tables(0).Columns(1).ColumnName = "CustomerName"
        stockDC.Tables(0).Columns(2).ColumnName = "CustomerCity"
        stockDC.Tables(0).Columns(3).ColumnName = "CSR"



        'Dim custid As New DataGridTextBoxColumn
        'custid.HeaderText = "Customer_ID"
        'custid.NullText = ""
        'custid.MappingName = stockDC.Tables(0).Columns(0).ToString()
        'custid.Width = 150


        'Dim custname As New DataGridTextBoxColumn
        'custname.HeaderText = "Customer_Name"
        'custname.NullText = ""
        'custname.MappingName = stockDC.Tables(0).Columns(1).ToString()
        'custname.Width = 500

        'Dim custcity As New DataGridTextBoxColumn
        'custcity.HeaderText = "Customer_City"
        'custcity.NullText = ""
        'custcity.MappingName = stockDC.Tables(0).Columns(2).ToString()
        'custcity.Width = 150

        'Dim csr As New DataGridTextBoxColumn
        'csr.HeaderText = "CSR"
        'csr.NullText = ""
        'csr.MappingName = stockDC.Tables(0).Columns(3).ToString()
        'csr.Width = 150


        DataGridCustomer.DataSource = stockDC.Tables(0)
        sqlcon.Close()
        DataGridCustomer.Expand(-1)
        'DataGrid1.TableStyles("Employees").GridColumnStyles("Title").Width = newwidth
        'DataGridCustomer.TableStyles("Table").GridColumnStyles("Customer").Width = 400


    End Sub

    Sub EnquiryEdit()

        DataGridEnquiryEdit.Show()


        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)



        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        ' Dim cmSQL As SqlCommand
        ' Dim drSQL As SqlDataReader

        strSql = "SELECT  Reg_No, Reg_Date, Enq_Ref_no, Enq_Ref_date, CustomerID, CustomerName FROM TSS_Enquiry_Edit order by Reg_No"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridEnquiryEdit.Width = 650

        DataGridEnquiryEdit.Height = 320


        DataGridEnquiryEdit.DataSource = stockDC.Tables(0)
        sqlCon.Close()
        DataGridEnquiryEdit.Expand(-1)


    End Sub


    Public Sub fillPartnumbers()

        DataGridPartNumbers.Show()

        'Dim sqlcon As New SqlConnection

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)

        'Dim sqlCon As SqlConnection = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"])

        'SQLConnection objConn = new SQLConnection(System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"]);


        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader

        txtpart.Text = txtpart.Text & "%"
        strSql = "SELECT ItemNumber, ItemDescription, ItemUM,ItemKey FROM FSDBBR.dbo.FS_Item " & _
                 "WHERE  ItemNumber like '" & txtpart.Text & "' " & _
                    "ORDER BY ItemNumber"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDAC.TableMappings.Add("Table", "Part")
        'get data
        stockDAC.Fill(stockDC)

        DataGridPartNumbers.Width = 650 '1000
        DataGridPartNumbers.Height = 320 ' 412


        DataGridPartNumbers.DataSource = stockDC.Tables(0)
        sqlCon.Close()
        DataGridPartNumbers.Expand(-1)


    End Sub

    Public Sub fillqty()

        DataGridQty.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet

        'Dim cmSQL As SqlCommand
        ' Dim drSQL As SqlDataReader

        ' If lblMode1.Text = "Add" Then
        strSql = "SELECT Qty,Qty_Type, Enq_Qty_IntCode FROM ENQ_Qty_Details " & _
                 "WHERE  Enq_Int_code=  '" & txtenqintcode.Text & "' and Enq_Detail_code = '" & txtdetailintcode.Text & "' " & _
                    "ORDER BY Qty"



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


    Sub fillDetailList()

        'datagridDetail.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDC As DataSet = New DataSet


        strSql = "SELECT Sl_no,FS_Yes_NO,Part_Source,PartNumber,PartDescription,CustPartNumber,CustPartDescription,uom,RecomVendor,Dimension,Material,Special,Enq_Detail_code,Req FROM ENQ_Details where Enq_Int_code = " & txtenqintcode.Text & " order by Sl_no"



        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd

        sqlCon.Open()
        stockDAC.TableMappings.Add("Table", "Part")
        'get data
        stockDAC.Fill(stockDC)


        datagridDetail.DataSource = stockDC.Tables(0)
        sqlCon.Close()
        datagridDetail.Expand(-1)

        'End If
        'End If


    End Sub











    Private Sub ListBoxCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub


    Private Sub DataGridCustomer_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridCustomer.Navigate

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtContact.TextChanged

    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox18_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPhone.TextChanged

    End Sub

    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub DataUpdation_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataUpdation.Enter

    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label41.Click

    End Sub

    Private Sub Label37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label37.Click

    End Sub

    Private Sub Label36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label36.Click

    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub TextBox29_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotalItems.TextChanged

    End Sub

    Private Sub TextBox28_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRecVend.TextChanged

    End Sub

    Private Sub Label46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label46.Click

    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxSource.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxSource_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxSource.GotFocus

    End Sub

    Public Sub COMBOLOAD()
        '-------------------SOURCE
        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand
        ' Dim drSQL As SqlDataReader
        strSql = "SELECT * FROM ENQ_Source " & _
                 "WHERE Status like 'A%' ORDER BY Source"
        cmSQL = New SqlCommand(strSql, sqlCon)
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource As SqlDataAdapter = New SqlDataAdapter
        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eSource")
        With ComboBoxSource
            .DataSource = source.Tables("eSource")
            .DisplayMember = "Source"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With
        'tax

        strSql = "SELECT * from ENQ_Tax " & _
                 "WHERE Status like 'A%' ORDER BY Tax_Details"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eclarity")
        With ComboBoxTax
            .DataSource = source.Tables("eclarity")
            .DisplayMember = "Tax_Details"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With



        '--------------------Clarity
        strSql = "SELECT * FROM ENQ_Clarity " & _
                   "WHERE Status like 'A%' ORDER BY Clarity"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eclarity")
        With ComboBoxClarity
            .DataSource = source.Tables("eclarity")
            .DisplayMember = "Clarity"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With
        '-----------------Enq Type

        strSql = "SELECT * FROM ENQ_Type " & _
                        "WHERE Status like 'A%' ORDER BY Type"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet

        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "etype")
        With ComboBoxEnquiryType
            .DataSource = source.Tables("etype")
            .DisplayMember = "Type"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With
        '-------------------PriceStatus

        strSql = "SELECT * FROM ENQ_PriceStatus " & _
                        "WHERE Status like 'A%' ORDER BY PriceStatus"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "estatus")
        With ComboBoxPriceStatus
            .DataSource = source.Tables("estatus")
            .DisplayMember = "PriceStatus"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With


        '-------------------Enquiry Category

        strSql = "SELECT * FROM ENQ_Category " & _
                        "WHERE Status like 'A%' ORDER BY Category"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "ecat")
        With ComboBoxCategory()
            .DataSource = source.Tables("ecat")
            .DisplayMember = "Category"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With

        '-------------------Enquiry forward

        strSql = "SELECT * FROM ENQ_Forward " & _
                                "WHERE Status like 'A%' ORDER BY Forward"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "efar")
        With ComboBoxForward()
            .DataSource = source.Tables("efar")
            .DisplayMember = "Forward"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With

        '-------------------Enquiry status
        strSql = "SELECT * FROM ENQ_Status " & _
                                "WHERE Status like 'A%' ORDER BY ENQ_Status"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eestatus")
        With ComboBoxEnquiryStatus()
            .DataSource = source.Tables("eestatus")
            .DisplayMember = "ENQ_Status"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With


        '-------------------status remarks
        strSql = "SELECT * FROM ENQ_Status_Remarks " & _
                                "WHERE Status like 'A%' ORDER BY Remarks"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eremarks")
        With ComboBoxStatusRemarks()

            .DataSource = source.Tables("eremarks")
            .DisplayMember = "Remarks"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With
        '-------------------Rejections
        strSql = "SELECT * FROM ENQ_Reason_Rejections " & _
                                "WHERE Status like 'A%' ORDER BY Reasons"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "ere")
        With ComboBoxRejectionReasons()
            .DataSource = source.Tables("ere")
            .DisplayMember = "Reasons"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With


        '-------------------Class
        strSql = "SELECT * FROM ENQ_CustClass " & _
                                "WHERE Status like 'A%' ORDER BY Class"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eclass")
        With ComboboxClass()
            .DataSource = source.Tables("eclass")
            .DisplayMember = "Class"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With


        '-------------------csr
        strSql = "SELECT * FROM ENQ_CSR " & _
                                "WHERE Status like 'A%' ORDER BY CSR"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "ecsr")
        With ComboBoxCSR()
            .DataSource = source.Tables("ecsr")
            .DisplayMember = "CSR"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With

        '-------------------isr
        strSql = "SELECT * FROM ENQ_ISR " & _
                                "WHERE Status like 'A%' ORDER BY ISR"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eisr")
        With ComboboxISR()
            .DataSource = source.Tables("eisr")
            .DisplayMember = "ISR"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With

        '-------------------tssiseg
        strSql = "SELECT * FROM ENQ_TSSISeg " & _
                                "WHERE Status like 'A%' ORDER BY Seg"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eseg")
        With ComboboxTSSISeg()
            .DataSource = source.Tables("eseg")
            .DisplayMember = "Seg"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With


        '-------------------tssseg
        strSql = "SELECT * FROM ENQ_TSSSeg " & _
                                "WHERE Status like 'A%' ORDER BY TSSSeg"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet


        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "esegment")
        With ComboboxSegment()
            .DataSource = source.Tables("esegment")
            .DisplayMember = "TSSSeg"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With



        Exit Sub



        'ComboBoxSource.Items.Clear()

    End Sub


    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        enablecustomerdata()
        clearcustomerdata()


    End Sub

    Private Sub txtVat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVat.TextChanged

    End Sub

    Private Sub btnHeaderSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHeaderSave.Click

        '  mailcustcreation()


        Dim checkdt As Date
        checkdt = Today

        'If Month(DTPRegDt.Value) = checkdt.Month Then

        'Else
        'MsgBox("Enq Reg Date is not saving properly, please contact the IT deptarment, IT has to check date in regional setting,", MsgBoxStyle.Critical)
        'Exit Sub


        'End If


        Dim strsql As String
        Dim strsql2 As String
        Dim cmSQL As SqlCommand
        Dim msgb As String
        Dim a As Integer
        Dim cust As String
        Dim tender As String
        Dim doc As String
        Dim markettype As String


        If ComboBoxEnquiryType.Text = "Project" Or ComboBoxEnquiryType.Text = "Project-Budgetary" Then

            If Val(txtBusinessPotential.Text) = 0 Then
                PanelProjectDetails.Visible = True
                PanelProjectDetails.Location = New Point(426, 16)
                PanelProjectDetails.Width = 814
                PanelProjectDetails.Height = 244
                Exit Sub

            End If


        End If

        strsql = ""
        strsql2 = ""

        If Val(txtRegNo.Text) > 0 Then

            ' If usertype = "S" Then
            '  purcheck = 0
            'Else

            purcheck = 0
            purchasecheck()

            ' End If
        End If

        If (purcheck = 1 And usertype <> "S") Then
            Exit Sub

        Else
            msgb = MsgBox("Are you sure of saving ?", vbYesNo)

            If msgb = vbYes Then


                If ComboBoxEnquiryType.Text = "Project" Or ComboBoxEnquiryType.Text = "Project-Budgetary" Then
                    If ComboBoxForward.Text <> "Forward to Apl. Dept" Then
                        MsgBox("Projects should be forwarded to Application Dept.", MsgBoxStyle.Exclamation)
                        Exit Sub
                    End If
                End If


                If ComboBoxEnquiryStatus.Text = "Closed" Then
                    MsgBox("Status closed is system generated,Don't select it", vbInformation)
                    Exit Sub

                End If


                If RBTenderYes.Checked = True Then
                    If dtpTenderDueDt.Checked = False Then
                        MsgBox("Tender due date to be selected", vbInformation)
                        Exit Sub
                    End If

                End If

                'End If

                If Len(txtCustomer.Text) <= 5 Then
                    MsgBox("Customer Name  should be selected or entered.", vbInformation)
                    Exit Sub
                End If

                If DTPEnqRecd.Value < dtpEnqDt.Value Then
                    MsgBox("Enquiry recd date should be greater or equal to Enquiry date.", vbInformation)
                    Exit Sub

                End If

                If RBTenderYes.Checked = True Then
                    If dtpTenderDueDt.Value < dtpEnqDt.Value Then
                        MsgBox("Tender due date should be greater or equal to Enquiry date.", vbInformation)
                        Exit Sub

                    End If
                End If

                If txtPartYesPriceYes.Text = "" Then
                    txtPartYesPriceYes.Text = 0

                End If


                If txtRejected.Text = "" Then
                    txtRejected.Text = 0

                End If

                If txtPartNot.Text = "" Then
                    txtPartNot.Text = 0
                End If


                If txtPriceNot.Text = "" Then
                    txtPriceNot.Text = 0
                End If

                If txtBothNot.Text = "" Then
                    txtBothNot.Text = 0
                End If


                a = Val(txtPartYesPriceYes.Text) + Val(txtPartNot.Text) + Val(txtPriceNot.Text) + Val(txtRejected.Text) + Val(txtBothNot.Text)

                If a <> Val(txtTotalItems.Text) Then
                    MsgBox("Part number count is not matching", vbInformation)
                    Exit Sub
                ElseIf a = 0 Then
                    MsgBox("Total Number of Part should not be blank", vbInformation)
                    Exit Sub

                End If

                If ComboBoxSource.Text = "" Then
                    MsgBox("Enquiry source should not be blank ", vbInformation)

                    Exit Sub
                End If


                If txtEnqRef.Text = "" Then
                    MsgBox("Enquiry reference number should not be blank", vbInformation)
                    Exit Sub
                End If



                If rbdocyes.Checked = True Then
                    If Len(txtDocDetails.Text) <= 5 Then
                        MsgBox("Document details to be entered.", vbInformation)
                        Exit Sub

                    End If
                End If

                If RBCustomerExisting.Checked = False Then
                    If Len(txtCustPin.Text) = 0 Then
                        txtCustPin.Text = 0

                    End If
                End If



                If RBCustomerExisting.Checked = True Then
                    cust = "YES"
                Else
                    cust = "NO"
                End If

                If RadioButtonDomestic.Checked = True Then
                    markettype = "Domestic"
                Else
                    markettype = "Export"
                End If


                If RBTenderYes.Checked = True Then
                    tender = "YES"
                Else
                    tender = "NO"
                End If

                If rbdocyes.Checked = True Then
                    doc = "YES"
                Else
                    doc = "NO"
                End If

                'generating regno


                Dim tdate As Date

                If dtpTenderDueDt.Checked = False Then
                    tdate = "01-01-1900"
                Else
                    tdate = dtpTenderDueDt.Value

                End If

                If ComboBoxEnquiryStatus.Text = "Accepted" Then
                    If DTPStatusDt.Checked = False Then
                        MsgBox("Enquiry accepted date should be entered,vbinformatrion")
                        Exit Sub

                    End If

                End If


                If lblMode1.Text = "Add" Then

                    Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)


                    If RBCustomerNew.Checked = True And Len(txtCustID.Text) < 3 And Len(txtCustomer.Text) > 5 Then

                        custintcodegen()

                    Else
                        txtcustintcode.Text = 0
                    End If

                    enqregno()
                    enqinternalcode()


                    If txtRegNo.Text > 0 Then

                        cnSQL.Open()


                        If RBCustomerNew.Checked = True And Len(txtCustID.Text) < 3 And Len(txtCustomer.Text) > 5 Then

                            curdate = System.DateTime.Now()


                            strsql2 = "insert ENQ_New_Customers values (" & txtcustintcode.Text & ",'" & txtCustomer.Text & "', '" & txtCustAd1.Text & "','" & txtCustAdr2.Text & "'," & _
                                        "'" & txtCustAdr3.Text & "','" & txtCustcity.Text & "'," & txtCustPin.Text & ",'" & txtCustState.Text & "','" & txtCustCountry.Text & "','" & txtContact.Text & "', " & _
                                        "'" & txtDesignation.Text & "','" & txtDept.Text & "', '" & txtMobile.Text & "', '" & txtPhone.Text & "', '" & txtFax.Text & "','" & txtemail.Text & "', '" & txtEcc.Text & "'," & _
                                        "'" & txtVat.Text & "','" & txtCst.Text & "','" & txtRemarks.Text & "','" & ComboBoxClass3.Text & "','" & ComboBoxCSR.Text & "','" & ComboboxISR.Text & "'," & _
                                        "'" & ComboboxTSSISeg.Text & "','" & ComboboxSegment.Text & "','-','" & curdate & "','" & curdate & "','" & txtDunsno.Text & "','" & ComboBoxTax.Text & "','" & ComboboxClass.Text & "','" & markettype & "','01-01-1900','" & ComboBoxCurrency.Text & "')"


                            cmSQL = New SqlCommand(strsql2, cnSQL)


                            If cmSQL.ExecuteNonQuery() = 0 Then
                                MsgBox("Cannot Save Customer Details. " & strsql, MsgBoxStyle.Exclamation, "Error!")
                                'txtRegNo.Text = 0
                                Application.Exit()

                            End If

                        End If

                        'end customer saving

                        'enq detail section saving

                        curdate = System.DateTime.Now()

                        strsql = "insert ENQ_Header values(" & txtenqintcode.Text & "," & txtRegNo.Text & ",'" & DTPRegDt.Value & "', '" & txtEnqRef.Text & "'," & _
                        "'" & dtpEnqDt.Value & "', '" & DTPEnqRecd.Value & "',' " & ComboBoxSource.Text & " ', '" & tender & "'," & _
                        "'" & tdate & "'," & _
                        "'" & cust & "'," & txtcustintcode.Text & ",'" & txtCustID.Text & "','" & ComboBoxCategory.Text & "', '" & ComboBoxEnquiryType.Text & "'," & _
                        "'" & ComboBoxClarity.Text & "','" & ComboBoxForward.Text & "', '" & ComboBoxPriceStatus.Text & "'," & _
                        "'" & ComboBoxEnquiryStatus.Text & "', '" & DTPStatusDt.Value & "', '" & ComboBoxStatusRemarks.Text & "','" & ComboBoxRejectionReasons.Text & "'," & _
                        " " & txtTotalItems.Text & ", " & txtPartYesPriceYes.Text & ", " & txtPartNot.Text & ", " & txtPriceNot.Text & ", " & txtRejected.Text & ", " & _
                        "'" & doc & "','" & txtSpecial.Text & "', '" & txtDocDetails.Text & "','" & curdate & "','" & curdate & "','" & username & "','" & markettype & "'," & txtBothNot.Text & ")"

                        'cnSQL.Open()
                        cmSQL = New SqlCommand(strsql, cnSQL)

                        If cmSQL.ExecuteNonQuery() = 0 Then
                            MsgBox("Cannot Save Header Section. " & strsql, MsgBoxStyle.Exclamation, "Error!")
                            txtRegNo.Text = 0
                            Application.Exit()

                        Else

                            'SaveCertDetails()
                            MsgBox("Header section saved.", vbInformation)
                            'update the regno.back to table
                            enqregupdate()
                            btnHeaderSave.Enabled = False

                            Exit Sub
                        End If

                    End If

                    'Else
                    '   Exit Sub
                    '  End If


                ElseIf lblMode1.Text = "Edit" Then


                    Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
                    cnSQL.Open()


                    If RBTenderYes.Checked = True Then
                        tender = "YES"
                    Else
                        tender = "NO"
                    End If

                    If rbdocyes.Checked = True Then
                        doc = "YES"
                    Else
                        doc = "NO"
                    End If


                    If RBCustomerNew.Checked = True And Len(txtCustID.Text) < 3 And Len(txtCustomer.Text) > 5 Then

                        ' strSQL = "update FSPrograms.dbo.CE_InvoiceDetail set CustomPartNo = '" & drSQL.Item(0) & "',  Rate = " & Format(drSQL.Item(1), "#.000") & " WHERE ItemNumber = '" & drSQL.Item(2) & "' AND ShipmentNumber = " & dNO & ""

                        curdate = System.DateTime.Now()

                        strsql2 = "update ENQ_New_Customers set " & _
                         "Name	= '" & txtCustomer.Text & "'," & _
                         "Addr1	= '" & txtCustAd1.Text & "'," & _
                         "Addr2	= '" & txtCustAdr2.Text & "'," & _
                         "Addr3	= '" & txtCustAdr3.Text & "'," & _
                         "City	=  '" & txtCustcity.Text & "'," & _
                         "Pin = " & txtCustPin.Text & ", " & _
                         "State	=  '" & txtCustState.Text & "'," & _
                         "Country	=  '" & txtCustCountry.Text & "'," & _
                         "ContactPerson	= '" & txtContact.Text & "'," & _
                         "Designation	= '" & txtDesignation.Text & "'," & _
                         "Mobile	= '" & txtMobile.Text & "'," & _
                         "Phone	= '" & txtPhone.Text & "'," & _
                         "Fax	= '" & txtFax.Text & "'," & _
                         "Email	= '" & txtemail.Text & "'," & _
                         "Ecc = '" & txtEcc.Text & "'," & _
                         "Vat	= '" & txtVat.Text & "'," & _
                         "CST	=  '" & txtCst.Text & "'," & _
                         "Remarks	= '" & txtRemarks.Text & "'," & _
                         "Class3	= '" & ComboBoxClass3.Text & "'," & _
                         "CSR	= '" & ComboBoxCSR.Text & "'," & _
                         "ISR	= '" & ComboboxISR.Text & "'," & _
                         "TSSISeg =  '" & ComboboxTSSISeg.Text & "'," & _
                         "TSSSeg	= '" & ComboboxSegment.Text & "'," & _
                         "CustomerID = '" & txtCustID.Text & "'," & _
                         "Date_Modify= '" & curdate & "'," & _
                         "Duns_No= '" & txtDunsno.Text & "'," & _
                         "Tax_Type = '" & ComboBoxTax.Text & "'," & _
                         "Class1 = '" & ComboboxClass.Text & "'," & _
                         "MarketType = '" & markettype & "'," & _
                         "Currency = '" & ComboBoxCurrency.Text & "'" & _
                        " where Cust_IntCode = " & txtcustintcode.Text & ""
                        cmSQL = New SqlCommand(strsql2, cnSQL)

                        If cmSQL.ExecuteNonQuery() = 0 Then
                            MsgBox("Cannot Save edited customer details. " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                            'txtRegNo.Text = 0
                            Application.Exit()

                        End If

                    End If


                    curdate = System.DateTime.Now()
                    If purcheck = 0 Then


                        strsql = "update ENQ_Header set " & _
                        "Enq_Ref_no	= '" & txtEnqRef.Text & "'," & _
                        "Enq_Ref_date =	'" & dtpEnqDt.Value & "'," & _
                        "Enq_Recd_date = '" & DTPEnqRecd.Value & "'," & _
                        "Enq_Source	= '" & ComboBoxSource.Text & "'," & _
                        "Tender_YesNo = '" & tender & "'," & _
                        "Enq_Due_date =	'" & dtpTenderDueDt.Value & "'," & _
                        "CustomerID	= '" & txtCustID.Text & "'," & _
                        "Enq_Category	= '" & ComboBoxCategory.Text & "'," & _
                        "Enq_Type	= '" & ComboBoxEnquiryType.Text & "'," & _
                        "Enq_Clarity  =	'" & ComboBoxClarity.Text & "'," & _
                        "Enq_Forward = 	'" & ComboBoxForward.Text & "'," & _
                        "Price_Status = '" & ComboBoxPriceStatus.Text & "'," & _
                        "Enq_Status	= 	'" & ComboBoxEnquiryStatus.Text & "'," & _
                        "Enq_Status_dt =	'" & DTPStatusDt.Value & "'," & _
                        "Enq_Status_Remarks = '" & ComboBoxStatusRemarks.Text & "'," & _
                        "Reason_Rejection	= '" & ComboBoxRejectionReasons.Text & "'," & _
                        "Total_no =	'" & txtTotalItems.Text & "'," & _
                        "Part_Price_Yes =	'" & txtPartYesPriceYes.Text & "'," & _
                        "Part_No	= '" & txtPartNot.Text & "'," & _
                        "price_No	= '" & txtPriceNot.Text & "'," & _
                        "Rejected	= '" & txtRejected.Text & "'," & _
                        "Doc_upload	= '" & doc & "'," & _
                        "Doc_Details =	'" & txtDocDetails.Text & "'," & _
                        "Special_instructions =	'" & txtSpecial.Text & "'," & _
                        "Date_Modify = '" & curdate & "'," & _
                        "MarketType = '" & markettype & "'," & _
                        "BothNot = " & txtBothNot.Text & "" & _
                        " where Enq_Reg_NO = " & txtRegNo.Text & ""

                    ElseIf purcheck = 1 And usertype = "S" Then

                        strsql = "update ENQ_Header set " & _
                        "Enq_Ref_no	= '" & txtEnqRef.Text & "'," & _
                        "Enq_Ref_date =	'" & dtpEnqDt.Value & "'," & _
                        "Enq_Recd_date = '" & DTPEnqRecd.Value & "'," & _
                        "Enq_Source	= '" & ComboBoxSource.Text & "'," & _
                        "Tender_YesNo = '" & tender & "'," & _
                        "Enq_Due_date =	'" & dtpTenderDueDt.Value & "'," & _
                        "CustomerID	= '" & txtCustID.Text & "'," & _
                        "Enq_Category	= '" & ComboBoxCategory.Text & "'," & _
                        "Enq_Type	= '" & ComboBoxEnquiryType.Text & "'," & _
                        "Enq_Clarity  =	'" & ComboBoxClarity.Text & "'," & _
                        "Enq_Forward = 	'" & ComboBoxForward.Text & "'," & _
                        "Price_Status = '" & ComboBoxPriceStatus.Text & "'," & _
                        "Enq_Status	= 	'" & ComboBoxEnquiryStatus.Text & "'," & _
                        "Enq_Status_Remarks = '" & ComboBoxStatusRemarks.Text & "'," & _
                        "Reason_Rejection	= '" & ComboBoxRejectionReasons.Text & "'," & _
                        "Total_no =	'" & txtTotalItems.Text & "'," & _
                        "Part_Price_Yes =	'" & txtPartYesPriceYes.Text & "'," & _
                        "Part_No	= '" & txtPartNot.Text & "'," & _
                        "price_No	= '" & txtPriceNot.Text & "'," & _
                        "Rejected	= '" & txtRejected.Text & "'," & _
                        "Doc_upload	= '" & doc & "'," & _
                        "Doc_Details =	'" & txtDocDetails.Text & "'," & _
                        "Special_instructions =	'" & txtSpecial.Text & "'," & _
                        "Date_Modify = '" & curdate & "'," & _
                        "MarketType = '" & markettype & "'," & _
                        "BothNot = " & txtBothNot.Text & "" & _
                        " where Enq_Reg_NO = " & txtRegNo.Text & ""

                        '   "Enq_Status_dt =	'" & DTPStatusDt.Value & "'," & _

                    End If


                    cmSQL = New SqlCommand(strsql, cnSQL)

                    If cmSQL.ExecuteNonQuery() = 0 Then
                        MsgBox("Cannot Save Header Section. " & strsql, MsgBoxStyle.Exclamation, "Error!")
                        txtRegNo.Text = 0
                        Application.Exit()

                    Else
                        MsgBox("Header section saved.", vbInformation)
                        'update the regno.back to table
                        btnHeaderSave.Enabled = False

                        'mail to finance
                        If purcheck = 0 And ComboBoxEnquiryStatus.Text = "Accepted" And RBCustomerNew.Checked = True Then

                            mailcustcreation()

                        End If

                        If purcheck = 0 And ComboBoxEnquiryStatus.Text = "Accepted" Then

                            mailEnquiryReg()

                        End If


                        'mail to appplication dept



                        If purcheck = 0 And ComboBoxEnquiryStatus.Text = "Accepted" Then

                            mailpartcreation()

                        End If

                        Exit Sub
                    End If

                End If
            End If
        End If

    End Sub
    Public Sub enqregno()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String


        strSQL1 = "select Enq_Reg_No_lastUsed from ENQ_RegNo_Control"
        cnSQL1.Open()

        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then
            txtRegNo.Text = drSQL1.Item(0) + 1
        End If


    End Sub
    Public Sub enqregupdate()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim strSQL1 As String

        strSQL1 = "update ENQ_RegNo_Control set Enq_Reg_No_lastUsed = '" & txtRegNo.Text & "'"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        If cmSQL1.ExecuteNonQuery() = 0 Then
            MsgBox("Enquiry Reg.No. Not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
            Exit Sub
        End If



    End Sub
    Public Sub enqinternalcode()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(Enq_Int_code)from ENQ_Header"
        'strSQL1 = "Select *,max(Enq_Int_code)  from ENQ_Header"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        'If drSQL1.Item(0) = 0 Then
        '    txtenqintcode.Text = 1
        'Else
        '    txtenqintcode.Text = drSQL1.Item(1) + 1
        'End If

        If txtRegNo.Text = 1 Then
            txtenqintcode.Text = 1
        Else
            If drSQL1.Read() Then

                txtenqintcode.Text = drSQL1.Item(0) + 1
            End If


        End If


        ' If drSQL1.FieldCount = 1 Then
        'txtenqintcode.Text = 1
        'Else
        '   txtenqintcode.Text = drSQL1.Item(0) + 1
        'End If


    End Sub

    Public Sub custintcodegen()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(Cust_IntCode)from ENQ_New_Customers"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        ' If txtRegNo.Text = 1 Then
        'txtenqintcode.Text = 1
        ' Else
        '--

        'If IsDBNull(drSQL1.Item(0)) Then
        'txtdetailintcode.Text = 1
        'Else
        '   txtdetailintcode.Text = drSQL1.Item(0) + 1
        ' End If

        '--




        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtcustintcode.Text = 1
            Else


                txtcustintcode.Text = drSQL1.Item(0) + 1


            End If

        End If

        'End If

    End Sub

    Public Sub detailintcodegen()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(Enq_Detail_code)from ENQ_Details"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        '  If Val(txtRegNo.Text) = 1 Then
        ' txtdetailintcode.Text = 1
        'Else
        If drSQL1.Read() Then


            If IsDBNull(drSQL1.Item(0)) Then
                txtdetailintcode.Text = 1
            Else
                txtdetailintcode.Text = drSQL1.Item(0) + 1
            End If



            ' txtdetailintcode.Text = drSQL1.Item(0) + 1
        End If


        'End If


    End Sub

    Public Sub qtyintcode()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(Enq_Qty_IntCode)from ENQ_Qty_Details"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then


            If IsDBNull(drSQL1.Item(0)) Then
                txtqtyintcode.Text = 1
            Else
                txtqtyintcode.Text = drSQL1.Item(0) + 1
            End If

        End If





        'If txtRegNo.Text = 1 And txtdetailintcode.Text = 1 Then
        '    txtqtyintcode.Text = 1
        'Else
        '    If drSQL1.Read() Then

        '        txtqtyintcode.Text = drSQL1.Item(0) + 1
        '    End If


        ' End If


    End Sub

    Public Sub slnogen()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(Sl_no)from ENQ_Details where Enq_Int_code = " & txtenqintcode.Text & ""
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        If txtslno.Text = "" Then
            txtslno.Text = 1
        Else
            If drSQL1.Read() Then

                If IsDBNull(drSQL1.Item(0)) Then
                    txtslno.Text = 1
                Else

                    txtslno.Text = drSQL1.Item(0) + 1
                End If


            End If
        End If



    End Sub



    Private Sub Label52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label52.Click

    End Sub

    Private Sub txtDocNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub RBTenderNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBTenderNo.CheckedChanged
        If RBTenderNo.Checked = False Then
            dtpTenderDueDt.Checked = False

        End If
    End Sub

    Private Sub RBTenderYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBTenderYes.CheckedChanged
        If RBTenderYes.Checked = True Then
            dtpTenderDueDt.Checked = True

        End If
    End Sub

    Private Sub RBCustomerExisting_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If RBCustomerExisting.Checked = True Then
            txtCustID.Enabled = True
            clearcustomerdata()
            diablecustomerdata()

            Exit Sub


        End If
    End Sub
    Public Sub diablecustomerdata()
        'txtCustID.Enabled = False
        txtCustomer.Enabled = False
        txtCustAd1.Enabled = False
        txtCustAdr2.Enabled = False
        txtCustAdr3.Enabled = False
        txtCustcity.Enabled = False
        txtCustState.Enabled = False
        txtCustPin.Enabled = False
        txtContact.Enabled = False
        txtDesignation.Enabled = False
        txtDept.Enabled = False
        txtCustCountry.Enabled = False
        txtPhone.Enabled = False
        txtMobile.Enabled = False
        txtEcc.Enabled = False
        txtVat.Enabled = False
        txtCst.Enabled = False
        txtFax.Enabled = False
        txtemail.Enabled = False
        txtRemarks.Enabled = False

        ComboboxClass.Text = ""
        ComboboxClass.Enabled = False

        ComboBoxCSR.Text = ""
        ComboBoxCSR.Enabled = False

        ComboboxISR.Text = ""
        ComboboxISR.Enabled = False

        ComboboxTSSISeg.Text = ""
        ComboboxTSSISeg.Enabled = False

        ComboboxSegment.Text = ""
        ComboboxSegment.Enabled = False

        txtDunsno.Enabled = False

        ComboBoxTax.Text = ""
        ComboBoxTax.Enabled = False

        ComboBoxClass3.Text = ""
        ComboBoxClass3.Enabled = False

        ComboBoxCurrency.Text = ""
        ComboBoxCurrency.Enabled = False


    End Sub
    Public Sub clearcustomerdata()
        txtCustID.Text = ""
        txtCustomer.Text = ""
        txtCustAd1.Text = ""
        txtCustAdr2.Text = ""
        txtCustAdr3.Text = ""
        txtCustcity.Text = ""
        txtCustState.Text = ""
        txtCustPin.Text = ""
        txtContact.Text = ""
        txtDesignation.Text = ""
        txtDept.Text = ""
        txtCustCountry.Text = ""
        txtPhone.Text = ""
        txtMobile.Text = ""
        txtEcc.Text = ""
        txtVat.Text = ""
        txtCst.Text = ""
        txtFax.Text = ""
        txtemail.Text = ""
        txtRemarks.Text = ""
        ComboboxClass.Text = ""
        ComboBoxCSR.Text = ""
        ComboboxISR.Text = ""
        ComboboxTSSISeg.Text = ""
        ComboboxSegment.Text = ""
        txtDunsno.Text = ""
        ComboBoxCurrency.Text = ""

    End Sub


    Public Sub enablecustomerdata()
        txtCustID.Enabled = False
        txtCustomer.Enabled = True
        txtCustAd1.Enabled = True
        txtCustAdr2.Enabled = True
        txtCustAdr3.Enabled = True
        txtCustcity.Enabled = True
        txtCustState.Enabled = True
        txtCustPin.Enabled = True
        txtContact.Enabled = True
        txtDesignation.Enabled = True
        txtDept.Enabled = True
        txtCustCountry.Enabled = True
        txtPhone.Enabled = True
        txtMobile.Enabled = True
        txtEcc.Enabled = True
        txtVat.Enabled = True
        txtCst.Enabled = True
        txtFax.Enabled = True
        txtemail.Enabled = True
        txtRemarks.Enabled = True
        ComboboxClass.Enabled = True
        ComboBoxCSR.Enabled = True
        ComboboxISR.Enabled = True
        ComboboxTSSISeg.Enabled = True

        ComboboxSegment.Enabled = True
        txtDunsno.Enabled = True
        ComboBoxTax.Enabled = True
        ComboBoxClass3.Enabled = True
        ComboBoxCurrency.Enabled = True





    End Sub
    Public Sub detailsectiondisable()
        txtslno.Enabled = False
        ComboBoxFSYesNo.Enabled = False
        ComboBoxItemSource.Enabled = False
        txtpart.Enabled = False
        txtPartDesc.Enabled = False

    End Sub

    Public Sub detailsectionenable()
        txtslno.Enabled = True
        ComboBoxFSYesNo.Enabled = True
        ComboBoxItemSource.Enabled = True
        txtpart.Enabled = True
        txtPartDesc.Enabled = True

    End Sub
    Private Sub txtCustID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustID.TextChanged

    End Sub

    Private Sub txtCustID_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustID.Enter

    End Sub

    Private Sub txtCustID_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustID.DoubleClick
        DataGridCustomer.Visible = True
        fillcustomerlist()

    End Sub

    Private Sub DataGridCustomer_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridCustomer.KeyDown



    End Sub

    Private Sub DataGridCustomer_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridCustomer.CurrentCellChanged

        Dim a As Integer
        'Dim custid As String



        a = DataGridCustomer.CurrentCell.ColumnNumber()

        If a = 0 Then
            txtCustID.Text = DataGridCustomer.Item(DataGridCustomer.CurrentCell)

            txtCustomer.Text = DataGridCustomer.Item(DataGridCustomer.CurrentCell.RowNumber, 1)

            txtCustcity.Text = DataGridCustomer.Item(DataGridCustomer.CurrentCell.RowNumber, 2)

            txtCustID.Enabled = False


            'txtCustomer.Text = DataGridCustomer.Item(


        Else
            MsgBox("Click on CustomerID to select the customer", vbInformation)
            Exit Sub
        End If

        DataGridCustomer.Hide()

        '2.datagrid1.item(0,0)<-----it gets the first column/row data of your datagrid
        '3. 4.'if you want selected
        '5.datagrid1.item(datagrid1.currentcell.rownumber,0)<---it gets the selected row and the first column 



    End Sub

    Private Sub DataGridCustomer_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridCustomer.DoubleClick


        'Private Sub DataGridView1_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        'frm_dashboard.lbl_pnum_text.Text = DataGridView1.SelectedRows(0).Cells(0).Value
        'frm_dashboard.lbl_pname_text.Text = DataGridView1.SelectedRows(0).Cells(2).Value + " " + DataGridView1.SelectedRows(0).Cells(1).Value
        'frm_dashboard.lbl_paddress_text.Text = DataGridView1.SelectedRows(0).Cells(4).Value
        'frm_dashboard.Enabled = True

        'If DataGridView1.SelectedRows.Count <> 0 Then
        '    Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
        '    fName = row.Cells("fNameColumnName").Value
        '    sName = row.Cells("sNameColumnName").Value
        'End If


        'Me.Hide()

        ' If DataGridCustomer.slectedrows.count <> 0 Then

        'End If

        'End If

        'txtCustID.Text = row.cells("").value


    End Sub


    Private Sub Label27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label27.Click

    End Sub

    Private Sub Label34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label34.Click

    End Sub

    Private Sub txtTotalItems_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTotalItems.KeyPress

        'Dim allowedChars As String = "0123456789$,"& Chr(Keys.Back)


        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If


    End Sub

    Private Sub txtPartYesPriceYes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPartYesPriceYes.TextChanged

    End Sub

    Private Sub txtPartYesPriceYes_ReadOnlyChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPartYesPriceYes.ReadOnlyChanged

    End Sub

    Private Sub txtPartYesPriceYes_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPartYesPriceYes.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtPriceNot_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPriceNot.TextChanged

    End Sub

    Private Sub txtPriceNot_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPriceNot.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtPartNot_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPartNot.TextChanged

    End Sub

    Private Sub txtPartNot_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPartNot.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtRejected_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRejected.TextChanged

    End Sub

    Private Sub txtRejected_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRejected.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtCustPin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustPin.TextChanged

    End Sub

    Private Sub txtCustPin_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustPin.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxFSYesNo.SelectedIndexChanged

    End Sub

    Private Sub ComboBox14_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxItemSource.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxuom.SelectedIndexChanged

    End Sub

    Private Sub TextBox26_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtqty.TextChanged

    End Sub

    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click

        '  checkpriceavailability()

        '  RFQHistory_Load1()

        Dim strsql As String
        strsql = ""
        purcheck = 0
        Dim linest As String


        If Val(txtRegNo.Text) > 0 Then
            purchasecheckdetail()
        End If



        If purcheck = 1 Then
            Exit Sub
        Else

            If ComboBoxFSYesNo.Text = "Yes" And lblMode1.Text <> "Edit" Then
                If Val(txtitemkey.Text) = 0 Then
                    MsgBox("Select part number from list", vbInformation)
                    Exit Sub

                End If
            End If

            If ComboboxReq.Text = "Part Creation" Or ComboboxReq.Text = "Both" Then
                If CheckBoxPartCreation.Checked = False Then
                    MsgBox("Part details to be filled, for part number creation", vbInformation)
                    Exit Sub
                Else
                    If (ComboBoxItemType.Text) = "" And (ComboBoxBuyer.Text) = "" And (ComboBoxPlanner.Text) = "" Then
                        MsgBox("Part details to be filled, for part number creation", vbInformation)
                        Exit Sub
                    End If

                End If

            End If


            ' If ComboBoxReqType.Text = "" Then
            ' MsgBox("Select Qty Type", vbInformation)
            ' Exit Sub

            ' End If


            If ComboboxReq.Text = "" Then
                MsgBox("Requirement should not be blank", vbInformation)
                Exit Sub
            End If


            If txtslno.Text = "" Then
                Exit Sub
            Else



                Dim cmSQL As SqlCommand
                Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)



                If ComboBoxFSYesNo.Text = "Yes" Then
                    If Len(txtpart.Text) < 5 Then
                        MsgBox("Pl select the part number from the list", vbInformation)
                        Exit Sub

                    End If
                ElseIf ComboBoxFSYesNo.Text = "No" Then
                    If Len(txtPartDesc.Text) < 3 Then
                        MsgBox("Pl enter part description", vbInformation)
                        Exit Sub
                    End If

                End If


                If ComboBoxuom.Text = "" Then
                    MsgBox("UOM should not be blank.", vbInformation)
                    Exit Sub

                End If

                If txtrun.Text = "" Then
                    txtrun.Text = 0
                End If

                If txtfix.Text = "" Then
                    txtfix.Text = 0
                End If

                If txtinsp.Text = "" Then
                    txtinsp.Text = 0
                End If


                If Val(txtRegNo.Text) > 0 Then


                    If (ComboBoxEnquiryType.Text = "Project" Or ComboBoxEnquiryType.Text = "Project-Budgetary") And (ComboBoxForward.Text = "Forward to Apl. Dept") Then

                        linest = "T"  'PROJECT ENQUIRIES
                    Else
                        linest = "P"   'ALL PENDING FOR PRICE

                    End If




                    If lblMode1.Text = "Add" Or transmode = "EditAdd" Then

                        curdate = System.DateTime.Now()

                        strsql = "insert ENQ_Details values(" & txtenqintcode.Text & "," & txtdetailintcode.Text & "," & txtslno.Text & "," & _
                        "'" & ComboBoxFSYesNo.Text & "','" & ComboBoxItemSource.Text & "', '" & txtpart.Text & "','" & txtPartDesc.Text & "','" & txtCustPart.Text & "'," & _
                        "'" & txtCustDesc.Text & "','" & ComboBoxuom.Text & "','" & txtRecVend.Text & "','" & txtDimension.Text & "'," & _
                        "'" & txtMaterial.Text & "','" & txtDetailSpecial.Text & "','" & curdate & "','" & curdate & "','" & linest & "','" & username & "','" & ComboboxReq.Text & "'," & _
                        "'" & ComboBoxItemType.Text & "','" & ComboBoxPlanner.Text & "','" & ComboBoxBuyer.Text & "'," & txtrun.Text & ", " & txtfix.Text & "," & _
                        "" & txtinsp.Text & ",'" & txtchilditemDesc.Text & "', '" & ComboBoxprodline.Text & "','" & ComboBoxInvAc.Text & "'," & _
                        "'" & txtsp1.Text & "','" & txtsp2.Text & "','01-01-1900','','','')"


                    ElseIf lblMode1.Text = "Edit" Then

                        curdate = System.DateTime.Now()

                        strsql = "Update ENQ_Details set " & _
                        "FS_Yes_NO = '" & ComboBoxFSYesNo.Text & "'," & _
                        "Part_Source =  '" & ComboBoxItemSource.Text & "'," & _
                        "PartNumber =   '" & txtpart.Text & "'," & _
                        "PartDescription = '" & txtPartDesc.Text & "', " & _
                        "CustPartNumber = '" & txtCustPart.Text & "'," & _
                        "CustPartDescription = '" & txtCustDesc.Text & "'," & _
                        "uom = '" & ComboBoxuom.Text & "', " & _
                        "RecomVendor = '" & txtRecVend.Text & "'," & _
                        "Dimension = '" & txtDimension.Text & "'," & _
                        "Material =  '" & txtMaterial.Text & "'," & _
                        "Special =  '" & txtDetailSpecial.Text & "'," & _
                        "ItemStatus = '" & linest & "'," & _
                         "Date_Modify = '" & curdate & "'," & _
                        "Req = '" & ComboboxReq.Text & "'," & _
                        "Item_Type	= '" & ComboBoxItemType.Text & "'," & _
                        "Planner	='" & ComboBoxPlanner.Text & "'," & _
                        "Buyer	= '" & ComboBoxBuyer.Text & "'," & _
                        "Lead_Run = " & txtrun.Text & "," & _
                        "Lead_Fix	= " & txtfix.Text & "," & _
                        "Lead_Insp	= " & txtinsp.Text & "," & _
                        "Child_desc	= '" & txtchilditemDesc.Text & "'," & _
                        "prod_Line	=  '" & ComboBoxprodline.Text & "'," & _
                        "Inv_Ac	= '" & ComboBoxInvAc.Text & "'," & _
                        "Sp_note1	= '" & txtsp1.Text & "'," & _
                        "sp_note2	= '" & txtsp2.Text & "'" & _
                        " Where Enq_Int_code = " & txtenqintcode.Text & "  and " & _
                        " Enq_Detail_code = " & txtdetailintcode.Text & ""

                    End If

                    cnSQL.Open()

                    cmSQL = New SqlCommand(strsql, cnSQL)

                    If cmSQL.ExecuteNonQuery() = 0 Then
                        MsgBox("Cannot save detail section " & strsql, MsgBoxStyle.Exclamation, "Error!")
                        Application.Exit()

                    Else
                        MsgBox("Detail section saved.", vbInformation)
                        btnsave.Enabled = False

                        SaveCertDetails()

                        fillDetailList()
                        'fillqty()

                        Exit Sub
                    End If

                End If
            End If
            txtitemkey.Text = 0
        End If

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        listclearCertificate()
        'listloadCertificate()

        If txtRegNo.Text = "" Then
            MsgBox("Save the header section before entering part numbers !!", vbInformation)
            Exit Sub

        End If


        If lblMode1.Text = "Edit" Then
            transmode = "EditAdd"
            If btnQtyAdd.Visible = False Then
                btnQtyAdd.Visible = True
                btnQtyAdd.Enabled = True

            End If

            slnogen()

        End If

        qtycheck = False

        If Val(txtslno.Text) >= 1 Then
            checkqtyentered()

        End If
        If qtycheck = False Then


            If txtslno.Text = "" Then
                txtslno.Text = 1
            Else
                slnogen()

            End If

            lblqty.Text = txtslno.Text

            detailintcodegen()
            cleardetailsection()
            clearpartdesc()
            fillqty()



            detailsectionenable()
            btnsave.Enabled = True

        End If


    End Sub
    Public Sub checkqtyentered()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(Enq_Detail_code)from ENQ_Details where Enq_Detail_code not  in( Select Enq_Detail_code from ENQ_Qty_Details)and Enq_Int_code = '" & txtenqintcode.Text & "' "
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        '  If Val(txtslno.Text) <> 1 Then
        'If drSQL1.Read() And drSQL1.Item(0) Is Null Then
        'If drSQL1.Item(0) Is Null Then
        '        If drSQL1.Read() = True Then
        If drSQL1.Read() Then


            If IsDBNull(drSQL1.Item(0)) Then
                qtycheck = False

            Else

                MsgBox("Qty to be entered  before proceeding next sl.no.", vbInformation)
                qtycheck = True

                Exit Sub
            End If
        Else
            qtycheck = False
        End If


        'Else
        'qtycheck = False


        'End If


    End Sub
    Public Sub cleardetailsection()
        txtpart.Text = ""
        txtPartDesc.Text = ""
        txtCustPart.Text = ""
        txtCustDesc.Text = ""
        txtRecVend.Text = ""
        txtDimension.Text = ""
        txtMaterial.Text = ""
        txtDetailSpecial.Text = ""


    End Sub
    Private Sub txtEnqRef_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEnqRef.TextChanged

    End Sub

    Private Sub DataGridQty_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridQty.Navigate

    End Sub

    Private Sub datagridDetail_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles datagridDetail.Navigate

    End Sub

    Private Sub txtpart_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpart.TextChanged

    End Sub

    Private Sub DataGridPartNumbers_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub

    Private Sub DataGridPartNumbers_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim b As Integer
        ''Dim custid As String
        'b = DataGridPartNumbers.CurrentCell.ColumnNumber()

        'If b = 0 Then
        '    txtpart.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell)

        '    txtPartDesc.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell.RowNumber, 1)

        '    ComboBoxuom.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell.RowNumber, 2)

        '    txtCustID.Enabled = False


        '    'txtCustomer.Text = DataGridCustomer.Item(


        'Else
        '    MsgBox("Click on Part Number ", vbInformation)
        '    Exit Sub
        'End If

        'DataGridPartNumbers.Hide()


    End Sub

    Private Sub txtpart_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtpart.DoubleClick
        If ComboBoxFSYesNo.Text = "Yes" And Val(txtRegNo.Text) >= 1 Then
            DataGridPartNumbers.Visible = True

            fillPartnumbers()

            Exit Sub

        End If



    End Sub

    Private Sub DataGridPartNumbers_Navigate_1(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub

    Private Sub DataGridPartNumbers_CurrentCellChanged1(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Dim b As Integer
        ''Dim custid As String
        'b = DataGridPartNumbers.CurrentCell.ColumnNumber()

        'If b = 0 Then
        '    txtpart.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell)

        '    txtPartDesc.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell.RowNumber, 1)

        '    ComboBoxuom.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell.RowNumber, 2)

        '    txtCustID.Enabled = False


        '    'txtCustomer.Text = DataGridCustomer.Item(


        'Else
        '    MsgBox("Click on Part Number ", vbInformation)
        '    Exit Sub
        'End If

        'DataGridPartNumbers.Hide()



    End Sub

    Private Sub DataGridPartNumbers_Navigate_2(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub

    Private Sub DataGridPartNumbers_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs)

    End Sub

    Private Sub DataGridPartNumbers_Navigate_3(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridPartNumbers.Navigate

    End Sub

    Private Sub DataGridPartNumbers_CurrentCellChanged2(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridPartNumbers.CurrentCellChanged
        Dim b As Integer
        'Dim custid As String
        b = DataGridPartNumbers.CurrentCell.ColumnNumber()

        If b = 0 Then
            txtpart.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell)

            txtPartDesc.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell.RowNumber, 1)

            ComboBoxuom.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell.RowNumber, 2)
            txtitemkey.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell.RowNumber, 3)

            '        txtCustID.Enabled = False


            'txtCustomer.Text = DataGridCustomer.Item(


        Else
            MsgBox("Click on Part Number ", vbInformation)
            Exit Sub
        End If

        DataGridPartNumbers.Hide()

    End Sub

    Private Sub btnItemsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnItemsave.Click
        purcheck = 0
        purchasecheckdetail()

        If purcheck = 1 Then

            Exit Sub
        Else

            Dim strsql As String
            Dim cmSQL As SqlCommand
            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

            If ComboBoxReqType.Text = "" Then
                MsgBox("Requirement should not be blank", vbInformation)
                Exit Sub
            End If

            If lblMode1.Text = "Add" Or transmode = "EditQtyAdd" Then


                'Dim txtdetialint As Integer



                If Val(txtRegNo.Text) > 0 And Val(txtslno.Text) > 0 And Val(txtqty.Text) > 0 Then

                    qtyintcode()

                    curdate = System.DateTime.Now()

                    strsql = "insert ENQ_Qty_Details values(" & txtenqintcode.Text & "," & txtdetailintcode.Text & "," & txtqtyintcode.Text & "," & _
                    "" & txtqty.Text & ",'" & curdate & "','" & curdate & "', '" & username & "','" & ComboBoxReqType.Text & "')"
                    'Qty_Type

                    cnSQL.Open()

                    cmSQL = New SqlCommand(strsql, cnSQL)

                    If cmSQL.ExecuteNonQuery() = 0 Then
                        MsgBox("Cannot save qty " & strsql, MsgBoxStyle.Exclamation, "Error!")
                        Application.Exit()

                    Else
                        MsgBox("Qty saved.", vbInformation)

                        fillqty()
                        txtqty.Text = ""


                        Exit Sub
                    End If

                End If
            ElseIf lblMode1.Text = "Edit" Then

                If Val(txtRegNo.Text) > 0 And Val(txtslno.Text) > 0 And Val(txtqtyintcode.Text) > 0 Then

                    curdate = System.DateTime.Now()

                    strsql = "update ENQ_Qty_Details  set " & _
                    "Qty = " & txtqty.Text & ",Date_Modify = '" & curdate & "',Qty_Type = '" & ComboBoxReqType.Text & "' where Enq_Qty_IntCode  = " & txtqtyintcode.Text & ""

                    cnSQL.Open()

                    cmSQL = New SqlCommand(strsql, cnSQL)

                    If cmSQL.ExecuteNonQuery() = 0 Then
                        MsgBox("Cannot save qty " & strsql, MsgBoxStyle.Exclamation, "Error!")
                        Application.Exit()

                    Else
                        MsgBox("Qty saved.", vbInformation)

                        fillqty()
                        txtqty.Text = ""


                        Exit Sub
                    End If

                End If

            End If

        End If

    End Sub

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click

        If rbdocyes.Checked = True And Val(txtRegNo.Text) > 0 Then

            Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL1 As SqlCommand
            Dim drSQL1 As SqlDataReader
            Dim strSQL1 As String

            strSQL1 = "select Enq_Reg_NO from ENQ_Header where Enq_Reg_NO = " & txtRegNo.Text & " "
            cnSQL1.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
            drSQL1 = cmSQL1.ExecuteReader()


            If drSQL1.Read() Then

                With OpenFileDialog1

                    OpenFileDialog1.ShowDialog()


                    If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then


                    End If

                End With

            Else
                MsgBox("please save the enquiry header section before uploading the files", vbInformation)


                Exit Sub



            End If

        Else
            MsgBox("please save the enquiry header section before uploading the files", vbInformation)


            Exit Sub


        End If


        'Dim streamWriter As streamWriter = File.CreateText(filename)
        'streamWriter.Write("Your Text here")
        'streamWriter.Flush()
        'Dim FOL As New FolderBrowserDialog
        'FOL.ShowDialog()

        ''Imports System.Data.SqlClient
        ''Imports System.Text
        ''Imports Microsoft.VisualBasic
        ''Imports System.Net.WebRequest
        ''Imports System.Net.WebClient
        ''Imports System.Net
        ''Imports System.IO


        'If (Directory.Exists(directoryPath)) Then
        '    arrFileList = Directory.GetFiles(directoryPath) '''Getting All files 
        'Else
        '    MessageBox.Show("Directory not Exist")
        'End If

        'For i As Integer = 0 To i < arrFileList.Length - 1

        '    Try
        '        Dim fileName As String = arrFileList.GetValue(i)
        '        Dim toUpload As New FileInfo(fileName)
        '        Dim client As New WebClient

        '        Dim nc As New NetworkCredential("xxxx", "xxxx")

        '        Dim addy As Uri
        '        addy = New Uri("ftp://1xx.xx.xxx.xxx/HYPOSII_FTP/Test/" & toUpload.Name.ToString())


        '        client.Credentials = nc
        '        Dim arrReturn As Byte() = client.UploadFile(addy.ToString(), fileName) //This Line Throwing error
        '        MessageBox.Show("File Uploaded Sucessfully")
        '        FLAG = True
        '    Catch ex As Exception
        '        MessageBox.Show(ex.Message)
        '    End Try
        'Next

    End Sub

    Private Sub SaveFileDialog1_FileOk_1(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Dim filename As String
        'Dim A As String
        'A = txtRegNo.Text
        'Dim dialogue As New SaveFileDialog1
        'dialogue.FileName = "\\TSSBLRDOM111\Public\RFQ\" + A

        'dialogue.ShowDialog()

        SaveFileDialog1.ShowDialog()



        ' = dDilaoge.FileName
    End Sub

    Private Sub GroupYesNo_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupYesNo.Enter

    End Sub

    Private Sub rbdocyes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbdocyes.CheckedChanged
        If rbdocyes.Checked = True Then
            'btnuploadsave.Enabled = True
            btnUpload.Enabled = True

        End If

    End Sub

    Private Sub FolderBrowserDialog1_HelpRequest(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtCustAd1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustAd1.TextChanged

    End Sub

    Private Sub datagridDetail_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridDetail.CurrentCellChanged
        Dim c As Integer
        'Dim custid As String

        btnsave.Enabled = True


        c = datagridDetail.CurrentCell.ColumnNumber()

        If c = 0 Then
            lblMode1.Text = "Edit"


            txtslno.Text = datagridDetail.Item(datagridDetail.CurrentCell)

            ComboBoxFSYesNo.Text = Trim(datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 1))
            ComboBoxItemSource.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 2)
            txtpart.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 3)
            txtPartDesc.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 4)
            txtCustPart.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 5)
            txtCustDesc.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 6)
            ComboBoxuom.Text = Trim(datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 7))
            txtRecVend.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 8)
            txtDimension.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 9)
            txtMaterial.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 10)
            txtDetailSpecial.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 11)
            'ComboBoxReqType.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 12)
            txtdetailintcode.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 12)
            ComboboxReq.Text = datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 13)
            If Trim(datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 13)) = "Part Creation" Or Trim(datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 13)) = "Both" Then
                CheckBoxPartCreation.Visible = True
                CheckBoxPartCreation.Checked = True
                GroupBoxPartCreation.Visible = False

            End If

        Else
            MsgBox("Click on Sl.No. to select the Partnumber", vbInformation)
            Exit Sub
        End If

        listclearCertificate()
        EditCertDetails()

        If Trim(datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 13)) = "Part Creation" Or Trim(datagridDetail.Item(datagridDetail.CurrentCell.RowNumber, 13)) = "Both" Then
            fillpartdetails()

        End If
        'fillpartdetails()


        clearqty()
        If lblMode1.Text = "Add" Then
            btnQtyAdd.Visible = False
            btnItemsave.Enabled = False
            btnItemDelete.Enabled = True

        ElseIf lblMode1.Text = "Edit" Then
            btnQtyAdd.Visible = True
            btnQtyAdd.Enabled = True
            btnItemsave.Enabled = True

            btnItemDelete.Enabled = True
            detailsectionenable()
            '            fillqty()


        End If

    End Sub

    Private Sub btnItemDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnItemDelete.Click

        Dim msgb As String
        msgb = MsgBox("This item along with its qty slabs will be deleted, Are you sure ? ", vbYesNo)
        If msgb = vbNo Then
            Exit Sub
        Else

            Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL1 As SqlCommand
            'Dim drSQL1 As SqlDataReader
            Dim strSQL1 As String
            Dim strSQL2 As String
            Dim d As Integer
            d = 0
            cnSQL1.Open()

            strSQL1 = "Delete from ENQ_Details where Enq_Int_code = " & txtenqintcode.Text & " and Sl_no = " & txtslno.Text & " and Enq_Detail_code = " & txtdetailintcode.Text & " "
            strSQL2 = "Delete from ENQ_Qty_Details where Enq_Int_code = " & txtenqintcode.Text & " and Enq_Detail_code = " & txtdetailintcode.Text & ""

            cmSQL1 = New SqlCommand(strSQL1, cnSQL1)

            If cmSQL1.ExecuteNonQuery() = 0 Then
                d = 1
            End If

            cmSQL1 = New SqlCommand(strSQL2, cnSQL1)

            If cmSQL1.ExecuteNonQuery() = 0 Then
                d = d + 1
            End If

            If d <= 1 Then
                MsgBox("Selected record is deleted successfully", vbInformation)
                btnItemDelete.Enabled = False
                fillDetailList()
                fillqty()

                Exit Sub
            Else
                MsgBox("Selected record is not deleted", vbInformation)
                Exit Sub
            End If

            '      fillDetailList()

        End If


    End Sub

    Private Sub Enquiry_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown




    End Sub

    Private Sub txtEnqRef_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEnqRef.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If

    End Sub


    Private Sub dtpEnqDt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpEnqDt.ValueChanged

    End Sub

    Private Sub dtpEnqDt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpEnqDt.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxSource_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxSource.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub DTPEnqRecd_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPEnqRecd.ValueChanged

    End Sub

    Private Sub DTPEnqRecd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTPEnqRecd.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub RBTenderNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RBTenderNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub RBCustomerExisting_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustID.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustomer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomer.TextChanged

    End Sub

    Private Sub txtCustomer_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomer.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustAd1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustAd1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustAdr2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustAdr2.TextChanged

    End Sub

    Private Sub txtCustAdr2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustAdr2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustAdr3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustAdr3.TextChanged

    End Sub

    Private Sub txtCustAdr3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustAdr3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustcity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustcity.TextChanged

    End Sub

    Private Sub txtCustcity_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustcity.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustPin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustPin.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustState_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustState.TextChanged

    End Sub

    Private Sub txtCustState_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustState.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustCountry_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustCountry.TextChanged

    End Sub

    Private Sub txtCustCountry_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustCountry.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtContact_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtContact.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtDesignation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesignation.TextChanged

    End Sub

    Private Sub txtDesignation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDesignation.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtMobile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMobile.TextChanged

    End Sub

    Private Sub txtMobile_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMobile.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtPhone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPhone.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtPhone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPhone.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtFax_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFax.TextChanged

    End Sub

    Private Sub txtemail_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtemail.TextChanged

    End Sub

    Private Sub txtemail_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtemail.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtEcc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEcc.TextChanged

    End Sub

    Private Sub txtEcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEcc.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtVat_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVat.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCst_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCst.TextChanged

    End Sub

    Private Sub txtCst_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCst.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtRemarks_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRemarks.TextChanged

    End Sub

    Private Sub txtRemarks_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemarks.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboboxClass_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboboxClass.SelectedIndexChanged

    End Sub

    Private Sub ComboboxClass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboboxClass.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxCSR_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxCSR.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxCSR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxCSR.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxISR_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboboxTSSISeg_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboboxTSSISeg.SelectedIndexChanged

    End Sub

    Private Sub ComboboxTSSISeg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboboxTSSISeg.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboboxSegment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboboxSegment.SelectedIndexChanged

    End Sub

    Private Sub ComboboxSegment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboboxSegment.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxCategory.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxCategory_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxCategory.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxEnquiryType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEnquiryType.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxEnquiryType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxEnquiryType.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxClarity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxClarity.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxClarity_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxClarity.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxForward_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxForward.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxPriceStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxPriceStatus.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxEnquiryStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEnquiryStatus.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxPriceStatus_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxPriceStatus.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub DTPStatusDt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTPStatusDt.GotFocus
        If ComboBoxEnquiryStatus.Text = "Accepted" Then
            DTPStatusDt.Checked = True
            DTPStatusDt.Enabled = False
        End If

    End Sub

    Private Sub DTPStatusDt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPStatusDt.ValueChanged

    End Sub

    Private Sub DTPStatusDt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTPStatusDt.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxStatusRemarks_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxStatusRemarks.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxStatusRemarks_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxStatusRemarks.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxRejectionReasons_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxRejectionReasons.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxRejectionReasons_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxRejectionReasons.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtTotalItems_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotalItems.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtPartYesPriceYes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPartYesPriceYes.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtPriceNot_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPriceNot.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtPartNot_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPartNot.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtRejected_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRejected.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtDocDetails_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDocDetails.TextChanged

    End Sub

    Private Sub txtDocDetails_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDocDetails.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtSpecial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSpecial.TextChanged

    End Sub

    Private Sub txtSpecial_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSpecial.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxFSYesNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxFSYesNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxItemSource_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxItemSource.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
        If ComboBoxFSYesNo.Text = "Yes" Then
            txtpart.Text = "%"
        End If


    End Sub

    Private Sub txtpart_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtpart.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtPartDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPartDesc.TextChanged

    End Sub

    Private Sub txtCustPart_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustPart.TextChanged

    End Sub

    Private Sub txtCustPart_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustPart.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtCustDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustDesc.TextChanged

    End Sub

    Private Sub txtCustDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustDesc.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxuom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxuom.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtRecVend_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRecVend.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtDimension_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDimension.TextChanged

    End Sub

    Private Sub txtDimension_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDimension.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtMaterial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaterial.TextChanged

    End Sub

    Private Sub txtMaterial_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMaterial.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtDetailSpecial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDetailSpecial.TextChanged

    End Sub

    Private Sub txtDetailSpecial_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDetailSpecial.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub btnsave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnsave.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtqty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtqty.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub dtpTenderDueDt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpTenderDueDt.ValueChanged

    End Sub

    Private Sub ComboBoxForward_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxForward.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxEnquiryStatus_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxEnquiryStatus.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtFax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFax.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxISR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub btnItemsave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnItemsave.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub Label44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblEnqAdd.Click
        MsgBox("Adding enquiry from this screen is not ready,use main menu ", vbInformation)
        Exit Sub

    End Sub

    Private Sub DataGridEnquiryEdit_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridEnquiryEdit.Navigate

    End Sub

    Private Sub DataGridEnquiryEdit_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridEnquiryEdit.CurrentCellChanged


        Dim b As Integer
        'Dim custid As String

        b = DataGridEnquiryEdit.CurrentCell.ColumnNumber()

        If b = 0 Then

            txtRegNo.Text = DataGridEnquiryEdit.Item(DataGridEnquiryEdit.CurrentCell)
            DTPRegDt.Value = DataGridEnquiryEdit.Item(DataGridEnquiryEdit.CurrentCell.RowNumber, 1)

            txtEnqRef.Text = DataGridEnquiryEdit.Item(DataGridEnquiryEdit.CurrentCell.RowNumber, 2)
            dtpEnqDt.Value = DataGridEnquiryEdit.Item(DataGridEnquiryEdit.CurrentCell.RowNumber, 3)

            txtCustID.Text = DataGridEnquiryEdit.Item(DataGridEnquiryEdit.CurrentCell.RowNumber, 4)
            txtCustomer.Text = DataGridEnquiryEdit.Item(DataGridEnquiryEdit.CurrentCell.RowNumber, 5)

        Else
            MsgBox("Click on Reg. No. to Edit ", vbInformation)
            Exit Sub
        End If

        DataGridEnquiryEdit.Hide()

        EditEnquiryHeaderDetails()

        'EditCertDetails()

        If Val(txtcustintcode.Text) > 0 Then
            newcustdetails()
            enablecustomerdata()

        End If

        fillDetailList()

        purchasecheck()

        ' If purcheck = 1 Then
        'MsgBox("Purchase is working/closed this enquiry, so changes are not possible", vbInformation)
        'Exit Sub

        'End IfD

    End Sub
    Private Sub EditEnquiryHeaderDetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select * from ENQ_Header where Enq_Reg_NO = " & txtRegNo.Text & ""


        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            txtenqintcode.Text = drSQL1.Item(0)

            DTPEnqRecd.Value = drSQL1.Item(5)
            ComboBoxSource.Text = drSQL1.Item(6)

            If drSQL1.Item(7) = "YES" Then
                RBTenderYes.Checked = True

            Else
                RBTenderNo.Checked = True
            End If

            dtpTenderDueDt.Value = drSQL1.Item(8)

            If drSQL1.Item(9) = "YES" Then
                RBCustomerExisting.Checked = True
                RBCustomerNew.Checked = False
                txtCustID.Enabled = True
                diablecustomerdata()


            Else
                RBCustomerNew.Checked = True
                RBCustomerExisting.Checked = False
                txtCustID.Enabled = False
                enablecustomerdata()

            End If
            ComboBoxCategory.Text = drSQL1.Item(12)
            ComboBoxEnquiryType.Text = drSQL1.Item(13)

            ComboBoxClarity.Text = drSQL1.Item(14)
            ComboBoxForward.Text = drSQL1.Item(15)
            ComboBoxPriceStatus.Text = drSQL1.Item(16)
            ComboBoxEnquiryStatus.Text = drSQL1.Item(17)
            DTPStatusDt.Value = drSQL1.Item(18)
            ComboBoxStatusRemarks.Text = drSQL1.Item(19)
            ComboBoxRejectionReasons.Text = drSQL1.Item(20)
            txtTotalItems.Text = drSQL1.Item(21)
            txtPartYesPriceYes.Text = drSQL1.Item(22)
            txtPartNot.Text = drSQL1.Item(23)
            txtPriceNot.Text = drSQL1.Item(24)
            txtRejected.Text = drSQL1.Item(25)
            ' txtBothNot.Text = drSQL1.Item(0)
            If drSQL1.Item(26) = "YES" Then
                rbdocyes.Checked = True
            Else
                rbDocNo.Checked = True

            End If
            txtDocDetails.Text = drSQL1.Item(28)

            If IsDBNull(drSQL1.Item(27)) Then
                txtSpecial.Text = "-"
            Else
                txtSpecial.Text = drSQL1.Item(27)
            End If

            RBCustomerExisting.Enabled = False
            RBCustomerNew.Enabled = False



            If IsDBNull(drSQL1.Item(10)) Then
                txtcustintcode.Text = 0
            Else
                txtcustintcode.Text = drSQL1.Item(10)
            End If

            If (drSQL1.Item(28)) = "Domestic" Then
                RadioButtonDomestic.Checked = True
                RadioButtonExport.Checked = False

            ElseIf (drSQL1.Item(28)) = "Export" Then
                RadioButtonDomestic.Checked = False
                RadioButtonExport.Checked = True
            Else
                RadioButtonDomestic.Checked = True

            End If

            If IsDBNull((drSQL1.Item(33))) Then
                txtBothNot.Text = 0
            Else


                txtBothNot.Text = (drSQL1.Item(33))
            End If

        End If


    End Sub

    Private Sub newcustdetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select * from ENQ_New_Customers where Cust_IntCode = " & txtcustintcode.Text & ""
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            '0Cust_IntCode	numeric(18, 0)	Unchecked
            '1Name	varchar(60)	Unchecked
            '2Addr1	varchar(60)	Checked
            '3Addr2	varchar(60)	Checked
            '4Addr3	varchar(60)	Checked
            '5City	varchar(60)	Checked
            '6Pin	numeric(18, 0)	Checked
            '7State	varchar(60)	Checked
            '8Country	varchar(60)	Checked
            '9ContactPerson	varchar(60)	Checked
            '10Designation	varchar(60)	Checked
            '11Mobile	varchar(50)	Checked
            '12Phone	varchar(50)	Checked
            '13Fax	varchar(50)	Checked
            '14Email	varchar(70)	Checked
            '15Ecc	varchar(50)	Checked
            '16Vat	varchar(50)	Checked
            '17CST	varchar(50)	Checked
            '18Remarks	varchar(150)	Checked
            '19Class3	nchar(10)	Checked
            '20CSR	nchar(10)	Checked
            '21ISR	nchar(10)	Checked
            '22TSSISeg	nchar(10)	Checked
            '23TSSSeg	nchar(10)	Checked
            '24CustomerID	varchar(30)	Checked

            txtCustomer.Text = drSQL1.Item(1)

            txtCustAd1.Text = drSQL1.Item(2)
            txtCustAdr2.Text = drSQL1.Item(3)
            txtCustAdr3.Text = drSQL1.Item(4)
            txtCustcity.Text = drSQL1.Item(5)
            txtCustPin.Text = drSQL1.Item(6)
            txtCustState.Text = drSQL1.Item(7)
            txtCustCountry.Text = drSQL1.Item(8)
            txtContact.Text = drSQL1.Item(9)
            txtDesignation.Text = drSQL1.Item(10)
            txtDept.Text = drSQL1.Item(11)
            txtMobile.Text = drSQL1.Item(12)
            txtPhone.Text = drSQL1.Item(13)
            txtFax.Text = drSQL1.Item(14)
            txtemail.Text = drSQL1.Item(15)
            txtEcc.Text = drSQL1.Item(16)
            txtVat.Text = drSQL1.Item(17)
            txtCst.Text = drSQL1.Item(18)
            txtRemarks.Text = drSQL1.Item(19)
            ComboBoxClass3.Text = ""

            ComboBoxClass3.Text = Trim(drSQL1.Item(20))
            ComboBoxCSR.Text = ""
            ComboBoxCSR.Text = Trim(drSQL1.Item(21))
            'ComboBoxISR.Text = ""
            ComboboxISR.Text = Trim(drSQL1.Item(22))

            txtISR.Visible = True
            txtISR.Text = Trim(drSQL1.Item(22))

            ComboboxTSSISeg.Text = ""
            ComboboxTSSISeg.Text = Trim(drSQL1.Item(23))
            ComboboxSegment.Text = ""
            ComboboxSegment.Text = Trim(drSQL1.Item(24))
            txtDunsno.Text = Trim(drSQL1.Item(28))
            ComboBoxTax.Text = ""
            ComboBoxTax.Text = Trim(drSQL1.Item(29))
            ComboboxClass.Text = ""
            ComboboxClass.Text = Trim(drSQL1.Item(30))

            If IsDBNull(drSQL1.Item(31)) Then
                RadioButtonDomestic.Checked = True
                RadioButtonExport.Checked = False
            ElseIf drSQL1.Item(31) = "Export" Then
                RadioButtonExport.Checked = True
                RadioButtonDomestic.Checked = False
            ElseIf drSQL1.Item(31) = "Domestic" Then
                RadioButtonDomestic.Checked = True
                RadioButtonExport.Checked = False
            End If

            If IsDBNull(drSQL1.Item(33)) Then
                ComboBoxCurrency.Text = ""
            Else
                ComboBoxCurrency.Text = Trim(drSQL1.Item(33))
            End If


        End If


    End Sub

    Private Sub datagridDetail_ControlRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs) Handles datagridDetail.ControlRemoved

    End Sub

    Private Sub DataGridQty_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridQty.CurrentCellChanged

        'Dim custid As String
        Dim c As Integer



        c = datagridDetail.CurrentCell.ColumnNumber()

        If c = 0 Then
            txtqty.Text = DataGridQty.Item(DataGridQty.CurrentCell)
            ComboBoxReqType.Text = DataGridQty.Item(DataGridQty.CurrentCell.RowNumber, 1)

            txtqtyintcode.Text = DataGridQty.Item(DataGridQty.CurrentCell.RowNumber, 2)


        Else
            MsgBox("Click on Qty", vbInformation)
            Exit Sub
        End If



    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QtyEdit.Click
        transmode = ""
        fillqty()

    End Sub

    Private Sub btnQtyAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQtyAdd.Click
        transmode = "EditQtyAdd"
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub listloadCertificate()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
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
    Private Sub SaveCertDetails()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        cnSQL.Open()

        Dim i As Integer
        i = CheckedListBoxCertificate.CheckedItems.Count
        Dim a As Integer
        Dim cert As String

        ' a = 1

        'it it is edit mode delete and update again


        If lblMode1.Text = "Edit" Then

            strsql = " delete from ENQ_EnqWise_Certificates where Enq_Reg_NO = " & txtRegNo.Text & " and Enq_Detail_code = " & txtdetailintcode.Text & " and Enq_Int_code = " & txtenqintcode.Text & ""

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                'MsgBox("Cannot delete existing cetificate Details. " & strsql, MsgBoxStyle.Exclamation, "Error!")
                'Exit Sub
            End If

        End If




        Do While a < i
            cert = ""
            cert = CheckedListBoxCertificate.CheckedItems.Item(a)
            a = a + 1

            curdate = System.DateTime.Now()

            strsql = "insert ENQ_EnqWise_Certificates values (" & txtRegNo.Text & "," & txtdetailintcode.Text & " ," & txtenqintcode.Text & ", '" & cert & "','" & curdate & "','" & curdate & "', '" & username & "' )"

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot Save cetificate Details. " & strsql, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If


        Loop
    End Sub

    Private Sub EditCertDetails()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert As String
        'Dim b As Integer

        strsql = "Select Certificates from ENQ_EnqWise_Certificates where Enq_Reg_NO = " & txtRegNo.Text & " and Enq_Detail_code = " & txtdetailintcode.Text & " "



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

    Private Sub CheckedListBoxCertificate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBoxCertificate.SelectedIndexChanged

    End Sub


    Private Sub txtCustomer_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomer.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtEnqRef_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEnqRef.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCustAd1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustAd1.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCustAdr2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustAdr2.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCustAdr3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustAdr3.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCustcity_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustcity.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCustState_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustState.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCustCountry_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCountry.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtContact_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtContact.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtDesignation_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesignation.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtSpecial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSpecial.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtMobile_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMobile.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtFax_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFax.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtemail_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtemail.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtEcc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEcc.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtVat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVat.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCst.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtRemarks_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRemarks.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtDocDetails_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDocDetails.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtpart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtpart.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtPartDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPartDesc.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCustPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustPart.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtCustDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustDesc.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtRecVend_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRecVend.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtDimension_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDimension.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtMaterial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaterial.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtDetailSpecial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDetailSpecial.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtEnqRef_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEnqRef.DoubleClick
        ' DataGridCustomer.Visible = True
        'fillcustomerlist()

    End Sub

    Private Sub Label49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label49.Click

    End Sub

    Private Sub txtDept_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDept.TextChanged

    End Sub

    Private Sub txtDept_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDept.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ComboboxReq_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboboxReq.KeyDown

    End Sub

    Private Sub ComboboxReq_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboboxReq.LostFocus
        If ComboboxReq.Text = "Part Creation" Or ComboboxReq.Text = "Both" Then
            CheckBoxPartCreation.Visible = True

        End If
    End Sub

    Private Sub ComboboxReq_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboboxReq.SelectedIndexChanged

    End Sub
    Private Sub purchasecheck()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cnsql1 As SqlConnection = New SqlConnection (system.configurationManager.appsettings(["ConnectionString"]);

        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select Enq_Reg_NO from ENQ_RFQ_PriceDetails where Enq_Reg_NO = " & txtRegNo.Text & " AND Status <> 'Released to Customer Sup'"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        'drSQL1 = cmSQL1.ExecuteReader
        'drSQL1 = cmSQL1.ExecuteReader
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then

            Else
                MsgBox("Purchase is working/closed this enquiry, editing is not possible", vbInformation)
                purcheck = 1
                Exit Sub

            End If

        End If


    End Sub
    Private Sub purchasecheckdetail()
        purcheck = 0
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cnsql1 As SqlConnection = New SqlConnection (system.configurationManager.appsettings(["ConnectionString"]);

        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select Enq_Reg_NO from ENQ_RFQ_PriceDetails where Enq_Reg_NO = " & txtRegNo.Text & " AND Enq_Detail_code = " & txtdetailintcode.Text & " AND Status <> 'Released to Customer Sup'"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        'drSQL1 = cmSQL1.ExecuteReader
        'drSQL1 = cmSQL1.ExecuteReader
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then

            Else
                MsgBox("Purchase is working/closed this enquiry, editing this line is not possible", vbInformation)
                purcheck = 1
                Exit Sub

            End If

        End If


    End Sub


    Private Sub Label56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label56.Click

    End Sub

    Private Sub txtDunsno_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDunsno.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtDunsno_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDunsno.TextChanged

    End Sub

    Private Sub BothNot_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBothNot.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If

    End Sub

    Private Sub BothNot_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBothNot.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub BothNot_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBothNot.TextChanged

    End Sub
    Private Sub listclearCertificate()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert As String
        'Dim b As Integer

        'strsql = "Select Certificates from ENQ_EnqWise_Certificates where Enq_Reg_NO = " & txtRegNo.Text & " and Enq_Detail_code = " & txtdetailintcode.Text & " "

        strsql = "SELECT Certificates FROM ENQ_Certificates " & _
             "WHERE  Status = 'A'"

        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        i = CheckedListBoxCertificate.Items.Count

        Do While drSQL1.Read()
            cert = drSQL1.Item(0)
            a = 0
            Do While a < i

                If cert = CheckedListBoxCertificate.Items(a) Then

                    CheckedListBoxCertificate.SetItemChecked(a, False)

                    a = i
                Else

                    a = a + 1

                End If


            Loop

        Loop

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


    Private Sub RadioButtonDomestic_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonDomestic.CheckedChanged

    End Sub

    Private Sub GroupBoxPartCreation_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxPartCreation.Enter

    End Sub

    Private Sub CheckBoxPartCreate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBoxPartCreation.Visible = True
        GroupBoxPartCreation.Width = 1038
        GroupBoxPartCreation.Height = 116
    End Sub

    Private Sub CheckBoxPartCreation_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPartCreation.CheckedChanged
        If CheckBoxPartCreation.Checked = True Then
            GroupBoxPartCreation.Visible = True
            GroupBoxPartCreation.Width = 1038
            GroupBoxPartCreation.Height = 116

            txtPartnum.Text = txtpart.Text
            txtpartDescription.Text = txtPartDesc.Text

        End If
    End Sub

    Private Sub LblClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LblClose.Click
        GroupBoxPartCreation.Visible = False
    End Sub

    Private Sub Label63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label63.Click

    End Sub

    Private Sub RBCustomerExisting_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBCustomerExisting.CheckedChanged
        diablecustomerdata()

    End Sub

    Private Sub RBCustomerNew_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBCustomerNew.CheckedChanged
        enablecustomerdata()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxBuyer.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxPlanner_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxPlanner.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxItemType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxItemType.KeyDown




    End Sub

    Private Sub ComboBoxItemType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxItemType.SelectedIndexChanged

    End Sub

    Private Sub txtchilditemDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtchilditemDesc.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtchilditemDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtchilditemDesc.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtchilditemDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtchilditemDesc.TextChanged

    End Sub

    Private Sub txtsp1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtsp1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtsp1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsp1.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtsp1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtsp1.TextChanged

    End Sub

    Private Sub txtsp2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsp2.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)


        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtrun_Invalidated(ByVal sender As Object, ByVal e As System.Windows.Forms.InvalidateEventArgs) Handles txtrun.Invalidated

    End Sub

    Private Sub txtrun_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtrun.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtrun_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtrun.TextChanged

    End Sub

    Private Sub txtfix_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtfix.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtfix_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtfix.TextChanged

    End Sub

    Private Sub txtinsp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtinsp.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtinsp_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtinsp.TextChanged

    End Sub
    Sub fillpartdetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "Select Item_Type,Planner,Buyer,Lead_Run,Lead_Fix,Lead_Insp,Child_desc,prod_Line,Inv_Ac,Sp_note1,sp_note2,Item_Created_Date,Item_Created_By  FROM ENQ_Details where Enq_Int_code = " & txtenqintcode.Text & " and Enq_Detail_code =" & txtdetailintcode.Text & ""
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then

            Else
                ComboBoxItemType.Text = Trim((drSQL1.Item(0)))
                ComboBoxPlanner.Text = Trim((drSQL1.Item(1)))
                ComboBoxBuyer.Text = Trim((drSQL1.Item(2)))
                txtrun.Text = (drSQL1.Item(3))
                txtfix.Text = (drSQL1.Item(4))
                txtinsp.Text = (drSQL1.Item(5))
                txtchilditemDesc.Text = (drSQL1.Item(6))
                ComboBoxprodline.Text = Trim((drSQL1.Item(7)))
                ComboBoxInvAc.Text = Trim((drSQL1.Item(8)))
                txtsp1.Text = (drSQL1.Item(9))
                txtsp2.Text = (drSQL1.Item(10))
            End If

        End If

    End Sub

    Sub clearpartdesc()

        ComboBoxItemType.Text = ""
        ComboBoxPlanner.Text = ""
        ComboBoxBuyer.Text = ""
        txtrun.Text = ""
        txtfix.Text = ""
        txtinsp.Text = ""
        txtchilditemDesc.Text = ""
        ComboBoxprodline.Text = ""
        ComboBoxInvAc.Text = ""
        txtsp1.Text = ""
        txtsp2.Text = ""
    End Sub


    Private Sub ComboBoxInvAc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        ' If ComboBoxprodline.Text = "Warehouse" Then
        'ComboBoxInvAc.Text = "02000-60100"
        'ElseIf ComboBoxprodline.Text = "Factory" Then
        'ComboBoxInvAc.Text = "01000-60100"
        'End If
    End Sub

    Private Sub ComboBoxInvAc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBoxprodline_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxprodline.LostFocus
        If ComboBoxprodline.Text = "Warehouse" Then
            ComboBoxInvAc.Text = Trim("02000-60100")
        ElseIf ComboBoxprodline.Text = "Factory" Then
            ComboBoxInvAc.Text = "01000-60100"
        End If
    End Sub

    Private Sub ComboBoxprodline_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxprodline.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxCurrency_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxCurrency.SelectedIndexChanged

    End Sub

    Private Sub btnProjectDetClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt.Click
        PanelProjectDetails.Visible = False

    End Sub

    Private Sub txtMob_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMob1.TextChanged

    End Sub

    Private Sub txtph1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtph1.TextChanged

    End Sub

    Private Sub lblMode1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMode1.Click

    End Sub
    Private Sub mailcustcreation()

        Dim cnn As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim REG As String

        REG = txtRegNo.Text


        ' Create an Outlook application.
        Dim oApp As outlook._Application
        oApp = New outlook.Application()

        ' Create a new MailItem.
        Dim oMsg As outlook._MailItem
        oMsg = oApp.CreateItem(outlook.OlItemType.olMailItem)
        oMsg.Subject = "FOCUS SOFTWARE :  AUTOMATED MAIL. Reg No." & REG & " Customer Creation Required"


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        '        Dim cmSQL1 As SqlCommand
        '       Dim drSQL1 As SqlDataReader
        '      Dim SQL As String
        Dim t As String
        Dim cc As String
        ' Dim name As String


        'SQL = "SELECT  slno,[EMAILID],[STATUS] ,[TYPE]  FROM [FSPrograms].[dbo].[ENQ_Email_list]  WHERE TYPE = 'CUST' AND STATUS = 'A' order by slno "
        'cnSQL1.Open()
        'cmSQL1 = New SqlCommand(sql, cnSQL1)
        'drSQL1 = cmSQL1.ExecuteReader()

        'Dim n As Integer
        ''Dim i As Integer
        'n = 0


        'If drSQL1.Read() Then

        '    ' Dim counter As Integer = 0
        '    ' Do While drSQL1.Read()
        '    While drSQL1.Read

        '        If n = 0 Then
        '            t = drSQL1.Item(1)
        '        Else


        '            cc = drSQL1.Item(1)

        '            ' cc = cc & ";" & drSQL1.Item(1)
        '        End If

        '        n = n + 1
        '     End While

        t = "Rajesh.Ramdas@trelleborg.com"
        cc = "Vijay.Kumar@trelleborg.com;indira.shetty@trelleborg.com;Libin.George@trelleborg.com"

        Dim name As String = t
        name = name.Substring(0, name.Length - 15)


        oMsg.Body = "Dear " & name & "," & vbCrLf & vbCrLf & "New customer creation required. Customer name : '" & txtCustomer.Text & "' " & vbCr & "For details please refer Focus Software." & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Customer Support Team "


        oMsg.To = t
        oMsg.CC = cc

        '  End If


        Dim sBodyLen As String = oMsg.Body.Length
        '        Dim oAttachs As outlook.Attachments = oMsg.Attachments
        '       Dim oAttach As outlook.Attachment
        '      oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)

        ' Send
        oMsg.Send()

        ' Clean up
        oApp = Nothing
        oMsg = Nothing
        '        oAttach = Nothing
        '       oAttachs = Nothing

    End Sub

    Private Sub mailpartcreation()

        ' Dim cnn As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim REG As String

        REG = txtRegNo.Text


        ' Create an Outlook application.
        Dim oApp As outlook._Application
        oApp = New outlook.Application()

        ' Create a new MailItem.
        Dim oMsg As outlook._MailItem
        oMsg = oApp.CreateItem(outlook.OlItemType.olMailItem)
        oMsg.Subject = "FOCUS SOFTWARE :  AUTOMATED MAIL. Reg No." & REG & "Partnumber Creation Required"


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim SQL As String
        Dim t As String
        Dim cc As String
        ' Dim name As String


        SQL = "SELECT PartNumber FROM ENQ_Details WHERE  (Req ='Both' OR  Req = 'Part') AND Enq_Int_code =  " & txtenqintcode.Text & " "
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(SQL, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        'Dim n As Integer
        ''Dim i As Integer
        'n = 0


        If drSQL1.Read() Then

            'Rupak Ranjan; Pradeep Manore

            t = "Rupak.Ranjan@trelleborg.com"
            cc = "Pradeep.Manore@trelleborg.com;Libin.George@trelleborg.com;indira.shetty@trelleborg.com"

            Dim name As String = t
            name = name.Substring(0, name.Length - 15)


            oMsg.Body = "Dear " & name & "," & vbCrLf & vbCrLf & "New Partnumber creation required. Customer name : '" & txtCustomer.Text & "' " & vbCr & "For details please refer Focus Software." & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Customer Support Team "


            oMsg.To = t
            oMsg.CC = cc

            '  End If


            Dim sBodyLen As String = oMsg.Body.Length
            '        Dim oAttachs As outlook.Attachments = oMsg.Attachments
            '       Dim oAttach As outlook.Attachment
            '      oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)

            ' Send
            oMsg.Send()

            ' Clean up
            oApp = Nothing
            oMsg = Nothing
            '        oAttach = Nothing
            '       oAttachs = Nothing

        End If


    End Sub

    Private Sub mailEnquiryReg()
        'Dim cnn As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim sql As String


        ' Create an Outlook application.
        Dim oApp As outlook._Application
        oApp = New outlook.Application()

        ' Create a new MailItem.
        Dim oMsg As outlook._MailItem
        oMsg = oApp.CreateItem(outlook.OlItemType.olMailItem)
        oMsg.Subject = "FOCUS SOFTWARE :  AUTOMATED MAIL. New Enquiry  Registered. Reg No.'" & txtRegNo.Text & "' "

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        'Dim strSQL1 As String
        Dim t As String
        Dim cc As String
        ' Dim name As String

        If RBCustomerNew.Checked = False Then

            sql = "Select ISRMAILID, CSRMAILID, SEGHEADMAILID FROM TSS_CUSTOMERID_ISR where CustomerID = '" & txtCustID.Text & "' "

        Else
            sql = "SELECT  ISRMAILID, CSRMAILID, SEGHEADMAILID FROM  TSS_CUSTOMER_ISR_NEWCUSTOMER WHERE RegNo  = '" & txtRegNo.Text & "' "
        End If

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(sql, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then

            Else

                t = drSQL1.Item(1)
                cc = drSQL1(0)
                cc = cc & ";" & drSQL1(2) & ";" & "indira.shetty@trelleborg.com;Libin.George@trelleborg.com;Gopikrishna.CH@trelleborg.com"

                ' t = "indira.shetty@trelleborg.com"
                'cc = "indira.shetty@trelleborg.com"


                If ComboBoxEnquiryType.Text = "Internal RFQ" Then
                    cc = cc & ";" & "Ajith.Kumar@trelleborg.com"
                End If


                Dim name As String = t
                name = name.Substring(0, name.Length - 15)

                oMsg.Body = "Dear " & name & "," & vbCrLf & vbCrLf & "Enquiry of customer '" & txtCustomer.Text & "' is registered in Focus Software. " & vbCr & " Registration Number as above. " & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Customer Support Team "


                oMsg.To = t
                oMsg.CC = cc


            End If

        End If


        Dim sBodyLen As String = oMsg.Body.Length

        oMsg.Send()

        ' Clean up
        oApp = Nothing
        oMsg = Nothing


    End Sub




    Private Sub btnsave_Validating(sender As Object, e As CancelEventArgs) Handles btnsave.Validating

    End Sub


    Private Sub checkpriceavailability()
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String
        Dim stockDC As DataSet = New DataSet

        '' strSQL1 = "select Enq_Reg_NO from ENQ_RFQ_PriceDetails where Enq_Reg_NO = " & txtRegNo.Text & " AND Enq_Detail_code = " & txtdetailintcode.Text & " AND Status <> 'Released to Customer Sup'"
        'strSQL1 = "SELECT [RegNo],[Reg.Date],[CustomerID],[CustomerName],[CSR],[ISR],[Date_Add] as PriceGivenDt  FROM [FSPrograms].[dbo].[TSS_Enquiry_Price_Completed_QtyPrice] where FinalPrice > 0 and Type = 'MTS' And days <=365"


        '    strSQL1 = " SELECT [RegNo],[Reg.Date],[CustomerID],[CustomerName],[CSR],[ISR],[Date_Add] as PriceGivenDt, FinalPrice  FROM [FSPrograms].[dbo].[TSS_Enquiry_Price_Completed_QtyPrice] where FinalPrice > 0 and Type = 'MTS' And days <=365 " & _
        '             "union select 0,Effective_Date,'Mfg Unit','Notional_Price','','',Effective_Date, Notional_Price  from  [FSDBBR].[dbo].TSS_Notional_price with (nolock)"


        strSQL1 = "SELECT [RegNo],[Reg.Date],[CustomerID],[CustomerName],[CSR],[ISR],[Date_Add] as PriceGivenDt,Qty,Qty_Type, FinalPrice   FROM [FSPrograms].[dbo].[TSS_Enquiry_Price_Completed_QtyPrice] where FinalPrice > 0 And days <=365 and PartNumber = '" & txtpart.Text & "' " & _
                 "union select 0,Effective_Date,'Mfg Unit','Notional_Price','','',Effective_Date,1,'', Notional_Price  from  [FSDBBR].[dbo].TSS_Notional_price with (nolock) where Part_no = '" & txtpart.Text & "'"



        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then

            Else
                MsgBox(" Price given by purchase is valid for 1 year. See the details...   ", vbInformation) '90 days to 1 year changed on 31-march 2020-after meeting -MOM from Neelapandiyan, SCM

                GroupBoxPriceAvble.Visible = True

                GroupBoxPriceAvble.Enabled = True
                GroupBoxPriceAvble.Location = New System.Drawing.Point(504, 72)
                GroupBoxPriceAvble.Width = 706
                GroupBoxPriceAvble.Height = 256


                datagridPriceAvble.Visible = True
                datagridPriceAvble.Enabled = True
                datagridPriceAvble.Location = New System.Drawing.Point(6, 33)
                datagridPriceAvble.Height = 210
                datagridPriceAvble.Width = 684


                Dim sqlCmd As SqlCommand = New SqlCommand(strSQL1, cnSQL1)
                Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
                cnSQL1.Close()

                stockDAC.SelectCommand = sqlCmd
                cnSQL1.Open()

                stockDAC.TableMappings.Add("Table", "Enq")
                'get data
                stockDAC.Fill(stockDC)

                datagridPriceAvble.DataSource = stockDC.Tables(0)
                Exit Sub
            End If

        End If

    End Sub

    Private Sub LabelClose_Click(sender As Object, e As EventArgs) Handles LabelClose.Click
        GroupBoxPriceAvble.Visible = False
    End Sub

    Private Sub RFQHistory_Load1()

        datagridPriceAvble.Enabled = True
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim strSQL As String


        GroupBoxPriceAvble.Visible = True
        datagridPriceAvble.Visible = True


        Dim stockDC As DataSet = New DataSet

        strSQL = "SELECT     RegNo, [Reg.Date], CustomerID, CustomerName, CSR, SlNo, Part_Source, PartNumber, PartDescription, Class1, Qty, Price, Qty_Type, Factor, FinalPrice, Source_Mtrl, " & _
            " MOQ, SPU, LeadTime, Type, Stock_Avble, Vendor_Ref, Name, Vendor_Quote, Special_Remarks FROM  TSS_Enquiry_Price_Completed_QtyPrice where PartNumber = '" & txtpart.Text & "' order by RegNo,SlNo"


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        datagridPriceAvble.DataSource = stockDC.Tables(0)

    End Sub



    Private Sub DTPRegDt_ValueChanged(sender As Object, e As EventArgs) Handles DTPRegDt.ValueChanged

    End Sub

    Private Sub txtcustintcode_TextChanged(sender As Object, e As EventArgs) Handles txtcustintcode.TextChanged

    End Sub

    Private Sub datagridPriceAvble_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridPriceAvble.CellContentClick

    End Sub

    Private Sub btnItemHistory_Click(sender As Object, e As EventArgs) Handles btnItemHistory.Click
        If Len(txtpart.Text) > 3 Then
            checkpriceavailability()
        End If

    End Sub
End Class
