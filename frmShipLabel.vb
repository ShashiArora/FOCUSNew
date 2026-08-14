Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports System
Imports System.Collections
Imports System.Data
Imports System.Math
Imports Microsoft.Reporting.WinForms

'<System.Windows.Forms jitdebugging="true" />
'Imports System
Imports System.Collections.Generic
Imports System.Text

Imports System.Configuration
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.IO


Public Class frmMain
    Inherits System.Windows.Forms.Form


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
    Friend WithEvents TxtCustItem As System.Windows.Forms.TextBox
    Friend WithEvents LblCustItem As System.Windows.Forms.Label
    Friend WithEvents LblCustDesc As System.Windows.Forms.Label
    Friend WithEvents TxtCustDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustPO As System.Windows.Forms.TextBox
    Friend WithEvents LblCustPO As System.Windows.Forms.Label
    Friend WithEvents LblNoCust As System.Windows.Forms.Label
    Friend WithEvents LblLabelType As System.Windows.Forms.Label
    Protected WithEvents TxtSelectReport As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtCOLnNo As System.Windows.Forms.TextBox
    Friend WithEvents LblCOLnNo As System.Windows.Forms.Label
    Friend WithEvents TxtUoM As System.Windows.Forms.TextBox
    Friend WithEvents BtnPrtMulti As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents LblInvoice As System.Windows.Forms.Label
    Friend WithEvents txtInvoiceNo As System.Windows.Forms.TextBox
    Friend WithEvents txtItem As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescription As System.Windows.Forms.TextBox
    Friend WithEvents LblLineNo As System.Windows.Forms.Label
    Friend WithEvents TxtLineNo As System.Windows.Forms.TextBox
    Friend WithEvents LblDescription As System.Windows.Forms.Label
    Friend WithEvents txtQty As System.Windows.Forms.TextBox
    Friend WithEvents lblLotNo As System.Windows.Forms.Label
    Friend WithEvents txtLotNo As System.Windows.Forms.TextBox
    Friend WithEvents lblQty As System.Windows.Forms.Label
    Friend WithEvents lblItem As System.Windows.Forms.Label
    Friend WithEvents LblCONo As System.Windows.Forms.Label
    Friend WithEvents TxtCustOrdNo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LblOrdQty As System.Windows.Forms.Label
    Friend WithEvents TxtLblQty As System.Windows.Forms.TextBox
    Friend WithEvents LblQtyLbl As System.Windows.Forms.Label
    Friend WithEvents LblQtyonLbl As System.Windows.Forms.Label
    Friend WithEvents txtCustomer As System.Windows.Forms.TextBox
    Friend WithEvents ListViewPart As System.Windows.Forms.ListView
    Friend WithEvents lblUOM As System.Windows.Forms.Label
    Friend WithEvents lblcustomer As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CheckFS As System.Windows.Forms.CheckBox
    Friend WithEvents RDBManual As System.Windows.Forms.RadioButton
    Friend WithEvents RDBAutomatic As System.Windows.Forms.RadioButton
    Friend WithEvents RdbShip As System.Windows.Forms.RadioButton
    Friend WithEvents RdbPORV As System.Windows.Forms.RadioButton
    Friend WithEvents rdbDCChecking As System.Windows.Forms.RadioButton
    Friend WithEvents CmbLotNo As System.Windows.Forms.ComboBox
    Friend WithEvents Groupporv As System.Windows.Forms.GroupBox
    Friend WithEvents lblPORV As System.Windows.Forms.Label
    Friend WithEvents lblQtyPerLabel As System.Windows.Forms.Label
    Friend WithEvents lblNoofLabel As System.Windows.Forms.Label
    Friend WithEvents txtNoofLabels As System.Windows.Forms.TextBox
    Friend WithEvents LblNoLot As System.Windows.Forms.Label
    Friend WithEvents lblQtyOnLabel As System.Windows.Forms.TextBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents datagridStock As System.Windows.Forms.DataGrid
    Friend WithEvents txtPart As System.Windows.Forms.TextBox
    Friend WithEvents txtLot As System.Windows.Forms.TextBox
    Friend WithEvents TXTKEY As System.Windows.Forms.TextBox
    Friend WithEvents TextlblqtyPORV As System.Windows.Forms.TextBox
    Friend WithEvents TXTLOTQTYPORV As System.Windows.Forms.TextBox
    Friend WithEvents btnPorvLblPrint As System.Windows.Forms.Button
    Friend WithEvents btnPorvPrintCancel As System.Windows.Forms.Button
    Friend WithEvents txtBin1 As System.Windows.Forms.TextBox
    Friend WithEvents txtItemDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtLotDate As System.Windows.Forms.TextBox
    Friend WithEvents txtStkRoom1 As System.Windows.Forms.TextBox
    Friend WithEvents TXTVENDCURE As System.Windows.Forms.TextBox
    Friend WithEvents txtconumber As System.Windows.Forms.TextBox
    Friend WithEvents txtcustid As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents BtnShipOK As System.Windows.Forms.Button
    Friend WithEvents DataGridShip As System.Windows.Forms.DataGrid
    Friend WithEvents txtshipqty As System.Windows.Forms.TextBox
    Friend WithEvents txtshiplot As System.Windows.Forms.TextBox
    Friend WithEvents txtshiplabel As System.Windows.Forms.TextBox
    Friend WithEvents txtshipqpl As System.Windows.Forms.TextBox
    Friend WithEvents txtshipLine As System.Windows.Forms.TextBox
    Friend WithEvents txtshippartno As System.Windows.Forms.TextBox
    Friend WithEvents btnshipprint As System.Windows.Forms.Button
    Friend WithEvents groupShip As System.Windows.Forms.GroupBox
    Friend WithEvents LblInfo As System.Windows.Forms.Label
    Friend WithEvents TxtInfo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnclear As System.Windows.Forms.Button
    Friend WithEvents btnporvok As System.Windows.Forms.Button
    Friend WithEvents txtbuyer As System.Windows.Forms.TextBox
    Friend WithEvents txtdesig As System.Windows.Forms.TextBox
    Friend WithEvents txtdept As System.Windows.Forms.TextBox
    Friend WithEvents txtphone As System.Windows.Forms.TextBox
    Friend WithEvents txtmobile As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents txtshipinvdate As System.Windows.Forms.TextBox
    Friend WithEvents RBPORVGoel As System.Windows.Forms.RadioButton
    Friend WithEvents RBDateWise As System.Windows.Forms.RadioButton
    Friend WithEvents RBPOwise As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFROM As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPONo As System.Windows.Forms.Label
    Friend WithEvents lblfrom As System.Windows.Forms.Label
    Friend WithEvents lblto As System.Windows.Forms.Label
    Friend WithEvents TXTPartSelect As System.Windows.Forms.TextBox
    Friend WithEvents txtPonumber As System.Windows.Forms.TextBox
    Friend WithEvents txtdate1 As System.Windows.Forms.TextBox
    Friend WithEvents txtdate2 As System.Windows.Forms.TextBox
    Friend WithEvents txtdate3 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents RBPorvIMTRGoel As System.Windows.Forms.RadioButton
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtcusto As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.TxtCustItem = New System.Windows.Forms.TextBox()
        Me.LblCustItem = New System.Windows.Forms.Label()
        Me.LblCustDesc = New System.Windows.Forms.Label()
        Me.TxtCustDesc = New System.Windows.Forms.TextBox()
        Me.TxtCustPO = New System.Windows.Forms.TextBox()
        Me.LblCustPO = New System.Windows.Forms.Label()
        Me.LblNoCust = New System.Windows.Forms.Label()
        Me.TxtSelectReport = New System.Windows.Forms.ComboBox()
        Me.LblLabelType = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RBPorvIMTRGoel = New System.Windows.Forms.RadioButton()
        Me.RBPORVGoel = New System.Windows.Forms.RadioButton()
        Me.rdbDCChecking = New System.Windows.Forms.RadioButton()
        Me.RdbPORV = New System.Windows.Forms.RadioButton()
        Me.RdbShip = New System.Windows.Forms.RadioButton()
        Me.CheckFS = New System.Windows.Forms.CheckBox()
        Me.RDBManual = New System.Windows.Forms.RadioButton()
        Me.RDBAutomatic = New System.Windows.Forms.RadioButton()
        Me.TxtCOLnNo = New System.Windows.Forms.TextBox()
        Me.LblCOLnNo = New System.Windows.Forms.Label()
        Me.TxtUoM = New System.Windows.Forms.TextBox()
        Me.lblUOM = New System.Windows.Forms.Label()
        Me.BtnPrtMulti = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.LblInvoice = New System.Windows.Forms.Label()
        Me.txtInvoiceNo = New System.Windows.Forms.TextBox()
        Me.txtItem = New System.Windows.Forms.TextBox()
        Me.TxtDescription = New System.Windows.Forms.TextBox()
        Me.LblLineNo = New System.Windows.Forms.Label()
        Me.TxtLineNo = New System.Windows.Forms.TextBox()
        Me.LblDescription = New System.Windows.Forms.Label()
        Me.txtQty = New System.Windows.Forms.TextBox()
        Me.lblLotNo = New System.Windows.Forms.Label()
        Me.txtLotNo = New System.Windows.Forms.TextBox()
        Me.lblQty = New System.Windows.Forms.Label()
        Me.lblItem = New System.Windows.Forms.Label()
        Me.LblCONo = New System.Windows.Forms.Label()
        Me.TxtCustOrdNo = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblQtyOnLabel = New System.Windows.Forms.TextBox()
        Me.CmbLotNo = New System.Windows.Forms.ComboBox()
        Me.ListViewPart = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.txtCustomer = New System.Windows.Forms.TextBox()
        Me.lblcustomer = New System.Windows.Forms.Label()
        Me.LblQtyonLbl = New System.Windows.Forms.Label()
        Me.LblNoLot = New System.Windows.Forms.Label()
        Me.LblQtyLbl = New System.Windows.Forms.Label()
        Me.TxtLblQty = New System.Windows.Forms.TextBox()
        Me.LblOrdQty = New System.Windows.Forms.Label()
        Me.Groupporv = New System.Windows.Forms.GroupBox()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblPONo = New System.Windows.Forms.Label()
        Me.lblfrom = New System.Windows.Forms.Label()
        Me.lblto = New System.Windows.Forms.Label()
        Me.TXTPartSelect = New System.Windows.Forms.TextBox()
        Me.txtPonumber = New System.Windows.Forms.TextBox()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFROM = New System.Windows.Forms.DateTimePicker()
        Me.RBPOwise = New System.Windows.Forms.RadioButton()
        Me.RBDateWise = New System.Windows.Forms.RadioButton()
        Me.btnporvok = New System.Windows.Forms.Button()
        Me.TXTVENDCURE = New System.Windows.Forms.TextBox()
        Me.txtStkRoom1 = New System.Windows.Forms.TextBox()
        Me.txtLotDate = New System.Windows.Forms.TextBox()
        Me.txtItemDesc = New System.Windows.Forms.TextBox()
        Me.txtBin1 = New System.Windows.Forms.TextBox()
        Me.btnPorvPrintCancel = New System.Windows.Forms.Button()
        Me.btnPorvLblPrint = New System.Windows.Forms.Button()
        Me.TXTLOTQTYPORV = New System.Windows.Forms.TextBox()
        Me.TXTKEY = New System.Windows.Forms.TextBox()
        Me.txtLot = New System.Windows.Forms.TextBox()
        Me.txtPart = New System.Windows.Forms.TextBox()
        Me.datagridStock = New System.Windows.Forms.DataGrid()
        Me.txtNoofLabels = New System.Windows.Forms.TextBox()
        Me.lblNoofLabel = New System.Windows.Forms.Label()
        Me.lblQtyPerLabel = New System.Windows.Forms.Label()
        Me.TextlblqtyPORV = New System.Windows.Forms.TextBox()
        Me.lblPORV = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.groupShip = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtshipinvdate = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.txtcusto = New System.Windows.Forms.TextBox()
        Me.txtmobile = New System.Windows.Forms.TextBox()
        Me.txtphone = New System.Windows.Forms.TextBox()
        Me.txtdept = New System.Windows.Forms.TextBox()
        Me.txtdesig = New System.Windows.Forms.TextBox()
        Me.txtbuyer = New System.Windows.Forms.TextBox()
        Me.btnclear = New System.Windows.Forms.Button()
        Me.LblInfo = New System.Windows.Forms.Label()
        Me.TxtInfo = New System.Windows.Forms.TextBox()
        Me.BtnShipOK = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.txtconumber = New System.Windows.Forms.TextBox()
        Me.txtcustid = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnshipprint = New System.Windows.Forms.Button()
        Me.txtshipqty = New System.Windows.Forms.TextBox()
        Me.txtshipLine = New System.Windows.Forms.TextBox()
        Me.txtshiplot = New System.Windows.Forms.TextBox()
        Me.txtshippartno = New System.Windows.Forms.TextBox()
        Me.DataGridShip = New System.Windows.Forms.DataGrid()
        Me.txtshiplabel = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtshipqpl = New System.Windows.Forms.TextBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtdate1 = New System.Windows.Forms.TextBox()
        Me.txtdate2 = New System.Windows.Forms.TextBox()
        Me.txtdate3 = New System.Windows.Forms.TextBox()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Groupporv.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.datagridStock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.groupShip.SuspendLayout()
        CType(Me.DataGridShip, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TxtCustItem
        '
        Me.TxtCustItem.AllowDrop = True
        Me.TxtCustItem.Enabled = False
        Me.TxtCustItem.Location = New System.Drawing.Point(452, 212)
        Me.TxtCustItem.Name = "TxtCustItem"
        Me.TxtCustItem.Size = New System.Drawing.Size(210, 22)
        Me.TxtCustItem.TabIndex = 9
        Me.TxtCustItem.TabStop = False
        '
        'LblCustItem
        '
        Me.LblCustItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LblCustItem.Location = New System.Drawing.Point(970, 110)
        Me.LblCustItem.Name = "LblCustItem"
        Me.LblCustItem.Size = New System.Drawing.Size(38, 26)
        Me.LblCustItem.TabIndex = 24
        Me.LblCustItem.Text = "Cust Item No"
        Me.LblCustItem.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LblCustItem.Visible = False
        '
        'LblCustDesc
        '
        Me.LblCustDesc.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LblCustDesc.Location = New System.Drawing.Point(872, 115)
        Me.LblCustDesc.Name = "LblCustDesc"
        Me.LblCustDesc.Size = New System.Drawing.Size(30, 19)
        Me.LblCustDesc.TabIndex = 27
        Me.LblCustDesc.Text = "Cust Desc."
        Me.LblCustDesc.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LblCustDesc.Visible = False
        '
        'TxtCustDesc
        '
        Me.TxtCustDesc.Enabled = False
        Me.TxtCustDesc.Location = New System.Drawing.Point(1052, 105)
        Me.TxtCustDesc.Name = "TxtCustDesc"
        Me.TxtCustDesc.Size = New System.Drawing.Size(28, 22)
        Me.TxtCustDesc.TabIndex = 11
        Me.TxtCustDesc.TabStop = False
        Me.TxtCustDesc.Visible = False
        '
        'TxtCustPO
        '
        Me.TxtCustPO.AllowDrop = True
        Me.TxtCustPO.Enabled = False
        Me.TxtCustPO.Location = New System.Drawing.Point(1016, 110)
        Me.TxtCustPO.Name = "TxtCustPO"
        Me.TxtCustPO.Size = New System.Drawing.Size(28, 22)
        Me.TxtCustPO.TabIndex = 7
        Me.TxtCustPO.TabStop = False
        Me.TxtCustPO.Visible = False
        '
        'LblCustPO
        '
        Me.LblCustPO.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LblCustPO.Cursor = System.Windows.Forms.Cursors.Cross
        Me.LblCustPO.Location = New System.Drawing.Point(924, 114)
        Me.LblCustPO.Name = "LblCustPO"
        Me.LblCustPO.Size = New System.Drawing.Size(38, 27)
        Me.LblCustPO.TabIndex = 28
        Me.LblCustPO.Text = "Cust PO No"
        Me.LblCustPO.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LblCustPO.Visible = False
        '
        'LblNoCust
        '
        Me.LblNoCust.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblNoCust.ForeColor = System.Drawing.Color.Red
        Me.LblNoCust.Location = New System.Drawing.Point(634, 415)
        Me.LblNoCust.Name = "LblNoCust"
        Me.LblNoCust.Size = New System.Drawing.Size(28, 27)
        Me.LblNoCust.TabIndex = 34
        '
        'TxtSelectReport
        '
        Me.TxtSelectReport.Items.AddRange(New Object() {"A)Item Labels Pre-Invoice", "B)Item Labels Post-Invoice", "C)Box Label", "D)KIT Child Items List"})
        Me.TxtSelectReport.Location = New System.Drawing.Point(108, 27)
        Me.TxtSelectReport.Name = "TxtSelectReport"
        Me.TxtSelectReport.Size = New System.Drawing.Size(268, 24)
        Me.TxtSelectReport.TabIndex = 0
        '
        'LblLabelType
        '
        Me.LblLabelType.BackColor = System.Drawing.Color.Transparent
        Me.LblLabelType.Location = New System.Drawing.Point(10, 37)
        Me.LblLabelType.Name = "LblLabelType"
        Me.LblLabelType.Size = New System.Drawing.Size(76, 18)
        Me.LblLabelType.TabIndex = 37
        Me.LblLabelType.Text = "Label Type"
        Me.LblLabelType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.RBPorvIMTRGoel)
        Me.GroupBox2.Controls.Add(Me.RBPORVGoel)
        Me.GroupBox2.Controls.Add(Me.rdbDCChecking)
        Me.GroupBox2.Controls.Add(Me.RdbPORV)
        Me.GroupBox2.Controls.Add(Me.RdbShip)
        Me.GroupBox2.Location = New System.Drawing.Point(10, 102)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(682, 69)
        Me.GroupBox2.TabIndex = 38
        Me.GroupBox2.TabStop = False
        '
        'RBPorvIMTRGoel
        '
        Me.RBPorvIMTRGoel.Location = New System.Drawing.Point(370, 22)
        Me.RBPorvIMTRGoel.Name = "RBPorvIMTRGoel"
        Me.RBPorvIMTRGoel.Size = New System.Drawing.Size(172, 17)
        Me.RBPorvIMTRGoel.TabIndex = 7
        Me.RBPorvIMTRGoel.Text = "PORV-GOEL-IMTR"
        '
        'RBPORVGoel
        '
        Me.RBPORVGoel.Location = New System.Drawing.Point(236, 21)
        Me.RBPORVGoel.Name = "RBPORVGoel"
        Me.RBPORVGoel.Size = New System.Drawing.Size(126, 27)
        Me.RBPORVGoel.TabIndex = 6
        Me.RBPORVGoel.Text = "PORV - GOEL"
        '
        'rdbDCChecking
        '
        Me.rdbDCChecking.Location = New System.Drawing.Point(20, 43)
        Me.rdbDCChecking.Name = "rdbDCChecking"
        Me.rdbDCChecking.Size = New System.Drawing.Size(134, 19)
        Me.rdbDCChecking.TabIndex = 5
        Me.rdbDCChecking.Text = "DC Validation"
        Me.rdbDCChecking.Visible = False
        '
        'RdbPORV
        '
        Me.RdbPORV.Location = New System.Drawing.Point(142, 20)
        Me.RdbPORV.Name = "RdbPORV"
        Me.RdbPORV.Size = New System.Drawing.Size(96, 18)
        Me.RdbPORV.TabIndex = 4
        Me.RdbPORV.Text = "PORV"
        '
        'RdbShip
        '
        Me.RdbShip.Location = New System.Drawing.Point(20, 18)
        Me.RdbShip.Name = "RdbShip"
        Me.RdbShip.Size = New System.Drawing.Size(114, 19)
        Me.RdbShip.TabIndex = 3
        Me.RdbShip.Text = "SHIP"
        '
        'CheckFS
        '
        Me.CheckFS.Checked = True
        Me.CheckFS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckFS.Location = New System.Drawing.Point(586, 37)
        Me.CheckFS.Name = "CheckFS"
        Me.CheckFS.Size = New System.Drawing.Size(86, 18)
        Me.CheckFS.TabIndex = 2
        Me.CheckFS.Text = "Std. FS"
        '
        'RDBManual
        '
        Me.RDBManual.Location = New System.Drawing.Point(490, 37)
        Me.RDBManual.Name = "RDBManual"
        Me.RDBManual.Size = New System.Drawing.Size(86, 18)
        Me.RDBManual.TabIndex = 1
        Me.RDBManual.Text = "Manual"
        '
        'RDBAutomatic
        '
        Me.RDBAutomatic.Location = New System.Drawing.Point(384, 37)
        Me.RDBAutomatic.Name = "RDBAutomatic"
        Me.RDBAutomatic.Size = New System.Drawing.Size(96, 18)
        Me.RDBAutomatic.TabIndex = 0
        Me.RDBAutomatic.Text = "Automatic"
        '
        'TxtCOLnNo
        '
        Me.TxtCOLnNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCOLnNo.Enabled = False
        Me.TxtCOLnNo.Location = New System.Drawing.Point(452, 138)
        Me.TxtCOLnNo.MaxLength = 20
        Me.TxtCOLnNo.Name = "TxtCOLnNo"
        Me.TxtCOLnNo.Size = New System.Drawing.Size(104, 22)
        Me.TxtCOLnNo.TabIndex = 5
        Me.TxtCOLnNo.TabStop = False
        '
        'LblCOLnNo
        '
        Me.LblCOLnNo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LblCOLnNo.Location = New System.Drawing.Point(346, 138)
        Me.LblCOLnNo.Name = "LblCOLnNo"
        Me.LblCOLnNo.Size = New System.Drawing.Size(96, 19)
        Me.LblCOLnNo.TabIndex = 22
        Me.LblCOLnNo.Text = "CO Line No"
        Me.LblCOLnNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TxtUoM
        '
        Me.TxtUoM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUoM.Enabled = False
        Me.TxtUoM.Location = New System.Drawing.Point(452, 286)
        Me.TxtUoM.MaxLength = 20
        Me.TxtUoM.Name = "TxtUoM"
        Me.TxtUoM.Size = New System.Drawing.Size(104, 22)
        Me.TxtUoM.TabIndex = 13
        Me.TxtUoM.TabStop = False
        '
        'lblUOM
        '
        Me.lblUOM.Location = New System.Drawing.Point(346, 286)
        Me.lblUOM.Name = "lblUOM"
        Me.lblUOM.Size = New System.Drawing.Size(76, 27)
        Me.lblUOM.TabIndex = 20
        Me.lblUOM.Text = "UoM"
        Me.lblUOM.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BtnPrtMulti
        '
        Me.BtnPrtMulti.Enabled = False
        Me.BtnPrtMulti.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BtnPrtMulti.Location = New System.Drawing.Point(480, 406)
        Me.BtnPrtMulti.Name = "BtnPrtMulti"
        Me.BtnPrtMulti.Size = New System.Drawing.Size(154, 32)
        Me.BtnPrtMulti.TabIndex = 18
        Me.BtnPrtMulti.TabStop = False
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancel.Location = New System.Drawing.Point(404, 406)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 32)
        Me.btnCancel.TabIndex = 17
        Me.btnCancel.TabStop = False
        Me.btnCancel.Text = "Cancel"
        '
        'btnPrint
        '
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPrint.Location = New System.Drawing.Point(268, 406)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(78, 32)
        Me.btnPrint.TabIndex = 16
        Me.btnPrint.TabStop = False
        Me.btnPrint.Text = "Print Label"
        '
        'LblInvoice
        '
        Me.LblInvoice.Location = New System.Drawing.Point(286, 25)
        Me.LblInvoice.Name = "LblInvoice"
        Me.LblInvoice.Size = New System.Drawing.Size(76, 27)
        Me.LblInvoice.TabIndex = 0
        Me.LblInvoice.Text = "1"
        Me.LblInvoice.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtInvoiceNo
        '
        Me.txtInvoiceNo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtInvoiceNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtInvoiceNo.Location = New System.Drawing.Point(370, 27)
        Me.txtInvoiceNo.MaxLength = 20
        Me.txtInvoiceNo.Name = "txtInvoiceNo"
        Me.txtInvoiceNo.Size = New System.Drawing.Size(124, 22)
        Me.txtInvoiceNo.TabIndex = 1
        '
        'txtItem
        '
        Me.txtItem.Enabled = False
        Me.txtItem.Location = New System.Drawing.Point(508, 445)
        Me.txtItem.Name = "txtItem"
        Me.txtItem.Size = New System.Drawing.Size(232, 22)
        Me.txtItem.TabIndex = 6
        Me.txtItem.TabStop = False
        '
        'TxtDescription
        '
        Me.TxtDescription.Enabled = False
        Me.TxtDescription.Location = New System.Drawing.Point(106, 212)
        Me.TxtDescription.Name = "TxtDescription"
        Me.TxtDescription.Size = New System.Drawing.Size(230, 22)
        Me.TxtDescription.TabIndex = 8
        Me.TxtDescription.TabStop = False
        '
        'LblLineNo
        '
        Me.LblLineNo.Location = New System.Drawing.Point(496, 23)
        Me.LblLineNo.Name = "LblLineNo"
        Me.LblLineNo.Size = New System.Drawing.Size(86, 28)
        Me.LblLineNo.TabIndex = 21
        Me.LblLineNo.Text = "Ship Line No"
        Me.LblLineNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TxtLineNo
        '
        Me.TxtLineNo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.TxtLineNo.ForeColor = System.Drawing.Color.Black
        Me.TxtLineNo.Location = New System.Drawing.Point(590, 24)
        Me.TxtLineNo.Name = "TxtLineNo"
        Me.TxtLineNo.Size = New System.Drawing.Size(66, 22)
        Me.TxtLineNo.TabIndex = 2
        Me.TxtLineNo.Text = "%"
        '
        'LblDescription
        '
        Me.LblDescription.Location = New System.Drawing.Point(10, 212)
        Me.LblDescription.Name = "LblDescription"
        Me.LblDescription.Size = New System.Drawing.Size(86, 27)
        Me.LblDescription.TabIndex = 23
        Me.LblDescription.Text = "Description"
        Me.LblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtQty
        '
        Me.txtQty.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtQty.Enabled = False
        Me.txtQty.Location = New System.Drawing.Point(106, 249)
        Me.txtQty.MaxLength = 20
        Me.txtQty.Name = "txtQty"
        Me.txtQty.Size = New System.Drawing.Size(86, 22)
        Me.txtQty.TabIndex = 10
        Me.txtQty.TabStop = False
        '
        'lblLotNo
        '
        Me.lblLotNo.Location = New System.Drawing.Point(10, 286)
        Me.lblLotNo.Name = "lblLotNo"
        Me.lblLotNo.Size = New System.Drawing.Size(86, 27)
        Me.lblLotNo.TabIndex = 17
        Me.lblLotNo.Text = "Lot No"
        Me.lblLotNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtLotNo
        '
        Me.txtLotNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtLotNo.Enabled = False
        Me.txtLotNo.Location = New System.Drawing.Point(106, 286)
        Me.txtLotNo.MaxLength = 15
        Me.txtLotNo.Name = "txtLotNo"
        Me.txtLotNo.Size = New System.Drawing.Size(240, 22)
        Me.txtLotNo.TabIndex = 12
        Me.txtLotNo.TabStop = False
        '
        'lblQty
        '
        Me.lblQty.Location = New System.Drawing.Point(10, 249)
        Me.lblQty.Name = "lblQty"
        Me.lblQty.Size = New System.Drawing.Size(86, 27)
        Me.lblQty.TabIndex = 18
        Me.lblQty.Text = "Quantity"
        Me.lblQty.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblItem
        '
        Me.lblItem.Location = New System.Drawing.Point(10, 175)
        Me.lblItem.Name = "lblItem"
        Me.lblItem.Size = New System.Drawing.Size(86, 27)
        Me.lblItem.TabIndex = 3
        Me.lblItem.Text = "Order item"
        Me.lblItem.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LblCONo
        '
        Me.LblCONo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LblCONo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LblCONo.Location = New System.Drawing.Point(346, 102)
        Me.LblCONo.Name = "LblCONo"
        Me.LblCONo.Size = New System.Drawing.Size(96, 18)
        Me.LblCONo.TabIndex = 30
        Me.LblCONo.Text = "Cust Order No"
        Me.LblCONo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TxtCustOrdNo
        '
        Me.TxtCustOrdNo.AllowDrop = True
        Me.TxtCustOrdNo.Enabled = False
        Me.TxtCustOrdNo.Location = New System.Drawing.Point(528, 157)
        Me.TxtCustOrdNo.Name = "TxtCustOrdNo"
        Me.TxtCustOrdNo.Size = New System.Drawing.Size(212, 22)
        Me.TxtCustOrdNo.TabIndex = 4
        Me.TxtCustOrdNo.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.lblQtyOnLabel)
        Me.GroupBox1.Controls.Add(Me.CmbLotNo)
        Me.GroupBox1.Controls.Add(Me.ListViewPart)
        Me.GroupBox1.Controls.Add(Me.txtCustomer)
        Me.GroupBox1.Controls.Add(Me.lblcustomer)
        Me.GroupBox1.Controls.Add(Me.LblQtyonLbl)
        Me.GroupBox1.Controls.Add(Me.LblNoLot)
        Me.GroupBox1.Controls.Add(Me.LblQtyLbl)
        Me.GroupBox1.Controls.Add(Me.TxtLblQty)
        Me.GroupBox1.Controls.Add(Me.LblOrdQty)
        Me.GroupBox1.Controls.Add(Me.TxtCOLnNo)
        Me.GroupBox1.Controls.Add(Me.LblCOLnNo)
        Me.GroupBox1.Controls.Add(Me.TxtUoM)
        Me.GroupBox1.Controls.Add(Me.lblUOM)
        Me.GroupBox1.Controls.Add(Me.BtnPrtMulti)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnPrint)
        Me.GroupBox1.Controls.Add(Me.LblCONo)
        Me.GroupBox1.Controls.Add(Me.TxtCustOrdNo)
        Me.GroupBox1.Controls.Add(Me.TxtDescription)
        Me.GroupBox1.Controls.Add(Me.LblDescription)
        Me.GroupBox1.Controls.Add(Me.txtQty)
        Me.GroupBox1.Controls.Add(Me.lblLotNo)
        Me.GroupBox1.Controls.Add(Me.txtLotNo)
        Me.GroupBox1.Controls.Add(Me.lblQty)
        Me.GroupBox1.Controls.Add(Me.lblItem)
        Me.GroupBox1.Controls.Add(Me.LblNoCust)
        Me.GroupBox1.Controls.Add(Me.TxtCustItem)
        Me.GroupBox1.Controls.Add(Me.LblLabelType)
        Me.GroupBox1.Controls.Add(Me.TxtSelectReport)
        Me.GroupBox1.Controls.Add(Me.RDBAutomatic)
        Me.GroupBox1.Controls.Add(Me.RDBManual)
        Me.GroupBox1.Controls.Add(Me.CheckFS)
        Me.GroupBox1.Location = New System.Drawing.Point(1438, 120)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(60, 59)
        Me.GroupBox1.TabIndex = 33
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "SHIP LABELS"
        Me.GroupBox1.Visible = False
        '
        'lblQtyOnLabel
        '
        Me.lblQtyOnLabel.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.lblQtyOnLabel.Enabled = False
        Me.lblQtyOnLabel.Location = New System.Drawing.Point(106, 406)
        Me.lblQtyOnLabel.MaxLength = 20
        Me.lblQtyOnLabel.Name = "lblQtyOnLabel"
        Me.lblQtyOnLabel.Size = New System.Drawing.Size(106, 22)
        Me.lblQtyOnLabel.TabIndex = 42
        Me.lblQtyOnLabel.TabStop = False
        '
        'CmbLotNo
        '
        Me.CmbLotNo.ItemHeight = 16
        Me.CmbLotNo.Location = New System.Drawing.Point(106, 286)
        Me.CmbLotNo.Name = "CmbLotNo"
        Me.CmbLotNo.Size = New System.Drawing.Size(192, 24)
        Me.CmbLotNo.TabIndex = 41
        Me.CmbLotNo.Visible = False
        '
        'ListViewPart
        '
        Me.ListViewPart.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.ListViewPart.FullRowSelect = True
        Me.ListViewPart.GridLines = True
        Me.ListViewPart.Location = New System.Drawing.Point(548, 65)
        Me.ListViewPart.Name = "ListViewPart"
        Me.ListViewPart.Size = New System.Drawing.Size(114, 27)
        Me.ListViewPart.TabIndex = 40
        Me.ListViewPart.UseCompatibleStateImageBehavior = False
        Me.ListViewPart.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ITEM"
        Me.ColumnHeader1.Width = 200
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "UOM"
        '
        'txtCustomer
        '
        Me.txtCustomer.Enabled = False
        Me.txtCustomer.Location = New System.Drawing.Point(106, 332)
        Me.txtCustomer.Name = "txtCustomer"
        Me.txtCustomer.Size = New System.Drawing.Size(442, 22)
        Me.txtCustomer.TabIndex = 14
        '
        'lblcustomer
        '
        Me.lblcustomer.Location = New System.Drawing.Point(10, 332)
        Me.lblcustomer.Name = "lblcustomer"
        Me.lblcustomer.Size = New System.Drawing.Size(76, 27)
        Me.lblcustomer.TabIndex = 38
        Me.lblcustomer.Text = "Customer"
        '
        'LblQtyonLbl
        '
        Me.LblQtyonLbl.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.LblQtyonLbl.Location = New System.Drawing.Point(10, 406)
        Me.LblQtyonLbl.Name = "LblQtyonLbl"
        Me.LblQtyonLbl.Size = New System.Drawing.Size(86, 27)
        Me.LblQtyonLbl.TabIndex = 37
        Me.LblQtyonLbl.Text = "Qty on Label"
        Me.LblQtyonLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LblNoLot
        '
        Me.LblNoLot.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblNoLot.ForeColor = System.Drawing.Color.Red
        Me.LblNoLot.Location = New System.Drawing.Point(346, 212)
        Me.LblNoLot.Name = "LblNoLot"
        Me.LblNoLot.Size = New System.Drawing.Size(96, 19)
        Me.LblNoLot.TabIndex = 36
        '
        'LblQtyLbl
        '
        Me.LblQtyLbl.Location = New System.Drawing.Point(172, 138)
        Me.LblQtyLbl.Name = "LblQtyLbl"
        Me.LblQtyLbl.Size = New System.Drawing.Size(78, 28)
        Me.LblQtyLbl.TabIndex = 35
        Me.LblQtyLbl.Text = "Qty Labels"
        Me.LblQtyLbl.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtLblQty
        '
        Me.TxtLblQty.Location = New System.Drawing.Point(260, 138)
        Me.TxtLblQty.Name = "TxtLblQty"
        Me.TxtLblQty.Size = New System.Drawing.Size(76, 22)
        Me.TxtLblQty.TabIndex = 3
        '
        'LblOrdQty
        '
        Me.LblOrdQty.Enabled = False
        Me.LblOrdQty.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.LblOrdQty.Location = New System.Drawing.Point(220, 249)
        Me.LblOrdQty.Name = "LblOrdQty"
        Me.LblOrdQty.Size = New System.Drawing.Size(78, 27)
        Me.LblOrdQty.TabIndex = 31
        Me.LblOrdQty.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.LblOrdQty.Visible = False
        '
        'Groupporv
        '
        Me.Groupporv.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Groupporv.Controls.Add(Me.ReportViewer1)
        Me.Groupporv.Controls.Add(Me.GroupBox3)
        Me.Groupporv.Controls.Add(Me.btnporvok)
        Me.Groupporv.Controls.Add(Me.TXTVENDCURE)
        Me.Groupporv.Controls.Add(Me.txtStkRoom1)
        Me.Groupporv.Controls.Add(Me.txtLotDate)
        Me.Groupporv.Controls.Add(Me.txtItemDesc)
        Me.Groupporv.Controls.Add(Me.txtBin1)
        Me.Groupporv.Controls.Add(Me.btnPorvPrintCancel)
        Me.Groupporv.Controls.Add(Me.btnPorvLblPrint)
        Me.Groupporv.Controls.Add(Me.TXTLOTQTYPORV)
        Me.Groupporv.Controls.Add(Me.TXTKEY)
        Me.Groupporv.Controls.Add(Me.txtLot)
        Me.Groupporv.Controls.Add(Me.txtPart)
        Me.Groupporv.Controls.Add(Me.datagridStock)
        Me.Groupporv.Controls.Add(Me.txtNoofLabels)
        Me.Groupporv.Controls.Add(Me.lblNoofLabel)
        Me.Groupporv.Controls.Add(Me.lblQtyPerLabel)
        Me.Groupporv.Controls.Add(Me.TextlblqtyPORV)
        Me.Groupporv.Controls.Add(Me.lblPORV)
        Me.Groupporv.ForeColor = System.Drawing.Color.Black
        Me.Groupporv.Location = New System.Drawing.Point(758, 179)
        Me.Groupporv.Name = "Groupporv"
        Me.Groupporv.Size = New System.Drawing.Size(816, 550)
        Me.Groupporv.TabIndex = 43
        Me.Groupporv.TabStop = False
        Me.Groupporv.Text = "PORV  LABELS"
        '
        'ReportViewer1
        '
        Me.ReportViewer1.Location = New System.Drawing.Point(78, 0)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Remote
        Me.ReportViewer1.ServerReport.ReportPath = "/Reports/IT-PendingForTesting/Modified Reports/LabelPrint"
        Me.ReportViewer1.ServerReport.ReportServerUrl = New System.Uri("http://tssblrfsh101/reportserver", System.UriKind.Absolute)
        Me.ReportViewer1.Size = New System.Drawing.Size(541, 423)
        Me.ReportViewer1.TabIndex = 72
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.lblPONo)
        Me.GroupBox3.Controls.Add(Me.lblfrom)
        Me.GroupBox3.Controls.Add(Me.lblto)
        Me.GroupBox3.Controls.Add(Me.TXTPartSelect)
        Me.GroupBox3.Controls.Add(Me.txtPonumber)
        Me.GroupBox3.Controls.Add(Me.dtpTo)
        Me.GroupBox3.Controls.Add(Me.dtpFROM)
        Me.GroupBox3.Controls.Add(Me.RBPOwise)
        Me.GroupBox3.Controls.Add(Me.RBDateWise)
        Me.GroupBox3.Location = New System.Drawing.Point(20, 15)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(738, 46)
        Me.GroupBox3.TabIndex = 39
        Me.GroupBox3.TabStop = False
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(596, 13)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(36, 23)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "Item"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPONo
        '
        Me.lblPONo.Location = New System.Drawing.Point(452, 9)
        Me.lblPONo.Name = "lblPONo"
        Me.lblPONo.Size = New System.Drawing.Size(54, 28)
        Me.lblPONo.TabIndex = 38
        Me.lblPONo.Text = "PO NO."
        Me.lblPONo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblfrom
        '
        Me.lblfrom.Location = New System.Drawing.Point(204, 13)
        Me.lblfrom.Name = "lblfrom"
        Me.lblfrom.Size = New System.Drawing.Size(20, 27)
        Me.lblfrom.TabIndex = 37
        Me.lblfrom.Text = "Fr"
        Me.lblfrom.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblto
        '
        Me.lblto.Location = New System.Drawing.Point(316, 10)
        Me.lblto.Name = "lblto"
        Me.lblto.Size = New System.Drawing.Size(28, 28)
        Me.lblto.TabIndex = 36
        Me.lblto.Text = "To"
        Me.lblto.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TXTPartSelect
        '
        Me.TXTPartSelect.Location = New System.Drawing.Point(632, 13)
        Me.TXTPartSelect.Name = "TXTPartSelect"
        Me.TXTPartSelect.Size = New System.Drawing.Size(100, 22)
        Me.TXTPartSelect.TabIndex = 35
        '
        'txtPonumber
        '
        Me.txtPonumber.Location = New System.Drawing.Point(514, 13)
        Me.txtPonumber.Name = "txtPonumber"
        Me.txtPonumber.Size = New System.Drawing.Size(74, 22)
        Me.txtPonumber.TabIndex = 34
        '
        'dtpTo
        '
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(340, 13)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(104, 22)
        Me.dtpTo.TabIndex = 33
        Me.dtpTo.Value = New Date(2016, 9, 22, 0, 0, 0, 0)
        '
        'dtpFROM
        '
        Me.dtpFROM.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFROM.Location = New System.Drawing.Point(220, 14)
        Me.dtpFROM.Name = "dtpFROM"
        Me.dtpFROM.Size = New System.Drawing.Size(94, 22)
        Me.dtpFROM.TabIndex = 32
        Me.dtpFROM.Value = New Date(2016, 9, 22, 0, 0, 0, 0)
        '
        'RBPOwise
        '
        Me.RBPOwise.AutoSize = True
        Me.RBPOwise.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.75!)
        Me.RBPOwise.Location = New System.Drawing.Point(8, 16)
        Me.RBPOwise.Name = "RBPOwise"
        Me.RBPOwise.Size = New System.Drawing.Size(82, 20)
        Me.RBPOwise.TabIndex = 30
        Me.RBPOwise.TabStop = True
        Me.RBPOwise.Text = "PO Wise"
        Me.RBPOwise.UseVisualStyleBackColor = True
        '
        'RBDateWise
        '
        Me.RBDateWise.AutoSize = True
        Me.RBDateWise.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.75!)
        Me.RBDateWise.Location = New System.Drawing.Point(94, 16)
        Me.RBDateWise.Name = "RBDateWise"
        Me.RBDateWise.Size = New System.Drawing.Size(76, 20)
        Me.RBDateWise.TabIndex = 31
        Me.RBDateWise.TabStop = True
        Me.RBDateWise.Text = "Dt Wise"
        Me.RBDateWise.UseVisualStyleBackColor = True
        '
        'btnporvok
        '
        Me.btnporvok.BackColor = System.Drawing.Color.Brown
        Me.btnporvok.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnporvok.ForeColor = System.Drawing.Color.Yellow
        Me.btnporvok.Location = New System.Drawing.Point(758, 25)
        Me.btnporvok.Name = "btnporvok"
        Me.btnporvok.Size = New System.Drawing.Size(58, 27)
        Me.btnporvok.TabIndex = 29
        Me.btnporvok.Text = "OK"
        Me.btnporvok.UseVisualStyleBackColor = False
        '
        'TXTVENDCURE
        '
        Me.TXTVENDCURE.Location = New System.Drawing.Point(768, 240)
        Me.TXTVENDCURE.Name = "TXTVENDCURE"
        Me.TXTVENDCURE.Size = New System.Drawing.Size(28, 22)
        Me.TXTVENDCURE.TabIndex = 24
        Me.TXTVENDCURE.Visible = False
        '
        'txtStkRoom1
        '
        Me.txtStkRoom1.Location = New System.Drawing.Point(768, 277)
        Me.txtStkRoom1.Name = "txtStkRoom1"
        Me.txtStkRoom1.Size = New System.Drawing.Size(28, 22)
        Me.txtStkRoom1.TabIndex = 23
        Me.txtStkRoom1.Visible = False
        '
        'txtLotDate
        '
        Me.txtLotDate.Location = New System.Drawing.Point(768, 342)
        Me.txtLotDate.Name = "txtLotDate"
        Me.txtLotDate.Size = New System.Drawing.Size(28, 22)
        Me.txtLotDate.TabIndex = 22
        Me.txtLotDate.Visible = False
        '
        'txtItemDesc
        '
        Me.txtItemDesc.Location = New System.Drawing.Point(768, 369)
        Me.txtItemDesc.Name = "txtItemDesc"
        Me.txtItemDesc.Size = New System.Drawing.Size(28, 22)
        Me.txtItemDesc.TabIndex = 21
        Me.txtItemDesc.Visible = False
        '
        'txtBin1
        '
        Me.txtBin1.Location = New System.Drawing.Point(768, 314)
        Me.txtBin1.Name = "txtBin1"
        Me.txtBin1.Size = New System.Drawing.Size(28, 22)
        Me.txtBin1.TabIndex = 20
        Me.txtBin1.Visible = False
        '
        'btnPorvPrintCancel
        '
        Me.btnPorvPrintCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPorvPrintCancel.ForeColor = System.Drawing.Color.IndianRed
        Me.btnPorvPrintCancel.Location = New System.Drawing.Point(692, 485)
        Me.btnPorvPrintCancel.Name = "btnPorvPrintCancel"
        Me.btnPorvPrintCancel.Size = New System.Drawing.Size(66, 32)
        Me.btnPorvPrintCancel.TabIndex = 11
        Me.btnPorvPrintCancel.TabStop = False
        Me.btnPorvPrintCancel.Text = "Exit"
        '
        'btnPorvLblPrint
        '
        Me.btnPorvLblPrint.BackColor = System.Drawing.Color.Brown
        Me.btnPorvLblPrint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPorvLblPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPorvLblPrint.ForeColor = System.Drawing.Color.Yellow
        Me.btnPorvLblPrint.Location = New System.Drawing.Point(586, 485)
        Me.btnPorvLblPrint.Name = "btnPorvLblPrint"
        Me.btnPorvLblPrint.Size = New System.Drawing.Size(96, 32)
        Me.btnPorvLblPrint.TabIndex = 10
        Me.btnPorvLblPrint.TabStop = False
        Me.btnPorvLblPrint.Text = "Print Label"
        Me.btnPorvLblPrint.UseVisualStyleBackColor = False
        '
        'TXTLOTQTYPORV
        '
        Me.TXTLOTQTYPORV.Location = New System.Drawing.Point(536, 448)
        Me.TXTLOTQTYPORV.Name = "TXTLOTQTYPORV"
        Me.TXTLOTQTYPORV.Size = New System.Drawing.Size(182, 22)
        Me.TXTLOTQTYPORV.TabIndex = 16
        '
        'TXTKEY
        '
        Me.TXTKEY.Location = New System.Drawing.Point(16, 448)
        Me.TXTKEY.Name = "TXTKEY"
        Me.TXTKEY.Size = New System.Drawing.Size(106, 22)
        Me.TXTKEY.TabIndex = 15
        '
        'txtLot
        '
        Me.txtLot.Location = New System.Drawing.Point(352, 448)
        Me.txtLot.Name = "txtLot"
        Me.txtLot.Size = New System.Drawing.Size(174, 22)
        Me.txtLot.TabIndex = 14
        '
        'txtPart
        '
        Me.txtPart.Location = New System.Drawing.Point(132, 448)
        Me.txtPart.Name = "txtPart"
        Me.txtPart.Size = New System.Drawing.Size(202, 22)
        Me.txtPart.TabIndex = 13
        '
        'datagridStock
        '
        Me.datagridStock.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.datagridStock.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridStock.CaptionVisible = False
        Me.datagridStock.DataMember = ""
        Me.datagridStock.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridStock.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datagridStock.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.datagridStock.Location = New System.Drawing.Point(16, 66)
        Me.datagridStock.Name = "datagridStock"
        Me.datagridStock.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.datagridStock.ParentRowsVisible = False
        Me.datagridStock.PreferredColumnWidth = 85
        Me.datagridStock.ReadOnly = True
        Me.datagridStock.RowHeadersVisible = False
        Me.datagridStock.Size = New System.Drawing.Size(632, 336)
        Me.datagridStock.TabIndex = 12
        Me.datagridStock.Visible = False
        '
        'txtNoofLabels
        '
        Me.txtNoofLabels.Location = New System.Drawing.Point(356, 488)
        Me.txtNoofLabels.Name = "txtNoofLabels"
        Me.txtNoofLabels.Size = New System.Drawing.Size(76, 22)
        Me.txtNoofLabels.TabIndex = 9
        '
        'lblNoofLabel
        '
        Me.lblNoofLabel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblNoofLabel.Location = New System.Drawing.Point(260, 488)
        Me.lblNoofLabel.Name = "lblNoofLabel"
        Me.lblNoofLabel.Size = New System.Drawing.Size(86, 27)
        Me.lblNoofLabel.TabIndex = 10
        Me.lblNoofLabel.Text = "No. of Labels"
        '
        'lblQtyPerLabel
        '
        Me.lblQtyPerLabel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblQtyPerLabel.Location = New System.Drawing.Point(8, 485)
        Me.lblQtyPerLabel.Name = "lblQtyPerLabel"
        Me.lblQtyPerLabel.Size = New System.Drawing.Size(104, 26)
        Me.lblQtyPerLabel.TabIndex = 9
        Me.lblQtyPerLabel.Text = "Qty Per Label"
        '
        'TextlblqtyPORV
        '
        Me.TextlblqtyPORV.Location = New System.Drawing.Point(120, 488)
        Me.TextlblqtyPORV.Name = "TextlblqtyPORV"
        Me.TextlblqtyPORV.Size = New System.Drawing.Size(86, 22)
        Me.TextlblqtyPORV.TabIndex = 8
        '
        'lblPORV
        '
        Me.lblPORV.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblPORV.Location = New System.Drawing.Point(224, 25)
        Me.lblPORV.Name = "lblPORV"
        Me.lblPORV.Size = New System.Drawing.Size(84, 17)
        Me.lblPORV.TabIndex = 0
        Me.lblPORV.Text = "PO Number"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(576, 102)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 63
        Me.PictureBox2.TabStop = False
        '
        'groupShip
        '
        Me.groupShip.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.groupShip.Controls.Add(Me.Button1)
        Me.groupShip.Controls.Add(Me.txtshipinvdate)
        Me.groupShip.Controls.Add(Me.TextBox6)
        Me.groupShip.Controls.Add(Me.txtcusto)
        Me.groupShip.Controls.Add(Me.txtmobile)
        Me.groupShip.Controls.Add(Me.txtphone)
        Me.groupShip.Controls.Add(Me.txtdept)
        Me.groupShip.Controls.Add(Me.txtdesig)
        Me.groupShip.Controls.Add(Me.txtbuyer)
        Me.groupShip.Controls.Add(Me.btnclear)
        Me.groupShip.Controls.Add(Me.LblInfo)
        Me.groupShip.Controls.Add(Me.TxtInfo)
        Me.groupShip.Controls.Add(Me.BtnShipOK)
        Me.groupShip.Controls.Add(Me.ComboBox1)
        Me.groupShip.Controls.Add(Me.txtconumber)
        Me.groupShip.Controls.Add(Me.txtcustid)
        Me.groupShip.Controls.Add(Me.TextBox5)
        Me.groupShip.Controls.Add(Me.TextBox7)
        Me.groupShip.Controls.Add(Me.Button2)
        Me.groupShip.Controls.Add(Me.btnshipprint)
        Me.groupShip.Controls.Add(Me.txtshipqty)
        Me.groupShip.Controls.Add(Me.txtshipLine)
        Me.groupShip.Controls.Add(Me.txtshiplot)
        Me.groupShip.Controls.Add(Me.txtshippartno)
        Me.groupShip.Controls.Add(Me.DataGridShip)
        Me.groupShip.Controls.Add(Me.txtshiplabel)
        Me.groupShip.Controls.Add(Me.Label1)
        Me.groupShip.Controls.Add(Me.Label2)
        Me.groupShip.Controls.Add(Me.txtshipqpl)
        Me.groupShip.Controls.Add(Me.lblType)
        Me.groupShip.Controls.Add(Me.LblInvoice)
        Me.groupShip.Controls.Add(Me.txtInvoiceNo)
        Me.groupShip.Controls.Add(Me.LblLineNo)
        Me.groupShip.Controls.Add(Me.TxtLineNo)
        Me.groupShip.Controls.Add(Me.txtItem)
        Me.groupShip.ForeColor = System.Drawing.Color.Black
        Me.groupShip.Location = New System.Drawing.Point(10, 178)
        Me.groupShip.Name = "groupShip"
        Me.groupShip.Size = New System.Drawing.Size(748, 551)
        Me.groupShip.TabIndex = 66
        Me.groupShip.TabStop = False
        Me.groupShip.Text = "SHIP LABELS"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(488, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(76, 21)
        Me.Button1.TabIndex = 45
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtshipinvdate
        '
        Me.txtshipinvdate.Location = New System.Drawing.Point(698, 422)
        Me.txtshipinvdate.Name = "txtshipinvdate"
        Me.txtshipinvdate.Size = New System.Drawing.Size(30, 22)
        Me.txtshipinvdate.TabIndex = 44
        Me.txtshipinvdate.Visible = False
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(700, 369)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(40, 22)
        Me.TextBox6.TabIndex = 43
        Me.TextBox6.Visible = False
        '
        'txtcusto
        '
        Me.txtcusto.Location = New System.Drawing.Point(700, 222)
        Me.txtcusto.Name = "txtcusto"
        Me.txtcusto.Size = New System.Drawing.Size(30, 22)
        Me.txtcusto.TabIndex = 42
        Me.txtcusto.Visible = False
        '
        'txtmobile
        '
        Me.txtmobile.Location = New System.Drawing.Point(700, 194)
        Me.txtmobile.Name = "txtmobile"
        Me.txtmobile.Size = New System.Drawing.Size(30, 22)
        Me.txtmobile.TabIndex = 41
        Me.txtmobile.Visible = False
        '
        'txtphone
        '
        Me.txtphone.Location = New System.Drawing.Point(700, 166)
        Me.txtphone.Name = "txtphone"
        Me.txtphone.Size = New System.Drawing.Size(30, 22)
        Me.txtphone.TabIndex = 40
        Me.txtphone.Visible = False
        '
        'txtdept
        '
        Me.txtdept.Location = New System.Drawing.Point(700, 138)
        Me.txtdept.Name = "txtdept"
        Me.txtdept.Size = New System.Drawing.Size(30, 22)
        Me.txtdept.TabIndex = 39
        Me.txtdept.Visible = False
        '
        'txtdesig
        '
        Me.txtdesig.Location = New System.Drawing.Point(700, 111)
        Me.txtdesig.Name = "txtdesig"
        Me.txtdesig.Size = New System.Drawing.Size(30, 22)
        Me.txtdesig.TabIndex = 38
        Me.txtdesig.Visible = False
        '
        'txtbuyer
        '
        Me.txtbuyer.Location = New System.Drawing.Point(700, 83)
        Me.txtbuyer.Name = "txtbuyer"
        Me.txtbuyer.Size = New System.Drawing.Size(30, 22)
        Me.txtbuyer.TabIndex = 37
        Me.txtbuyer.Visible = False
        '
        'btnclear
        '
        Me.btnclear.BackColor = System.Drawing.Color.Brown
        Me.btnclear.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnclear.ForeColor = System.Drawing.Color.Yellow
        Me.btnclear.Location = New System.Drawing.Point(692, 46)
        Me.btnclear.Name = "btnclear"
        Me.btnclear.Size = New System.Drawing.Size(56, 27)
        Me.btnclear.TabIndex = 36
        Me.btnclear.Text = "Clear"
        Me.btnclear.UseVisualStyleBackColor = False
        Me.btnclear.Visible = False
        '
        'LblInfo
        '
        Me.LblInfo.Location = New System.Drawing.Point(10, 445)
        Me.LblInfo.Name = "LblInfo"
        Me.LblInfo.Size = New System.Drawing.Size(86, 27)
        Me.LblInfo.TabIndex = 35
        Me.LblInfo.Text = "Package Info"
        Me.LblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TxtInfo
        '
        Me.TxtInfo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtInfo.Enabled = False
        Me.TxtInfo.Location = New System.Drawing.Point(96, 445)
        Me.TxtInfo.MaxLength = 40
        Me.TxtInfo.Name = "TxtInfo"
        Me.TxtInfo.Size = New System.Drawing.Size(404, 22)
        Me.TxtInfo.TabIndex = 34
        Me.TxtInfo.TabStop = False
        '
        'BtnShipOK
        '
        Me.BtnShipOK.BackColor = System.Drawing.Color.Brown
        Me.BtnShipOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnShipOK.ForeColor = System.Drawing.Color.Yellow
        Me.BtnShipOK.Location = New System.Drawing.Point(692, 18)
        Me.BtnShipOK.Name = "BtnShipOK"
        Me.BtnShipOK.Size = New System.Drawing.Size(56, 27)
        Me.BtnShipOK.TabIndex = 28
        Me.BtnShipOK.Text = "OK"
        Me.BtnShipOK.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ComboBox1.Items.AddRange(New Object() {"A)Item Labels Pre-Invoice", "B)Item Labels Post-Invoice", "C)Box Label", "D)KIT Child Items List"})
        Me.ComboBox1.Location = New System.Drawing.Point(96, 24)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(182, 24)
        Me.ComboBox1.TabIndex = 27
        '
        'txtconumber
        '
        Me.txtconumber.Location = New System.Drawing.Point(700, 256)
        Me.txtconumber.Name = "txtconumber"
        Me.txtconumber.Size = New System.Drawing.Size(40, 22)
        Me.txtconumber.TabIndex = 24
        Me.txtconumber.Visible = False
        '
        'txtcustid
        '
        Me.txtcustid.Location = New System.Drawing.Point(700, 286)
        Me.txtcustid.Name = "txtcustid"
        Me.txtcustid.Size = New System.Drawing.Size(40, 22)
        Me.txtcustid.TabIndex = 23
        Me.txtcustid.Visible = False
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(700, 342)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(40, 22)
        Me.TextBox5.TabIndex = 22
        Me.TextBox5.Visible = False
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(700, 314)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(40, 22)
        Me.TextBox7.TabIndex = 20
        Me.TextBox7.Visible = False
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button2.ForeColor = System.Drawing.Color.IndianRed
        Me.Button2.Location = New System.Drawing.Point(662, 510)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(68, 32)
        Me.Button2.TabIndex = 11
        Me.Button2.TabStop = False
        Me.Button2.Text = "Exit"
        '
        'btnshipprint
        '
        Me.btnshipprint.BackColor = System.Drawing.Color.Brown
        Me.btnshipprint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnshipprint.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnshipprint.ForeColor = System.Drawing.Color.Yellow
        Me.btnshipprint.Location = New System.Drawing.Point(566, 510)
        Me.btnshipprint.Name = "btnshipprint"
        Me.btnshipprint.Size = New System.Drawing.Size(86, 32)
        Me.btnshipprint.TabIndex = 10
        Me.btnshipprint.TabStop = False
        Me.btnshipprint.Text = "Print Label"
        Me.btnshipprint.UseVisualStyleBackColor = False
        '
        'txtshipqty
        '
        Me.txtshipqty.Location = New System.Drawing.Point(508, 482)
        Me.txtshipqty.Name = "txtshipqty"
        Me.txtshipqty.Size = New System.Drawing.Size(184, 22)
        Me.txtshipqty.TabIndex = 16
        '
        'txtshipLine
        '
        Me.txtshipLine.Location = New System.Drawing.Point(10, 482)
        Me.txtshipLine.Name = "txtshipLine"
        Me.txtshipLine.Size = New System.Drawing.Size(96, 22)
        Me.txtshipLine.TabIndex = 15
        '
        'txtshiplot
        '
        Me.txtshiplot.Location = New System.Drawing.Point(326, 482)
        Me.txtshiplot.Name = "txtshiplot"
        Me.txtshiplot.Size = New System.Drawing.Size(174, 22)
        Me.txtshiplot.TabIndex = 14
        '
        'txtshippartno
        '
        Me.txtshippartno.Location = New System.Drawing.Point(116, 482)
        Me.txtshippartno.Name = "txtshippartno"
        Me.txtshippartno.Size = New System.Drawing.Size(200, 22)
        Me.txtshippartno.TabIndex = 13
        '
        'DataGridShip
        '
        Me.DataGridShip.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DataGridShip.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridShip.CaptionVisible = False
        Me.DataGridShip.DataMember = ""
        Me.DataGridShip.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridShip.HeaderFont = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridShip.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridShip.Location = New System.Drawing.Point(10, 68)
        Me.DataGridShip.Name = "DataGridShip"
        Me.DataGridShip.ParentRowsForeColor = System.Drawing.Color.Yellow
        Me.DataGridShip.ParentRowsVisible = False
        Me.DataGridShip.PreferredColumnWidth = 85
        Me.DataGridShip.RowHeadersVisible = False
        Me.DataGridShip.Size = New System.Drawing.Size(682, 374)
        Me.DataGridShip.TabIndex = 12
        '
        'txtshiplabel
        '
        Me.txtshiplabel.Location = New System.Drawing.Point(316, 519)
        Me.txtshiplabel.Name = "txtshiplabel"
        Me.txtshiplabel.Size = New System.Drawing.Size(78, 22)
        Me.txtshiplabel.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(220, 519)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 27)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "No. of Labels"
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(10, 519)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(106, 19)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Qty Per Label"
        '
        'txtshipqpl
        '
        Me.txtshipqpl.Location = New System.Drawing.Point(116, 519)
        Me.txtshipqpl.Name = "txtshipqpl"
        Me.txtshipqpl.Size = New System.Drawing.Size(86, 22)
        Me.txtshipqpl.TabIndex = 8
        '
        'lblType
        '
        Me.lblType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblType.Location = New System.Drawing.Point(10, 24)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(76, 19)
        Me.lblType.TabIndex = 0
        Me.lblType.Text = "Label Type"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Monotype Corsiva", 40.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(586, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(988, 65)
        Me.Label3.TabIndex = 67
        Me.Label3.Text = "            L a b e l    S o f t w a r e"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Goldenrod
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(576, 83)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(998, 19)
        Me.Label4.TabIndex = 68
        Me.Label4.Text = "© 2012 Trelleborg Sealing Solutions (P) Ltd. IT Dept. All rights reserved. Ver 2." & _
    "0 Dt.26th Oct 2012"
        '
        'txtdate1
        '
        Me.txtdate1.AllowDrop = True
        Me.txtdate1.Enabled = False
        Me.txtdate1.Location = New System.Drawing.Point(740, 118)
        Me.txtdate1.Name = "txtdate1"
        Me.txtdate1.Size = New System.Drawing.Size(30, 22)
        Me.txtdate1.TabIndex = 69
        Me.txtdate1.TabStop = False
        Me.txtdate1.Visible = False
        '
        'txtdate2
        '
        Me.txtdate2.AllowDrop = True
        Me.txtdate2.Enabled = False
        Me.txtdate2.Location = New System.Drawing.Point(784, 115)
        Me.txtdate2.Name = "txtdate2"
        Me.txtdate2.Size = New System.Drawing.Size(30, 22)
        Me.txtdate2.TabIndex = 70
        Me.txtdate2.TabStop = False
        Me.txtdate2.Visible = False
        '
        'txtdate3
        '
        Me.txtdate3.AllowDrop = True
        Me.txtdate3.Enabled = False
        Me.txtdate3.Location = New System.Drawing.Point(836, 117)
        Me.txtdate3.Name = "txtdate3"
        Me.txtdate3.Size = New System.Drawing.Size(30, 22)
        Me.txtdate3.TabIndex = 71
        Me.txtdate3.TabStop = False
        Me.txtdate3.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(1910, 902)
        Me.Controls.Add(Me.txtdate3)
        Me.Controls.Add(Me.txtdate2)
        Me.Controls.Add(Me.txtdate1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.groupShip)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.TxtCustPO)
        Me.Controls.Add(Me.LblCustPO)
        Me.Controls.Add(Me.LblCustItem)
        Me.Controls.Add(Me.LblCustDesc)
        Me.Controls.Add(Me.TxtCustDesc)
        Me.Controls.Add(Me.Groupporv)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Groupporv.ResumeLayout(False)
        Me.Groupporv.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.datagridStock, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.groupShip.ResumeLayout(False)
        Me.groupShip.PerformLayout()
        CType(Me.DataGridShip, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Protected Const CONNECTION_STRING As String = "Server=10.56.40.5;Database=FSPrograms;User ID=sa;Password=Trelleborg123"
    'Protected Const CONNECTION_STRING As String = "Server=tssblrfsh101;Database=FSPrograms;User ID=sa;Password=TR3LL3B0RGFSH"

    'Protected Const connection_string As String = "Server = TSSBLRL211;Database=FSDBIN;User ID=sa;password=Tr3ll3b0rgfsh"

    Public Shared ConnectionString As String = CONNECTION_STRING


    Private Sub txtselectreport_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtSelectReport.Leave



    End Sub
    Private Sub txtLblQty_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtLblQty.Leave
        'Dim strCode As String
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        ' Dim strSQL As String

        Dim report As String

        report = Microsoft.VisualBasic.Left(TxtSelectReport.Text, 1)

        If txtInvoiceNo.Text.Length = 0 And report = "A" Then
            MsgBox("Please, enter Shipment no.!", MsgBoxStyle.Critical, "Error!")
            txtInvoiceNo.Focus()
            Exit Sub
        End If

        If txtInvoiceNo.Text.Length = 0 And report = "B" Then
            MsgBox("Please, enter Invoice No.!", MsgBoxStyle.Critical, "Error!")
            txtInvoiceNo.Focus()
            Exit Sub
        End If


        If report = "A" Or report = "B" Then
            If TxtLineNo.Text.Length = 0 Then
                MsgBox("Please, enter line no.!", MsgBoxStyle.Critical, "Error!")
                TxtLineNo.Focus()
                Exit Sub
            End If
        End If

        If TxtLblQty.Text.Length = 0 Then
            TxtLblQty.Text = "1"
        End If

        Dim strSQL As String

        If report = "B" Then


            If CheckFS.Checked = False Then
                strSQL = "Select InvoiceNumber, LineNumber, ShipmentNumber, SerialNumber, " & _
                     "CustomerID,CustomerOrderNumber, CustomerPONumber, ShipQuantity, " & _
                     "ItemNumber, ItemDescription, CustItemNumber, CustItemDesc, LotNo, " & _
                     "NoLot , NoCustItem, ItemUM, CustomerName " & _
                     "from TSS_ItemLabel " & _
                     "where InvoiceNumber  = '" & txtInvoiceNo.Text & "' and " & _
                     "LineNumber = '" & TxtLineNo.Text & "' "

                'TSS_ItemLabel_PartWise_Ship()

            ElseIf CheckFS.Checked = True Then
                'changes done on 12-1-2011

                'strSQL = "Select InvoiceNumber, LineNumber, ShipmentNumber, SerialNumber, " & _
                '            "CustomerID,CustomerOrderNumber, CustomerPONumber, ShipQuantity, " & _
                '            "ItemNumber, ItemDescription, CustItemNumber, CustItemDesc, LotNo, " & _
                '            "NoLot , NoCustItem, ItemUM, CustomerName " & _
                '            "from TSS_ItemLabel_WH_LotWise" & _
                '            "where InvoiceNumber  = '" & txtInvoiceNo.Text & "' and " & _
                '            "LineNumber = '" & TxtLineNo.Text & "' "

                strSQL = "Select  InvoiceNumber, COLineNumber, ShipmentNumber, SerialNumber, " & _
                            "CustomerOrderNumber,  ShipQty, " & _
                            "ItemNumber, ItemDescription, CustItemNumber, CustItemDesc, LotNumber, " & _
                            "NoLot , NoCustItem, ItemUM, CustomerName " & _
                            "from TSS_ItemLabel_WH_LotWise " & _
                            "where InvoiceNumber  = '" & txtInvoiceNo.Text & "' and " & _
                            "COLineNumber = '" & TxtLineNo.Text & "' "

            End If




        ElseIf report = "A" Then

            strSQL = "Select COLineNumber, ShipmentNumber, " & _
              "CustomerID,CONumber, CustomerPONumber, ShippedQuantity, " & _
              "ItemNumber, ItemDescription, CustItemNumber, CustItemDesc, LotNumber, " & _
              "ItemUM, CustomerName " & _
              "from TSS_ItemLabel2 " & _
              "where ShipmentNumber  = '" & txtInvoiceNo.Text & "' and " & _
              "COLineNumber = '" & TxtLineNo.Text & "' "

        ElseIf report = "C" Then


            If CheckFS.Checked = False Then
                strSQL = "Select ShipmentNumber,InvoiceDate,CustomerPONumber, ConsigneeName,ShipmentAddress1," & _
                         "ShipmentAddress2,ShipmentCity, ShipmentState,ShipmentZip," & _
                         "CustomerContact, CustomerContactPhone, cod from TSS_BOXLabel_DATA where " & _
                         "InvoiceNumber = '" & txtInvoiceNo.Text & "' "
            ElseIf CheckFS.Checked = True Then
                '-changes done on 12-1-2011  TSS_BOXLabel_DATA_WH

                strSQL = "Select  ShipmentNumber,InvoiceDate,CustomerPONumber, ConsigneeName,ShipmentAddress1," & _
                                "ShipmentAddress2,ShipmentCity, ShipmentState,ShipmentZip," & _
                                "CustomerContact, CustomerContactPhone, cod from TSS_BOXLabel_DATA_WH_VER2 where " & _
                                "InvoiceNumber = '" & txtInvoiceNo.Text & "' "

            End If

        End If


        cnSQL = New SqlConnection(ConnectionString)
        cnSQL.Open()

        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()
        Dim a As Integer
        Dim b As Integer
        b = 0
        a = drSQL.Item(0)
        a = drSQL.Item(0).count




        '  a = drSQL.FieldCount
        a = drSQL(0)


        Try
            If drSQL.Read() Then
                Do While a >= b



                    If report = "C" Then
                        TxtCustOrdNo.Text = drSQL.Item(0) 'shipmentnumber
                        txtQty.Text = drSQL.Item(1) 'invoicedate 
                        TxtCustPO.Text = drSQL.Item(2) 'custpo  
                        txtItem.Text = drSQL.Item(3) 'cust name
                        TxtDescription.Text = drSQL.Item(4) 'ship ad1
                        TxtCustItem.Text = drSQL.Item(5) 'ship ad2
                        TxtCustDesc.Text = drSQL.Item(6) 'shipcity
                        txtLotNo.Text = drSQL.Item(7)   'shipstate
                        TxtUoM.Text = drSQL.Item(8) 'shipzip
                        txtCustomer.Text = drSQL.Item(9) 'customercontact
                        TxtCOLnNo.Text = drSQL.Item(10) 'customercontactphone.

                        If Len(LTrim(RTrim(drSQL.Item(11)))) > 1 Then
                            TxtInfo.Text = drSQL.Item(11) 'cod pay details
                        End If

                        ' TxtCustDesc.Text = drSQL.Item(11) 'cod pay details
                        If TxtLblQty.Text < "1" Then
                            ' MsgBox("Please enter Label Quantity", MsgBoxStyle.Critical, "Label Qty Err")
                            TxtLblQty.Text = "1"
                            BtnPrtMulti.Enabled = False
                            btnPrint.Enabled = True
                            BtnPrtMulti.Text = ""
                        End If

                        If TxtLblQty.Text > "1" Then
                            BtnPrtMulti.Enabled = True
                            'BtnPrtMulti.Focus()
                            BtnPrtMulti.Text = "Print " & Microsoft.VisualBasic.FormatNumber(TxtLblQty.Text, 0) & " Labels"
                        Else
                            btnPrint.Enabled = True
                            'btnPrint.Focus()
                            BtnPrtMulti.Enabled = False
                            BtnPrtMulti.Text = ""
                        End If


                        TxtInfo.Enabled = True
                        TxtInfo.Focus()
                        Exit Sub

                    ElseIf report = "B" Then

                        TxtCustOrdNo.Text = drSQL.Item(5)
                        TxtCustPO.Text = drSQL.Item(6)

                        txtQty.Text = drSQL.Item(7).ToString()
                        txtQty.Text = FormatNumber(txtQty.Text, 3)
                        txtItem.Text = Trim(drSQL.Item(8).ToString())
                        TxtDescription.Text = drSQL.Item(9)
                        TxtCustItem.Text = drSQL.Item(10)
                        TxtCustDesc.Text = drSQL.Item(11)
                        txtLotNo.Text = drSQL.Item(12)
                        LblNoLot.Text = drSQL.Item(13)
                        LblNoCust.Text = drSQL.Item(14)
                        lblQtyOnLabel.Text = (txtQty.Text / TxtLblQty.Text)
                        TxtUoM.Text = drSQL.Item(15)
                        txtCustomer.Text = drSQL.Item(16)
                    ElseIf report = "A" Then

                        TxtCustOrdNo.Text = drSQL.Item(3)
                        TxtCustPO.Text = drSQL.Item(4)
                        txtQty.Text = drSQL.Item(5).ToString()
                        txtQty.Text = FormatNumber(txtQty.Text, 3)
                        txtItem.Text = Trim(drSQL.Item(6).ToString())
                        TxtDescription.Text = drSQL.Item(7)
                        TxtCustItem.Text = drSQL.Item(8)
                        TxtCustDesc.Text = drSQL.Item(9)
                        txtLotNo.Text = drSQL.Item(10)
                        'LblNoLot.Text = drSQL.Item(13)
                        'LblNoCust.Text = drSQL.Item(14)
                        lblQtyOnLabel.Text = (txtQty.Text / TxtLblQty.Text)
                        TxtUoM.Text = drSQL.Item(11)
                        txtCustomer.Text = drSQL.Item(12)

                    End If
                    b = b + 1
                Loop


            Else
                MsgBox("Wrong data entered! Check the Invoice number! ", MsgBoxStyle.Exclamation, "Error!")
                ClearAll()
                Exit Sub
            End If





            If TxtLblQty.Text < "1" Then
                ' MsgBox("Please enter Label Quantity", MsgBoxStyle.Critical, "Label Qty Err")
                TxtLblQty.Text = "1"
                BtnPrtMulti.Enabled = False
                BtnPrtMulti.Text = ""
            End If

            If TxtLblQty.Text > "1" Then
                BtnPrtMulti.Enabled = True
                BtnPrtMulti.Text = "Print " & Microsoft.VisualBasic.FormatNumber(TxtLblQty.Text, 0) & " Labels"
                TxtInfo.Enabled = True
                TxtInfo.Focus()
            Else
                BtnPrtMulti.Enabled = False
                BtnPrtMulti.Text = ""
                TxtInfo.Enabled = True
                TxtInfo.Focus()
            End If


            btnPrint.Enabled = True
            btnCancel.Enabled = True

            drSQL.Close()
            cnSQL.Close()
            cmSQL.Dispose()
            cnSQL.Dispose()


        Catch
            MsgBox("Wrong data entered! Check the Invoice number! ", MsgBoxStyle.Exclamation, "Error!")
            ClearAll()
        End Try

    End Sub
    Private Sub txtCOlnNo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCOLnNo.Leave
        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String

        If TxtCustOrdNo.Text.Length = 0 Then
            MsgBox("Please, enter Order no.!", MsgBoxStyle.Critical, "Error!")
            TxtCustOrdNo.Focus()
            Exit Sub
        End If

        If TxtCOLnNo.Text.Length = 0 Then
            MsgBox("Please, enter line no.!", MsgBoxStyle.Critical, "Error!")
            TxtCustOrdNo.Focus()
            Exit Sub
        End If


        Try
            strSQL = "Select H.CONumber,H.CustomerPONumber,L.COLineNumber,I.ItemNumber,I.ItemDescription,I.ItemUM, L.ItemOrderedQuantity, " & _
                     "isnull(CI.CustomerItemNumber,I.ItemNumber) as 'CustItemNumber', " & _
                     "isnull(CI.CustomerItemDescription,I.ItemDescription) as 'CustItemDesc', " & _
                     "case when isnull(CI.CustomerItemNumber,'no') = 'no' then 'Customer Item not defined' else '' end as 'NoCustItem' " & _
                    " from FS_COHeader H " & _
                    " Left outer join FS_COLine L on L.COHeaderKey = H.COHeaderKey " & _
                    " Left outer join FS_Item I on I.ItemKey = L.ItemKey " & _
                    " left outer join FS_CustomerItem CI on CI.ItemKey = L.ItemKey and CI.CustomerKey = H.CustomerKey " & _
                    " where H.CONumber  = '" & TxtCustOrdNo.Text & "' and " & _
                     " L.COLineNumber = '" & TxtCOLnNo.Text & "' "


            cnSQL = New SqlConnection(ConnectionString)
            cnSQL.Open()

            cmSQL = New SqlCommand(strSQL, cnSQL)
            drSQL = cmSQL.ExecuteReader()

            If drSQL.Read() Then

                TxtCustPO.Text = drSQL.Item(1)
                txtItem.Text = Trim(drSQL.Item(3).ToString())
                TxtDescription.Text = drSQL.Item(4)
                TxtUoM.Text = drSQL.Item(5)
                LblOrdQty.Text = drSQL.Item(6).ToString()
                TxtCustItem.Text = drSQL.Item(7)
                TxtCustDesc.Text = drSQL.Item(8)
                LblNoCust.Text = drSQL.Item(9)


                btnPrint.Enabled = True
                BtnPrtMulti.Enabled = True
                btnCancel.Enabled = True

                LblOrdQty.Visible = True
                txtQty.Enabled = True
                txtLotNo.Enabled = True
                txtLotNo.TabStop = True
                txtQty.Focus()
                '   BtnPrtMulti.Text = "Print " & Microsoft.VisualBasic.FormatNumber(txtQty.Text, 0) & " Labels"
                BtnPrtMulti.Visible = False

            Else
                MsgBox("Wrong data entered! Check the Customer Order number! ", MsgBoxStyle.Exclamation, "Error!")
                ClearAll()
            End If

            drSQL.Close()
            cnSQL.Close()
            cmSQL.Dispose()
            cnSQL.Dispose()
        Catch
            MsgBox("Wrong data entered! Check the Customer Order number! ", MsgBoxStyle.Exclamation, "Error!")
            ClearAll()
        End Try
    End Sub

    Sub PostInvoice()
        'Prepare fields for PostInvoice label
        textvisibleT()

        'TxtCustOrdNo.Text = ""
        'TxtCustPO.Text = ""
        'txtQty.Text = ""
        'txtItem.Text = ""
        'TxtDescription.Text = ""
        'TxtCustItem.Text = ""
        'TxtCustDesc.Text = ""
        'txtLotNo.Text = ""
        'LblNoLot.Text = ""
        'LblNoCust.Text = ""
        'BtnPrtMulti.Text = ""
        'TxtUoM.Text = ""
        'txtInvoiceNo.Text = ""
        'TxtLineNo.Text = ""
        'txtCustomer.Text = ""
        'LblOrdQty.Text = ""
        'TxtLblQty.Text = ""
        'LblQtyOnLabel.Text = ""
        'TxtInfo.Text = ""
        LblQtyLbl.Visible = True
        TxtLblQty.Visible = True
        lblQtyOnLabel.Visible = True
        LblQtyonLbl.Visible = True
        TxtInfo.Visible = True
        TxtInfo.Enabled = True
        LblInfo.Visible = True
        txtInvoiceNo.Visible = True
        TxtLineNo.Visible = True
        LblInvoice.Visible = True
        LblLineNo.Visible = True
        BtnPrtMulti.Visible = True
        TxtCustOrdNo.Enabled = False
        TxtCOLnNo.Enabled = False
        txtQty.Enabled = False
        txtLotNo.Enabled = False
        TxtCOLnNo.TabStop = False
        BtnPrtMulti.Enabled = False
        btnPrint.Enabled = False
        btnCancel.Enabled = False

        txtInvoiceNo.Focus()
    End Sub
    Sub PreInvoice()
        'Prepare fields for Preinvoice label
        textvisibleT()

        'TxtCustOrdNo.Text = ""
        'TxtCustPO.Text = ""
        'txtQty.Text = ""
        'txtItem.Text = ""
        'TxtDescription.Text = ""
        'TxtCustItem.Text = ""
        'TxtCustDesc.Text = ""
        'txtLotNo.Text = ""
        'LblNoLot.Text = ""
        'LblNoCust.Text = ""
        'BtnPrtMulti.Text = ""
        'TxtUoM.Text = ""
        'txtInvoiceNo.Text = ""
        'TxtLineNo.Text = ""
        'LblOrdQty.Text = ""
        'TxtLblQty.Text = ""
        'LblQtyOnLabel.Text = ""
        TxtCustOrdNo.Focus()
        TxtCOLnNo.TabStop = True
    End Sub
    Sub kitlabel()
        textvisibleF()
        'TxtCustOrdNo.Text = ""
        'TxtCustPO.Text = ""
        'txtQty.Text = ""
        'txtItem.Text = ""
        'TxtDescription.Text = ""
        'TxtCustItem.Text = ""
        'TxtCustDesc.Text = ""
        'txtLotNo.Text = ""
        'LblNoLot.Text = ""
        'LblNoCust.Text = ""
        'BtnPrtMulti.Text = ""
        'TxtUoM.Text = ""
        'txtInvoiceNo.Text = ""
        'TxtLineNo.Text = ""
        'LblOrdQty.Text = ""
        'TxtLblQty.Text = ""
        'LblQtyOnLabel.Text = ""

        'txtInvoiceNo.Focus()
    End Sub
    Sub boxlabel()
        textvisibleF()
        LblInfo.Visible = True
        TxtInfo.Visible = True
        LblInfo.Text = "Transporter:"
        Exit Sub
    End Sub
    Sub boxlabelpacking()
        MsgBox("Under development")
        Exit Sub
    End Sub
    Sub textvisibleF()

        TxtLineNo.Visible = False
        LblLineNo.Visible = False

        TxtCustOrdNo.Visible = False
        LblCONo.Visible = False

        TxtCOLnNo.Visible = False
        LblCOLnNo.Visible = False

        LblCustPO.Visible = False
        TxtCustPO.Visible = False

        txtQty.Visible = False
        lblQty.Visible = False

        txtItem.Visible = False
        lblItem.Visible = False

        TxtDescription.Visible = False
        LblDescription.Visible = False

        TxtCustItem.Visible = False
        LblCustItem.Visible = False

        TxtCustDesc.Visible = False
        LblCustDesc.Visible = False

        txtLotNo.Visible = False
        lblLotNo.Visible = False

        LblNoLot.Visible = False
        LblNoCust.Visible = False
        BtnPrtMulti.Text = ""

        TxtUoM.Visible = False
        lblUOM.Visible = False

        'LblOrdQty.Text = ""
        'TxtInfo.Text = ""
        'TxtLblQty.Text = ""

        lblQtyOnLabel.Visible = False
        LblQtyonLbl.Visible = False

        txtCustomer.Visible = False
        lblcustomer.Visible = False

        txtInvoiceNo.Enabled = True
        txtInvoiceNo.Focus()
    End Sub
    Sub textvisibleT()

        TxtLineNo.Visible = True
        LblLineNo.Visible = True

        TxtCustOrdNo.Visible = True
        LblCONo.Visible = True

        TxtCOLnNo.Visible = True
        LblCOLnNo.Visible = True

        LblCustPO.Visible = True
        TxtCustPO.Visible = True

        txtQty.Visible = True
        lblQty.Visible = True

        txtItem.Visible = True
        lblItem.Visible = True

        TxtDescription.Visible = True
        LblDescription.Visible = True

        TxtCustItem.Visible = True
        LblCustItem.Visible = True

        TxtCustDesc.Visible = True
        LblCustDesc.Visible = True

        txtLotNo.Visible = True
        lblLotNo.Visible = True

        LblNoLot.Visible = True
        LblNoCust.Visible = True
        BtnPrtMulti.Text = ""

        TxtUoM.Visible = True
        lblUOM.Visible = True

        'LblOrdQty.Text = ""
        'TxtInfo.Text = ""
        'TxtLblQty.Text = ""

        lblQtyOnLabel.Visible = True
        LblQtyonLbl.Visible = True

        txtCustomer.Visible = True
        lblcustomer.Visible = True
    End Sub


    Sub ClearAll()
        'Clear screen boxes
        'Leave Invoice number

        TxtCustOrdNo.Text = ""
        TxtCustPO.Text = ""
        txtQty.Text = ""
        txtItem.Text = ""
        TxtDescription.Text = ""
        TxtCustItem.Text = ""
        TxtCustDesc.Text = ""
        txtLotNo.Text = ""
        LblNoLot.Text = ""
        LblNoCust.Text = ""
        BtnPrtMulti.Text = ""
        TxtUoM.Text = ""
        LblOrdQty.Text = ""
        TxtInfo.Text = ""
        TxtLblQty.Text = ""
        txtCustomer.Text = ""
        'txtInvoiceNo.Text = ""
        'TxtLineNo.Text = ""

        BtnPrtMulti.Enabled = False
        BtnPrtMulti.Visible = True
        btnPrint.Enabled = False
        btnCancel.Enabled = False

        TxtSelectReport.Focus()
    End Sub
    Sub ClearEverything()
        'Empty all fields
        TxtCustOrdNo.Text = ""
        TxtCustPO.Text = ""
        txtQty.Text = ""
        txtItem.Text = ""
        TxtDescription.Text = ""
        TxtCustItem.Text = ""
        TxtCustDesc.Text = ""
        txtLotNo.Text = ""
        LblNoLot.Text = ""
        LblNoCust.Text = ""
        BtnPrtMulti.Text = ""
        TxtUoM.Text = ""
        txtInvoiceNo.Text = ""
        TxtLineNo.Text = ""
        TxtSelectReport.Text = ""
        LblOrdQty.Text = ""
        TxtInfo.Text = ""
        TxtLblQty.Text = ""
        lblQtyOnLabel.Text = ""
        txtInvoiceNo.Text = ""
        txtCustomer.Text = ""
        TxtLblQty.Visible = False
        LblQtyLbl.Visible = False
        txtInvoiceNo.Visible = True
        TxtLineNo.Visible = True
        LblInvoice.Visible = True
        LblLineNo.Visible = True
        TxtCustOrdNo.Enabled = False
        TxtCOLnNo.Enabled = False
        txtQty.Enabled = False
        txtLotNo.Enabled = False
        BtnPrtMulti.Visible = True
        BtnPrtMulti.Enabled = False
        btnPrint.Enabled = False
        btnCancel.Enabled = False

        TxtSelectReport.Focus()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        'Print one label
        Dim report As String
        report = Microsoft.VisualBasic.Left(TxtSelectReport.Text, 1)

        Dim rptLabel As New ReportDocument
        Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection2 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection3 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection4 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection5 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection6 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection7 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection8 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection9 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection10 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection11 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection12 As New CrystalDecisions.Shared.ParameterValues
        Dim PVcollection13 As New CrystalDecisions.Shared.ParameterValues

        Dim pdvPartNo As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvPartName As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvQty As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvInvoiceNo As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvSINo As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvMfgRef As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvCustPO As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvUom As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvItemNumber As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvItemDescription As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvPackInfo As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvcustomer As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvoption As New CrystalDecisions.Shared.ParameterDiscreteValue

        '    Dim path As String
        '   path = System.IO.Path.GetDirectoryName( _
        '     System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) 
        ' MessageBox.Show(path)


        'load the report template to document object be sure that report is in main directory
        If report = "A" Or report = "B" Then
            rptLabel.Load("Label2.rpt")
        ElseIf report = "C" Then
            rptLabel.Load("BoxLabel.rpt")
        End If

        ' Set the discreet value to the order name and line no

        If report = "A" Or report = "B" Then

            pdvPartNo.Value = Trim(TxtCustItem.Text)
            pdvPartName.Value = Trim(TxtCustDesc.Text)
            pdvQty.Value = Trim(txtQty.Text)
            pdvInvoiceNo.Value = Trim(txtInvoiceNo.Text)
            pdvSINo.Value = Trim(TxtLineNo.Text)
            pdvMfgRef.Value = Trim(txtLotNo.Text)
            pdvCustPO.Value = Trim(TxtCustPO.Text)
            pdvUom.Value = Trim(TxtUoM.Text)
            pdvItemNumber.Value = Trim(txtItem.Text)
            pdvItemDescription.Value = Trim(TxtDescription.Text)
            pdvPackInfo.Value = Trim(TxtInfo.Text)
            pdvcustomer.Value = Trim(txtCustomer.Text)

        ElseIf report = "C" Then ' box label printing

            pdvPackInfo.Value = Trim(TxtInfo.Text) ' Transport details
            pdvInvoiceNo.Value = Trim(txtInvoiceNo.Text) 'ok invoice
            pdvPartNo.Value = Trim(TxtCustOrdNo.Text) ' shipmentnumber 
            pdvPartName.Value = Trim(txtQty.Text) 'invoice date
            pdvCustPO.Value = Trim(TxtCustPO.Text) 'ok customer po
            pdvItemNumber.Value = Trim(txtItem.Text) 'custname, shipto
            pdvItemDescription.Value = Trim(TxtDescription.Text) 'sp ad1
            pdvQty.Value = Trim(TxtCustItem.Text) 'ship ad2
            pdvSINo.Value = Trim(TxtCustDesc.Text) 'shipcity
            pdvMfgRef.Value = Trim(txtLotNo.Text) 'ship state
            pdvUom.Value = Trim(TxtUoM.Text) 'ship zip
            pdvcustomer.Value = Trim(txtCustomer.Text) ' customr contact
            pdvoption.Value = Trim(TxtCOLnNo.Text) ' customr contact

        Else
            MsgBox("Program is not yet ready", vbInformation)
            Exit Sub
        End If

        If report = "A" Then
            pdvoption.Value = "A"
        ElseIf report = "B" Then
            pdvoption.Value = "B"

        End If


        ' Add it to the parameter collection.
        pvCollection.Add(pdvPartNo)
        pvCollection2.Add(pdvPartName)
        pvCollection3.Add(pdvQty)
        pvCollection4.Add(pdvInvoiceNo)
        pvCollection5.Add(pdvSINo)
        pvCollection6.Add(pdvMfgRef)
        pvCollection7.Add(pdvCustPO)
        pvCollection8.Add(pdvUom)
        pvCollection9.Add(pdvItemNumber)
        pvCollection10.Add(pdvItemDescription)
        pvCollection11.Add(pdvPackInfo)
        pvCollection12.Add(pdvcustomer)
        PVcollection13.Add(pdvoption)
        ' Apply the current parameter values.
        rptLabel.DataDefinition.ParameterFields("PartNo").ApplyCurrentValues(pvCollection)
        rptLabel.DataDefinition.ParameterFields("PartName").ApplyCurrentValues(pvCollection2)
        rptLabel.DataDefinition.ParameterFields("Qty").ApplyCurrentValues(pvCollection3)

        rptLabel.DataDefinition.ParameterFields("InvoiceNo").ApplyCurrentValues(pvCollection4)
        rptLabel.DataDefinition.ParameterFields("SINo").ApplyCurrentValues(pvCollection5)
        rptLabel.DataDefinition.ParameterFields("MfgRef").ApplyCurrentValues(pvCollection6)
        rptLabel.DataDefinition.ParameterFields("CustPO").ApplyCurrentValues(pvCollection7)
        rptLabel.DataDefinition.ParameterFields("UoM").ApplyCurrentValues(pvCollection8)
        rptLabel.DataDefinition.ParameterFields("ItemNumber").ApplyCurrentValues(pvCollection9)
        rptLabel.DataDefinition.ParameterFields("ItemDescription").ApplyCurrentValues(pvCollection10)
        rptLabel.DataDefinition.ParameterFields("PackInfo").ApplyCurrentValues(pvCollection11)
        rptLabel.DataDefinition.ParameterFields("customer").ApplyCurrentValues(pvCollection12)
        rptLabel.DataDefinition.ParameterFields("option").ApplyCurrentValues(PVcollection13)


        rptLabel.PrintToPrinter(1, False, 0, 0)

        '   rptLabel.ExportToDisk(CrystalDecisions.[Shared].ExportFormatType.RichText, "c:\label.txt")

        'rptLabel.ExportToDisk(CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat, "c:\label.pdf")

        pvCollection.Clear()
        pvCollection2.Clear()
        pvCollection3.Clear()
        pvCollection4.Clear()
        pvCollection5.Clear()
        pvCollection6.Clear()
        pvCollection7.Clear()
        pvCollection8.Clear()
        pvCollection9.Clear()
        pvCollection10.Clear()
        pvCollection11.Clear()
        pvCollection12.Clear()
        PVcollection13.Clear()

        rptLabel.Close()


    End Sub

    Public Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        ClearEverything()
    End Sub
    <System.STAThread()> _
    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.Run(New frmMain)
    End Sub

    Private Sub BtnPrtMulti_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrtMulti.Click

        Dim report As String
        report = Microsoft.VisualBasic.Left(TxtSelectReport.Text, 1)

        Dim copy As Integer
        Dim Qty As String
        '  copy = CInt(txtQty.Text)
        copy = CInt(TxtLblQty.Text)

        If report <> "C" Then
            Qty = txtQty.Text / TxtLblQty.Text
        End If

        '-new coding copiied.

        Dim rptLabel As New ReportDocument
        Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection2 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection3 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection4 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection5 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection6 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection7 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection8 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection9 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection10 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection11 As New CrystalDecisions.Shared.ParameterValues
        Dim pvCollection12 As New CrystalDecisions.Shared.ParameterValues
        Dim PVcollection13 As New CrystalDecisions.Shared.ParameterValues

        Dim pdvPartNo As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvPartName As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvQty As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvInvoiceNo As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvSINo As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvMfgRef As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvCustPO As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvUom As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvItemNumber As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvItemDescription As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvPackInfo As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvcustomer As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim pdvoption As New CrystalDecisions.Shared.ParameterDiscreteValue

        '    Dim path As String
        '   path = System.IO.Path.GetDirectoryName( _
        '     System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) 
        ' MessageBox.Show(path)


        'load the report template to document object be sure that report is in main directory
        If report = "A" Or report = "B" Then
            rptLabel.Load("Label2.rpt")
        ElseIf report = "C" Then
            rptLabel.Load("BoxLabel.rpt")
        End If

        ' Set the discreet value to the order name and line no

        If report = "A" Or report = "B" Then

            pdvPartNo.Value = Trim(TxtCustItem.Text)
            pdvPartName.Value = Trim(TxtCustDesc.Text)
            '    pdvQty.Value = Trim(txtQty.Text)
            pdvQty.Value = Trim(lblQtyOnLabel.Text)
            pdvInvoiceNo.Value = Trim(txtInvoiceNo.Text)
            pdvSINo.Value = Trim(TxtLineNo.Text)
            pdvMfgRef.Value = Trim(txtLotNo.Text)
            pdvCustPO.Value = Trim(TxtCustPO.Text)
            pdvUom.Value = Trim(TxtUoM.Text)
            pdvItemNumber.Value = Trim(txtItem.Text)
            pdvItemDescription.Value = Trim(TxtDescription.Text)
            pdvPackInfo.Value = Trim(TxtInfo.Text)
            pdvcustomer.Value = Trim(txtCustomer.Text)

        ElseIf report = "C" Then ' box label printing


            pdvPackInfo.Value = Trim(TxtInfo.Text) ' Transport details
            pdvInvoiceNo.Value = Trim(txtInvoiceNo.Text) 'ok invoice
            pdvPartNo.Value = Trim(TxtCustOrdNo.Text) ' shipmentnumber 
            pdvPartName.Value = Trim(txtQty.Text) 'invoice date
            pdvCustPO.Value = Trim(TxtCustPO.Text) 'ok customer po
            pdvItemNumber.Value = Trim(txtItem.Text) 'custname, shipto
            pdvItemDescription.Value = Trim(TxtDescription.Text) 'sp ad1
            pdvQty.Value = Trim(TxtCustItem.Text) 'ship ad2
            pdvSINo.Value = Trim(TxtCustDesc.Text) 'shipcity
            pdvMfgRef.Value = Trim(txtLotNo.Text) 'ship state
            pdvUom.Value = Trim(TxtUoM.Text) 'ship zip
            pdvcustomer.Value = Trim(txtCustomer.Text) ' customr contact
            pdvoption.Value = Trim(TxtCOLnNo.Text) ' customr contact

        Else
            MsgBox("Program is not yet ready", vbInformation)
            Exit Sub
        End If

        If report = "A" Then
            pdvoption.Value = "A"
        ElseIf report = "B" Then
            pdvoption.Value = "B"

        End If


        ' Add it to the parameter collection.
        pvCollection.Add(pdvPartNo)
        pvCollection2.Add(pdvPartName)
        pvCollection3.Add(pdvQty)
        pvCollection4.Add(pdvInvoiceNo)
        pvCollection5.Add(pdvSINo)
        pvCollection6.Add(pdvMfgRef)
        pvCollection7.Add(pdvCustPO)
        pvCollection8.Add(pdvUom)
        pvCollection9.Add(pdvItemNumber)
        pvCollection10.Add(pdvItemDescription)
        pvCollection11.Add(pdvPackInfo)
        pvCollection12.Add(pdvcustomer)
        PVcollection13.Add(pdvoption)
        ' Apply the current parameter values.
        rptLabel.DataDefinition.ParameterFields("PartNo").ApplyCurrentValues(pvCollection)
        rptLabel.DataDefinition.ParameterFields("PartName").ApplyCurrentValues(pvCollection2)
        rptLabel.DataDefinition.ParameterFields("Qty").ApplyCurrentValues(pvCollection3)

        rptLabel.DataDefinition.ParameterFields("InvoiceNo").ApplyCurrentValues(pvCollection4)
        rptLabel.DataDefinition.ParameterFields("SINo").ApplyCurrentValues(pvCollection5)
        rptLabel.DataDefinition.ParameterFields("MfgRef").ApplyCurrentValues(pvCollection6)
        rptLabel.DataDefinition.ParameterFields("CustPO").ApplyCurrentValues(pvCollection7)
        rptLabel.DataDefinition.ParameterFields("UoM").ApplyCurrentValues(pvCollection8)
        rptLabel.DataDefinition.ParameterFields("ItemNumber").ApplyCurrentValues(pvCollection9)
        rptLabel.DataDefinition.ParameterFields("ItemDescription").ApplyCurrentValues(pvCollection10)
        rptLabel.DataDefinition.ParameterFields("PackInfo").ApplyCurrentValues(pvCollection11)
        rptLabel.DataDefinition.ParameterFields("customer").ApplyCurrentValues(pvCollection12)
        rptLabel.DataDefinition.ParameterFields("option").ApplyCurrentValues(PVcollection13)


        rptLabel.PrintToPrinter(copy, False, 0, 0)


        ' rptLabel.ExportToDisk(CrystalDecisions.[Shared].ExportFormatType.RichText, "c:\label.txt")

        'rptLabel.ExportToDisk(CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat, "c:\label.pdf")

        pvCollection.Clear()
        pvCollection2.Clear()
        pvCollection3.Clear()
        pvCollection4.Clear()
        pvCollection5.Clear()
        pvCollection6.Clear()
        pvCollection7.Clear()
        pvCollection8.Clear()
        pvCollection9.Clear()
        pvCollection10.Clear()
        pvCollection11.Clear()
        pvCollection12.Clear()
        PVcollection13.Clear()

        rptLabel.Close()





    End Sub
    Private Sub TxtSelectReport_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtSelectReport.SelectedIndexChanged


    End Sub
    Private Sub porvclear()

        TXTKEY.Clear()
        txtPart.Clear()
        txtLot.Clear()
        TXTLOTQTYPORV.Clear()
        TextlblqtyPORV.Clear()
        txtNoofLabels.Clear()

    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LblInvoice.Text = "Invoice No."
        RDBAutomatic.Checked = True
        ListViewPart.Visible = False
        RdbShip.Checked = False
        RdbPORV.Checked = False
        rdbDCChecking.Checked = False
        groupShip.Enabled = False
        Groupporv.Enabled = False
        dtpFROM.Value = Today()
        dtpTo.Value = Today()



        ' Me.ReportViewer1.RefreshReport()
        'Me.ReportViewer1.RefreshReport()
        'Me.ReportViewer2.RefreshReport()
    End Sub

    Sub shipdatatclear()

        Dim sqlCon As SqlConnection
        sqlCon = New SqlConnection(ConnectionString)

        Dim strSql As String
        Dim stockDS As DataSet = New DataSet


        strSql = ""
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDA As SqlDataAdapter = New SqlDataAdapter
        'Try
        stockDA.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDA.TableMappings.Add("Table", "Stock")


        datagridStock.DataSource = stockDS.Tables(0)
        sqlCon.Close()
        datagridStock.Expand(-1)

        stockDS.Tables(0).Clear()
        sqlCon.Close()
        Exit Sub
        '  End Try
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub txtInvoiceNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.TextChanged

    End Sub

    Private Sub TxtLblQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtLblQty.TextChanged

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub ListViewPart_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewPart.SelectedIndexChanged

    End Sub

    Private Sub txtInvoiceNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoiceNo.KeyPress

        If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then

            Dim allowedChars As String = "0123456789" & Chr(Keys.Back)


            If allowedChars.IndexOf(e.KeyChar) = -1 Then
                ' Invalid Character
                e.Handled = True
            End If

        End If




    End Sub

    Private Sub txtInvoiceNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceNo.KeyDown
        btnshipprint.Enabled = False


        ''Dim listviewpart1 As New ListView
        'Dim strCode As String
        'Dim cnSQL As SqlConnection
        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        'Dim strSQL As String
        'Dim I As Integer = 0


        'If e.KeyValue = 112 Then
        '    If TxtSelectReport.Text = Trim("D)KIT Child Items Label") Then
        '        ListViewPart.Visible = True
        '        ListViewPart.Left = 280
        '        ListViewPart.Width = 280
        '        ListViewPart.Height = 264
        '    End If
        'End If

        'strSQL = "Select ItemNumber,ItemUM from FS_Item"

        'cnSQL = New SqlConnection(ConnectionString)
        'cnSQL.Open()
        'cmSQL = New SqlCommand(strSQL, cnSQL)
        'drSQL = cmSQL.ExecuteReader()

        'Do While drSQL.Read()

        '    ListViewPart.Items.Add(drSQL.Item(0))
        '    ListViewPart.Items(I).SubItems.Add(drSQL.Item(1))
        '    I = I + 1

        'Loop

    End Sub

    Private Sub txtInvoiceNo_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtInvoiceNo.MouseDown


    End Sub

    Private Sub TxtCustDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustDesc.TextChanged

    End Sub

    Private Sub TxtLblQty_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLblQty.EnabledChanged

    End Sub

    Private Sub TxtDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDescription.TextChanged

    End Sub

    Private Sub TxtSelectReport_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSelectReport.GotFocus

    End Sub

    Private Sub TxtInfo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        btnPrint.Enabled = True
        btnPrint.Focus()
        Exit Sub
    End Sub

    Private Sub TxtLineNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtLineNo.TextChanged

    End Sub

    Private Sub TxtLineNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLineNo.Leave
        'If RBFACT.Checked = True And RBWH.Checked = True Then
        '    MsgBox("Please select Warehouse or Factory not both", MsgBoxStyle.Information)
        '    Exit Sub
        'ElseIf RBFACT.Checked = False And RBWH.Checked = False Then
        '    MsgBox("Please select Warehouse or Factory", MsgBoxStyle.Information)
        '    Exit Sub
        'End If
    End Sub

    Private Sub RBFACT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If RBWH.Checked = True Then
        '    RBWH.Checked = False
        'End If
    End Sub

    Private Sub RBWH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If RBFACT.Checked = True Then
        '    RBFACT.Checked = False
        'End If
    End Sub

    Private Sub TxtLineNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLineNo.GotFocus

    End Sub

    Private Sub TxtLineNo_ModifiedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLineNo.ModifiedChanged

    End Sub

    Private Sub txtCustomer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomer.TextChanged

    End Sub

    Private Sub TxtInfo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TxtInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub TxtLblQty_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles TxtLblQty.Layout

    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub RdbPORV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdbPORV.CheckedChanged
        Groupporv.Enabled = True
        TXTPartSelect.Text = "%"
        groupShip.Enabled = False
        txtPonumber.Text = ""
        porvclear()


    End Sub

    Private Sub RdbPORV_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RdbPORV.Click
        Groupporv.Visible = True
        'LblLabelType.Visible = False
        'RDBAutomatic.Visible = False
        'RDBManual.Visible = False
        'CheckFS.Visible = False


    End Sub

    Private Sub RdbShip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdbShip.CheckedChanged
        'Groupporv.Visible = False
        groupShip.Enabled = True
        Groupporv.Enabled = False
        txtInvoiceNo.Text = ""
        shipclear()




    End Sub

    Private Sub RdbShip_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RdbShip.Click

    End Sub

    Private Sub ComboPart_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtPonumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtPonumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs)

        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        'Dim strSQL As String


        ''    Dim strCode As String
        'Dim cnSQL As SqlConnection




        'Try

        '    '            strSQL = "Select *  from TSS_PORV_LABLES where PONumber = ' & txtPonumber.txt & '"
        '    strSQL = "Select *  from _NoLock_FS_HistoryPOReceipt where PONumber = '" & txtPonumber.Text & "'"


        '    cnSQL = New SqlConnection(ConnectionString)
        '    cnSQL.Open()

        '    cmSQL = New SqlCommand(strSQL, cnSQL)
        '    drSQL = cmSQL.ExecuteReader()

        '    If drSQL.Read() Then

        '                FillDataGrid()

        '        Exit Sub

        '    Else
        '        MsgBox("Wrong data entered! Check the Customer Order number! ", MsgBoxStyle.Exclamation, "Error!")
        '        'ClearAll()
        '    End If

        '    drSQL.Close()
        '    cnSQL.Close()
        '    cmSQL.Dispose()
        '    cnSQL.Dispose()
        'Catch
        '    MsgBox("Wrong data entered! Check the Purchase Order number! ", MsgBoxStyle.Exclamation, "Error!")
        '    ClearAll()
        'End Try




    End Sub

    Private Sub Groupporv_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Groupporv.Enter

    End Sub

    Private Sub CmbLotNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbLotNo.SelectedIndexChanged

    End Sub

    Private Sub ComboPart_GiveFeedback(ByVal sender As System.Object, ByVal e As System.Windows.Forms.GiveFeedbackEventArgs)

    End Sub

    Private Sub ComboPart_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        'Dim strSQL As String
        'Dim cnSQL As SqlConnection
        'Dim ComboPart As ComboBox


        'Try
        'strSQL = "Select HistoryPOReceiptKey,ItemNumber  from TSS_PORV_LABLES where PONumber = '" & txtPonumber.Text & "'"

        'cnSQL = New SqlConnection(ConnectionString)
        'cnSQL.Open()

        'cmSQL = New SqlCommand(strSQL, cnSQL)
        'drSQL = cmSQL.ExecuteReader()

        'If drSQL.Read() Then

        'Do While Not cmSQL.eof
        'ListBox1.Items.Add(Rs(0))


        'Rs.MoveNext()
        'Loop


        '              Dim conn As New SqlConnection(connString)

        '             Dim strSQL As String = "SELECT * FROM Disk"
        'Dim da As New SqlDataAdapter(strSQL, cnSQL)
        'Dim dt As New DataTable
        'dt.BeginLoadData()
        'dt.Fill(da, "Disk")
        'With ComboPart
        '.DataSource = dt
        'ComboPart.DisplayMember = "ItemNumber"
        'ComboPart.ValueMember = "1"


        '.DisplayMember = "ItemNumber"
        '.ValueMember = "HistoryPOReceiptKey"
        '.SelectedIndex = 0
        'End With




        '                ComboPart.Focus()
        'Exit Sub

        'Else
        'MsgBox("No Data! ", MsgBoxStyle.Exclamation, "Error!")
        'MsgBox(Err.Description)


        'Exit Sub
        'ClearAll()
        'End If

        'drSQL.Close()
        'cnSQL.Close()
        'cmSQL.Dispose()
        'cnSQL.Dispose()
        'Catch
        'MsgBox("Wrong data entered! Check the Purchase Order number! ", MsgBoxStyle.Exclamation, "Error!")
        'MsgBox(Err.Description)


        'End Try



        ' If rs.RecordCount > 0 Then
        'rs.MoveFirst()
        'For i As Integer = 0 To rs.RecordCount - 1
        'cmbEmpName.items.Add(rs.Fields("EmployeeName").val ue)
        '       rs.MoveNext()
        '      Next
        '     End If






        'Do While Not Rs.EOF
        'ListBox1.Items.Add(Rs(0))
        'Rs.MoveNext()
        'Loop






    End Sub

    Private Sub datagridStock_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles datagridStock.Navigate

    End Sub

    Sub FillDataGrid()

        ' Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        Dim strSQL As String
        Dim cnSQL As SqlConnection

        'Dim FDATE As Date
        'Dim TDATE As Date
        'dtpFROM.Format = DateTimePickerFormat.Custom
        'dtpFROM.CustomFormat = "dd/MM/yyyy"
        'dtpTo.Format = DateTimePickerFormat.Custom
        'dtpTo.CustomFormat = "dd/MM/yyyy"
        'DateTimePicker1.Format = DateTimePickerFormat.Custom
        'DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        'dtpFROM.Value = Format(dtpFROM.Value, "mm / dd / yyyy")
        'dtpTo.Value = Format(dtpTo.Value, "mm /dd / yyyy")
        'dtpFROM.Value = Format(dtpFROM.Value, "dd / mmm / yyyy")
        'dtpTo.Value = Format(dtpTo.Value, "dd / mmm / yyyy")
        'ffdate = Convert(varchar, dtpFROM.Value, 6)
        ' Convert(varchar, getdate(), 106)

        Try


            If RdbPORV.Checked = True Then


                If RBPOwise.Checked = True Then
                    strSQL = "Select HistoryPOReceiptKey,ItemNumber,LotNumber,PORV_QTY,VendorLotNumber,Stockroom1,Bin1,TransactionDate,ItemDescription from TSS_PORV_LABLES where PONumber = '" & txtPonumber.Text & "' AND ItemNumber LIKE '" & TXTPartSelect.Text & "'"

                ElseIf RBDateWise.Checked = True Then

                    strSQL = "Select HistoryPOReceiptKey,ItemNumber,LotNumber,PORV_QTY,VendorLotNumber,Stockroom1,Bin1,TransactionDate,ItemDescription from TSS_PORV_LABLES where POReceiptDate >= '" & dtpFROM.Value & "' AND POReceiptDate <='" & dtpTo.Value & "' "


                End If


            ElseIf RBPORVGoel.Checked = True Then


                If RBPOwise.Checked = True Then
                    strSQL = "Select [HistoryPOReceiptKey],[POReceiptDate], [PONumber],[POLineNumber],[ItemNumber],[Stockroom1] ,[Bin1],[PORV_QTY] ,[LotNumber] ,[BoxQty]  ,[NoofBoxs] ,[ManufacturingDate] ,[VendorID]   ,[VendorName], " & _
                             "[AvailabilityDays] ,[ShelfLifeDays]      ,[RetestDays]  FROM [FSDBBR].[dbo].[TSS_PORV_LABLES _GOEL] where PONumber = '" & txtPonumber.Text & "' AND ItemNumber LIKE '" & TXTPartSelect.Text & "'"

                ElseIf RBDateWise.Checked = True Then

                    strSQL = "Select [HistoryPOReceiptKey],[POReceiptDate], [PONumber],[POLineNumber],[ItemNumber],[Stockroom1] ,[Bin1],[PORV_QTY] ,[LotNumber] ,[BoxQty]  ,[NoofBoxs] ,[ManufacturingDate] ,[VendorID]   ,[VendorName], " & _
                             "[ITestDate] ,[2TestDate],[3TestDate] FROM [FSDBBR].[dbo].[TSS_PORV_LABLES _GOEL] where[Stockroom1]  like 'EL' and Transdate >= '" & dtpFROM.Value & "' AND Transdate <='" & dtpTo.Value & "' AND ItemNumber LIKE '" & TXTPartSelect.Text & "'"
                End If


            ElseIf RBPorvIMTRGoel.Checked = True Then

                'QUERY NEED TO BE CHANGED ON 16TH SEPT 2016

                If RBPOwise.Checked = True Then
                    strSQL = "Select  [HistoryPOReceiptKey], [POReceiptDate], [PONumber],[POLineNumber],[ItemNumber],[StockroomTo] ,[BinTo],[InventoryQuantity] ,[LotNumber] ,[BOXQTY]  ,[NOOFBOXES] ,[ManufacturingDate] ,[VendorID]   ,[VendorName] , " & _
                             "ITestDate, [2TestDate] ,[3TestDate]  FROM [FSDBBR].[dbo].[TSS_PORV_LABLES _GOEL_IMTR]  where PONumber = '" & txtPonumber.Text & "' AND ItemNumber LIKE '" & TXTPartSelect.Text & "'"

                ElseIf RBDateWise.Checked = True Then

                    strSQL = "Select  [HistoryPOReceiptKey], [POReceiptDate], [PONumber],[POLineNumber],[ItemNumber],[StockroomTo] ,[BinTo],[InventoryQuantity] ,[LotNumber] ,[BOXQTY]  ,[NOOFBOXES] ,[ManufacturingDate] ,[VendorID]   ,[VendorName] , " & _
                             "ITestDate, [2TestDate] ,[3TestDate]  FROM [FSDBBR].[dbo].[TSS_PORV_LABLES _GOEL_IMTR] where  [DateTime]  >= '" & dtpFROM.Value & "' AND  [DateTime]  <='" & dtpTo.Value & "' AND ItemNumber LIKE '" & TXTPartSelect.Text & "'"
                End If
            End If

            cnSQL = New SqlConnection(ConnectionString)
            'cnSQL.Open()

            Dim stockDS As DataSet = New DataSet
            'Try
            'Catch ex As Exception
            'End Try
            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDA As SqlDataAdapter = New SqlDataAdapter
            'Try
            stockDA.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDA.TableMappings.Add("Table", "Stock")
            'get data
            stockDA.Fill(stockDS)

            If RdbPORV.Checked = True Then


                stockDS.Tables(0).Columns(0).ColumnName = "Key"
                stockDS.Tables(0).Columns(1).ColumnName = "ItemNumber"
                stockDS.Tables(0).Columns(2).ColumnName = "LotNumber"
                stockDS.Tables(0).Columns(3).ColumnName = "LotQty"
                stockDS.Tables(0).Columns(4).ColumnName = "CureDate"
                stockDS.Tables(0).Columns(5).ColumnName = "StockRoom1"
                stockDS.Tables(0).Columns(6).ColumnName = "Bin1"
                stockDS.Tables(0).Columns(7).ColumnName = "LotDate"
                stockDS.Tables(0).Columns(8).ColumnName = "ItemDescription"

            ElseIf RBPORVGoel.Checked = True Or RBPorvIMTRGoel.Checked = True Then

                stockDS.Tables(0).Columns(0).ColumnName = "Key"
                stockDS.Tables(0).Columns(1).ColumnName = " PORV Dt"
                stockDS.Tables(0).Columns(2).ColumnName = "PONumber"
                stockDS.Tables(0).Columns(3).ColumnName = "POLn"
                stockDS.Tables(0).Columns(4).ColumnName = "ItemNumber"
                stockDS.Tables(0).Columns(5).ColumnName = "StockRoom1"
                stockDS.Tables(0).Columns(6).ColumnName = "Bin1"
                stockDS.Tables(0).Columns(7).ColumnName = "Qty"
                stockDS.Tables(0).Columns(8).ColumnName = "LotNumber"
                stockDS.Tables(0).Columns(9).ColumnName = "BoxQty"
                stockDS.Tables(0).Columns(10).ColumnName = "No.Of Boxes"
                stockDS.Tables(0).Columns(11).ColumnName = "MfgDate"
                stockDS.Tables(0).Columns(12).ColumnName = "VendorID"
                stockDS.Tables(0).Columns(13).ColumnName = "Name"
                stockDS.Tables(0).Columns(14).ColumnName = "I Test Dt"
                stockDS.Tables(0).Columns(15).ColumnName = "II Test Dt"
                stockDS.Tables(0).Columns(16).ColumnName = "Expiry Dt"
            End If


            datagridStock.DataSource = stockDS.Tables(0)
            cnSQL.Close()
            datagridStock.Expand(-1)
            MsgBox("Select Key for label print", MsgBoxStyle.Information, "PORV-LABELS")
            Exit Sub

        Catch
            'MsgBox("Error occured in filldata!!")
            'btnPrint.Enabled = False
            MsgBox(Err.Description)

            '   cnSQL.Close()
            Exit Sub
        End Try
    End Sub


    Private Sub datagridStock_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridStock.CurrentCellChanged

        Dim a As Integer
        Dim b As Integer
        Dim cnSQL As SqlConnection
        Dim strsql As String

        b = datagridStock.CurrentCell.ColumnNumber
        TextlblqtyPORV.Text = ""
        txtNoofLabels.Text = ""
        If b = 0 Then
            a = datagridStock.Item(datagridStock.CurrentCell)
            '    a = "2525865"

            cnSQL = New SqlConnection(ConnectionString)
            Dim cmSQL As SqlCommand
            Dim drSQL As SqlDataReader

            '   strsql = "Select HistoryPOReceiptKey,ItemNumber,LotNumber,PORV_QTY,VendorLotNumber,Stockroom1,Bin1,TransactionDate,ItemDescription from TSS_PORV_LABLES where HistoryPOReceiptKey = '" & a & "'"

            strsql = "Select [HistoryPOReceiptKey],[POReceiptDate], [ItemNumber],[Stockroom1] ,[Bin1],[PORV_QTY] ,[LotNumber] ,[BoxQty]  ,[NoofBoxs] ,[ManufacturingDate] ,[VendorID],[VendorName], " & _
                            "[ITestDate] ,[2TestDate],[3TestDate] FROM [FSDBBR].[dbo].[TSS_PORV_LABLES _GOEL] where HistoryPOReceiptKey = '" & a & "'"

            cnSQL.Open()

            cmSQL = New SqlCommand(strsql, cnSQL)
            drSQL = cmSQL.ExecuteReader()
            If drSQL.Read() Then

                TXTKEY.Text = drSQL.Item(0)
                txtLotDate.Text = drSQL.Item(1) 'porv date, transactiondate
                txtPart.Text = drSQL.Item(2)
                txtStkRoom1.Text = drSQL.Item(3)
                txtBin1.Text = drSQL.Item(4)
                TXTLOTQTYPORV.Text = drSQL.Item(5)
                txtLot.Text = drSQL.Item(6)


                If IsDBNull(drSQL.Item(7)) Then

                    TextlblqtyPORV.Text = 0
                Else

                    TextlblqtyPORV.Text = drSQL.Item(7)
                End If

                If IsDBNull(drSQL.Item(8)) Then

                    txtNoofLabels.Text = 0
                Else

                    txtNoofLabels.Text = drSQL.Item(8)
                End If

                TXTVENDCURE.Text = drSQL.Item(9) 'Mfg date
                txtItemDesc.Text = drSQL.Item(11) 'vendor name

                txtdate1.Text = drSQL.Item(12)
                txtdate2.Text = drSQL.Item(13)
                txtdate3.Text = drSQL.Item(14)

                'If IsDBNull(drSQL.Item(4)) Then
                'TXTVENDCURE.Text = ""
                'Else

                'TXTVENDCURE.Text = drSQL.Item(4)
                'End If
                TextlblqtyPORV.Focus()
            End If

        Else
            MsgBox("Pl click on Key", vbInformation)
            'MsgBox("CurrentCellChanged event")
            Exit Sub


        End If

    End Sub

    Private Sub datagridStock_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridStock.Resize

    End Sub

    Private Sub Textlblqty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextlblqtyPORV.TextChanged

    End Sub

    Private Sub TextlblqtyPORV_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextlblqtyPORV.Leave

        If Val(TextlblqtyPORV.Text) >= 1 Then

            Dim C As Integer
            Dim D As Integer

            'C = TORight(InStr(txtNoofLabels.Text, "."), 1)
            'C = Right(txtNoofLabels.Text, 
            'txtNoofLabels.Text = Trim(txtNoofLabels.Text)

            'If (txtNoofLabels.Text.Length) = 1 Then
            'C = (txtNoofLabels.Text.Substring(txtNoofLabels.Text.Length - 1, 1))
            'ElseIf (txtNoofLabels.Text.Length) = 2 Then
            '   C = (txtNoofLabels.Text.Substring(txtNoofLabels.Text.Length - 2, 2))

            'ElseIf (txtNoofLabels.Text.Length) = 3 Then
            '    C = (txtNoofLabels.Text.Substring(txtNoofLabels.Text.Length - 3, 3))

            C = Val(TXTLOTQTYPORV.Text) Mod Val(TextlblqtyPORV.Text)

            D = Val(TXTLOTQTYPORV.Text) - Val(C)
            txtNoofLabels.Text = Round(D / Val(TextlblqtyPORV.Text), 0)


            If C > 0 Then
                txtNoofLabels.Text = txtNoofLabels.Text + 1

            End If

        End If


    End Sub

    Private Sub datagridStock_RightToLeftChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridStock.RightToLeftChanged

    End Sub

    Private Sub datagridStock_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridStock.Click


    End Sub

    Private Sub datagridStock_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridStock.DataSourceChanged

    End Sub

    Private Sub txtNoofLabels_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNoofLabels.TextChanged

    End Sub

    Private Sub btnPorvLblPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPorvLblPrint.Click
        If RdbPORV.Checked = True Then


            Dim report As String
            report = Microsoft.VisualBasic.Left(TxtSelectReport.Text, 1)

            Dim rptporvLabel As New ReportDocument

            Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection2 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection3 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection4 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection5 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection6 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection7 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection8 As New CrystalDecisions.Shared.ParameterValues

            Dim pvCollection9 As New CrystalDecisions.Shared.ParameterValues 'item barcode
            Dim pvCollection10 As New CrystalDecisions.Shared.ParameterValues 'for stock room
            Dim pvcollection11 As New CrystalDecisions.Shared.ParameterValues
            Dim pvcollection12 As New CrystalDecisions.Shared.ParameterValues

            Dim pdvPartNo As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvPartDesc As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvLotNumber As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvLotDate As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvcureDate As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvStkroom As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvBin As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvqty As New CrystalDecisions.Shared.ParameterDiscreteValue

            Dim pdvItemBarcode As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvStockRoomBarcode As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvlotnumberbarcode As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim Pdvbarcodeqty As New CrystalDecisions.Shared.ParameterDiscreteValue



            'Dim path As String
            'path = System.IO.Path.GetDirectoryName( _
            ' System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
            'MessageBox.Show(path)


            ' rptporvLabel.Load("rptporvLabel.rpt")

            pdvPartNo.Value = Trim(txtPart.Text)
            pdvPartDesc.Value = Trim(txtItemDesc.Text)
            pdvLotNumber.Value = Trim(txtLot.Text)
            txtLotDate.Text = Format(txtLotDate.Text, "dd/mmm/yyyy")
            pdvLotDate.Value = (txtLotDate.Text)
            pdvcureDate.Value = Trim(TXTVENDCURE.Text)
            pdvStkroom.Value = Trim(txtStkRoom1.Text)
            pdvBin.Value = Trim(txtBin1.Text)
            pdvqty.Value = Trim(TextlblqtyPORV.Text)

            pdvItemBarcode.Value = RTrim(LTrim(txtPart.Text))
            pdvlotnumberbarcode.Value = RTrim(LTrim(txtLot.Text)) '+ Trim(TextlblqtyPORV.Text)
            pdvStockRoomBarcode.Value = RTrim(LTrim(txtStkRoom1.Text)) + RTrim(LTrim(txtBin1.Text))
            Pdvbarcodeqty.Value = Trim(TextlblqtyPORV.Text)
            ' Add it to the parameter collection.
            pvCollection.Add(pdvPartNo)
            pvCollection2.Add(pdvPartDesc)
            pvCollection3.Add(pdvLotNumber)
            pvCollection4.Add(pdvLotDate)
            pvCollection5.Add(pdvcureDate)
            pvCollection6.Add(pdvStkroom)
            pvCollection7.Add(pdvBin)
            pvCollection8.Add(pdvqty)
            pvCollection9.Add(pdvItemBarcode)
            pvCollection10.Add(pdvStockRoomBarcode)
            pvcollection11.Add(pdvlotnumberbarcode)
            pvcollection12.Add(Pdvbarcodeqty)

            ' Apply the current parameter values.
            rptporvLabel.DataDefinition.ParameterFields("PartNo").ApplyCurrentValues(pvCollection)
            rptporvLabel.DataDefinition.ParameterFields("ItemDescription").ApplyCurrentValues(pvCollection2)
            rptporvLabel.DataDefinition.ParameterFields("LotNumber").ApplyCurrentValues(pvCollection3)
            rptporvLabel.DataDefinition.ParameterFields("LotDate").ApplyCurrentValues(pvCollection4)
            rptporvLabel.DataDefinition.ParameterFields("CureDate").ApplyCurrentValues(pvCollection5)
            rptporvLabel.DataDefinition.ParameterFields("Stockroom1").ApplyCurrentValues(pvCollection6)
            rptporvLabel.DataDefinition.ParameterFields("Bin1").ApplyCurrentValues(pvCollection7)
            rptporvLabel.DataDefinition.ParameterFields("Qty").ApplyCurrentValues(pvCollection8)

            rptporvLabel.DataDefinition.ParameterFields("ItemBarcode").ApplyCurrentValues(pvCollection9)
            rptporvLabel.DataDefinition.ParameterFields("StkroomBarcode").ApplyCurrentValues(pvCollection10)
            rptporvLabel.DataDefinition.ParameterFields("LotBarcode").ApplyCurrentValues(pvcollection11)
            rptporvLabel.DataDefinition.ParameterFields("QtyBarcode").ApplyCurrentValues(pvcollection12)

            Dim CTR As Integer
            Dim LABELCOUNT As Integer
            LABELCOUNT = Val(txtNoofLabels.Text)
            Dim lastlabelqty As Integer

            lastlabelqty = Val(TXTLOTQTYPORV.Text) - ((Val(txtNoofLabels.Text) - 1) * Val(TextlblqtyPORV.Text))
            'pdvItemBarcode.Value = RTrim(LTrim(txtPart.Text)) + RTrim(LTrim(txtLot.Text))
            'pdvlotnumberbarcode.Value = RTrim(LTrim(txtLot.Text)) + Trim(TextlblqtyPORV.Text)


            '+ Trim(TextlblqtyPORV.Text)


            CTR = 1

            Do While CTR <= LABELCOUNT
                If CTR = LABELCOUNT Then

                    pdvqty.Value = (lastlabelqty)
                    pvCollection8.Add(pdvqty)
                    rptporvLabel.DataDefinition.ParameterFields("Qty").ApplyCurrentValues(pvCollection8)

                    pdvlotnumberbarcode.Value = RTrim(LTrim(txtLot.Text)) + Trim(lastlabelqty)


                End If
                rptporvLabel.Load("rptporvLabel.rpt")
                rptporvLabel.PrintToPrinter(1, False, 0, 0)
                'rptporvLabel.PrintToPrinter(1, False, 1, 1)


                'If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
                'r'ptShipLabel.Load("Label2.rpt")

                'ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
                '   rptShipLabel.Load("Label2Invoice.rpt")

                'End If




                'MsgBox("pdvqty.value")

                CTR = CTR + 1
            Loop

            pvCollection.Clear()
            pvCollection2.Clear()
            pvCollection3.Clear()
            pvCollection4.Clear()
            pvCollection5.Clear()
            pvCollection6.Clear()
            pvCollection7.Clear()
            pvCollection8.Clear()
            pvCollection9.Clear()
            pvCollection10.Clear()
            pvcollection11.Clear()
            pvcollection12.Clear()

            rptporvLabel.Close()

        ElseIf RBPORVGoel.Checked = True Or RBPorvIMTRGoel.Checked = True Then

            Dim report As String
            report = Microsoft.VisualBasic.Left(TxtSelectReport.Text, 1)

            Dim rptporvLabel As New ReportDocument

            Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection1 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection2 As New CrystalDecisions.Shared.ParameterValues


            Dim pdvKey As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvQtyperLabel As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim PdvNoOfBox As New CrystalDecisions.Shared.ParameterDiscreteValue



            'Dim path As String
            'path = System.IO.Path.GetDirectoryName( _
            ' System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
            'MessageBox.Show(path)

            If RBPORVGoel.Checked = True Then
                rptporvLabel.Load("rptporvLabelGoel.rpt")
            ElseIf RBPorvIMTRGoel.Checked = True Then
                rptporvLabel.Load("rptporvLabelGoelIMTR.rpt")

            End If

            pdvKey.Value = Trim(TXTKEY.Text)
            pdvQtyperLabel.Value = Trim(TextlblqtyPORV.Text)
            PdvNoOfBox.Value = Trim(txtNoofLabels.Text)
            'pdvPartDesc.Value = Trim(txtItemDesc.Text)
            'pdvLotNumber.Value = Trim(txtLot.Text)
            'txtLotDate.Text = Format(txtLotDate.Text, "dd/mmm/yyyy")
            'pdvLotDate.Value = (txtLotDate.Text)
            'pdvcureDate.Value = Trim(TXTVENDCURE.Text)
            'pdvStkroom.Value = Trim(txtStkRoom1.Text)
            'pdvBin.Value = Trim(txtBin1.Text)
            '  pdvQtyperLabel.Value = Trim(TextlblqtyPORV.Text)

            'pdvItemBarcode.Value = RTrim(LTrim(txtPart.Text))
            'pdvlotnumberbarcode.Value = RTrim(LTrim(txtLot.Text)) '+ Trim(TextlblqtyPORV.Text)
            'pdvStockRoomBarcode.Value = RTrim(LTrim(txtStkRoom1.Text)) + RTrim(LTrim(txtBin1.Text))
            'Pdvbarcodeqty.Value = Trim(TextlblqtyPORV.Text)


            ' Add it to the parameter collection.
            pvCollection.Add(pdvKey)
            pvCollection1.Add(pdvQtyperLabel)
            pvCollection2.Add(PdvNoOfBox)
            'pvCollection4.Add(pdvLotDate)
            'pvCollection5.Add(pdvcureDate)
            'pvCollection6.Add(pdvStkroom)
            'pvCollection7.Add(pdvBin)
            'pvCollection8.Add(pdvqty)
            'pvCollection9.Add(pdvItemBarcode)
            'pvCollection10.Add(pdvStockRoomBarcode)
            'pvcollection11.Add(pdvlotnumberbarcode)
            'pvcollection12.Add(Pdvbarcodeqty)

            ' Apply the current parameter values.

            rptporvLabel.DataDefinition.ParameterFields("porvkey").ApplyCurrentValues(pvCollection)
            rptporvLabel.DataDefinition.ParameterFields("qty").ApplyCurrentValues(pvCollection1)
            rptporvLabel.DataDefinition.ParameterFields("NoOfBox").ApplyCurrentValues(pvCollection2)
            'rptporvLabel.DataDefinition.ParameterFields("LotNumber").ApplyCurrentValues(pvCollection3)
            'rptporvLabel.DataDefinition.ParameterFields("LotDate").ApplyCurrentValues(pvCollection4)
            'rptporvLabel.DataDefinition.ParameterFields("CureDate").ApplyCurrentValues(pvCollection5)
            'rptporvLabel.DataDefinition.ParameterFields("Stockroom1").ApplyCurrentValues(pvCollection6)
            'rptporvLabel.DataDefinition.ParameterFields("Bin1").ApplyCurrentValues(pvCollection7)
            'rptporvLabel.DataDefinition.ParameterFields("Qty").ApplyCurrentValues(pvCollection8)

            'rptporvLabel.DataDefinition.ParameterFields("ItemBarcode").ApplyCurrentValues(pvCollection9)
            'rptporvLabel.DataDefinition.ParameterFields("StkroomBarcode").ApplyCurrentValues(pvCollection10)
            'rptporvLabel.DataDefinition.ParameterFields("LotBarcode").ApplyCurrentValues(pvcollection11)
            'rptporvLabel.DataDefinition.ParameterFields("QtyBarcode").ApplyCurrentValues(pvcollection12)

            Dim CTR As Integer
            Dim LABELCOUNT As Integer
            LABELCOUNT = Val(txtNoofLabels.Text)
            Dim lastlabelqty As Integer

            lastlabelqty = Val(TXTLOTQTYPORV.Text) - ((Val(txtNoofLabels.Text) - 1) * Val(TextlblqtyPORV.Text))
            'pdvItemBarcode.Value = RTrim(LTrim(txtPart.Text)) + RTrim(LTrim(txtLot.Text))
            'pdvlotnumberbarcode.Value = RTrim(LTrim(txtLot.Text)) + Trim(TextlblqtyPORV.Text)


            '+ Trim(TextlblqtyPORV.Text)


            CTR = 1

            Do While CTR <= LABELCOUNT
                If CTR = LABELCOUNT Then

                    pdvQtyperLabel.Value = (lastlabelqty)
                    pvCollection1.Add(pdvQtyperLabel)
                    rptporvLabel.DataDefinition.ParameterFields("Qty").ApplyCurrentValues(pvCollection1)
                    '    rptporvLabel.DataDefinition.ParameterFields("qty").ApplyCurrentValues(pvCollection1)
                    '     pdvlotnumberbarcode.Value = RTrim(LTrim(txtLot.Text)) + Trim(lastlabelqty)


                End If

                rptporvLabel.PrintToPrinter(1, False, 0, 0)
                'rptporvLabel.PrintToPrinter(1, False, 1, 1)


                'If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
                'r'ptShipLabel.Load("Label2.rpt")

                'ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
                '   rptShipLabel.Load("Label2Invoice.rpt")

                'End If

                'MsgBox("pdvqty.value")

                CTR = CTR + 1
            Loop

            pvCollection.Clear()
            pvCollection1.Clear()
            pvCollection2.Clear()
            'pvCollection4.Clear()
            'pvCollection5.Clear()
            'pvCollection6.Clear()
            'pvCollection7.Clear()
            'pvCollection8.Clear()
            'pvCollection9.Clear()
            'pvCollection10.Clear()
            'pvcollection11.Clear()
            'pvcollection12.Clear()

            rptporvLabel.Close()

        End If

    End Sub

    Private Sub btnporvok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


        'If RdbPORV.Checked = True Then
        '    datagridStock.Visible = True

        '    datagridStock.Enabled = True

        '    FillDataGrid()
        'Else
        '    MsgBox("Select PORV before clicking ok", vbInformation, "Label Software")
        '    Exit Sub
        'End If
        'porvclear()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BtnShipOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShipOK.Click

        Dim report As String

        shipclear()
        btnshipprint.Enabled = True


        Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String




        If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
            report = "A"
            shipenable()

        ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
            report = "B"
            shipenable()

        ElseIf ComboBox1.Text = Trim("C)Box Label") Then
            report = "C"
            shipdisable()
        ElseIf ComboBox1.Text = Trim("D)KIT Child Items List") Then
            report = "D"
            MsgBox("Program will be developed after BILL modification", vbInformation)
            Exit Sub
        End If


        If txtInvoiceNo.Text.Length = 0 And report = "A" Then
            MsgBox("Please, enter Shipment no.!", MsgBoxStyle.Critical, "Error!")
            txtInvoiceNo.Focus()
            Exit Sub
        End If

        If txtInvoiceNo.Text.Length = 0 And report = "B" Then
            MsgBox("Please, enter Invoice No.!", MsgBoxStyle.Critical, "Error!")
            txtInvoiceNo.Focus()
            Exit Sub
        End If


        If report = "B" Then


            strSQL = "Select LotNumber,CONumber,ItemNumber,COLineNumber,ShipmentNumber, 0 as qtyPerLabel,ShippedQuantity, CustomerID, InvoiceNumber,ItemUM, ShipmentDate " & _
                        "from TSS_ItemLabel_Ver2 " & _
                        "where InvoiceNumber  = '" & txtInvoiceNo.Text & "' and " & _
                        "COLineNumber like '" & TxtLineNo.Text & "' order by COLineNumber "



        ElseIf report = "A" Then

            ' strSQL = "Select LotNumber,CONumber,ItemNumber,COLineNumber,ShippedQuantity,ItemDescription,CustItemNumber," & _
            '  "ItemUM, CustomerName, CustomerPONumber  " & _
            ' "from TSS_ItemLabel2_New " & _
            ' "where ShipmentNumber  = " & txtInvoiceNo.Text & " and " & _
            ' "COLineNumber like '" & TxtLineNo.Text & "' order by COLineNumber "


            strSQL = "Select LotNumber,CONumber,ItemNumber,COLineNumber,ShipmentNumber, 0 as qtyPerLabel,ShippedQuantity, CustomerID, 0 as InvoiceNumber,ItemUM,ShipmentDate " & _
             "from TSS_History_Shipment " & _
             "where ShippedQuantity > 0 and ShipmentNumber  = " & txtInvoiceNo.Text & " and " & _
             "COLineNumber like '" & TxtLineNo.Text & "' order by COLineNumber "

        ElseIf report = "C" Then


            strSQL = "Select  ShipmentNumber,InvoiceDate,CustomerPONumber, ConsigneeName,ShipmentAddress1," & _
                            "ShipmentAddress2,ShipmentCity, ShipmentState,ShipmentZip," & _
                            "CustomerContact, CustomerContactPhone, cod, BuyerName,Designation,Dept,Phone,Mobile,CustomerName from TSS_BOXLabel_DATA_WH_VER2 where " & _
                            "InvoiceNumber = '" & txtInvoiceNo.Text & "' "

        End If


        If report = "C" Then

            cnSQL = New SqlConnection(ConnectionString)
            cnSQL.Open()

            cmSQL = New SqlCommand(strSQL, cnSQL)
            drSQL = cmSQL.ExecuteReader()
            If drSQL.Read() Then



                TxtCustOrdNo.Text = drSQL.Item(0) 'shipmentnumber
                txtQty.Text = drSQL.Item(1) 'invoicedate 
                TxtCustPO.Text = drSQL.Item(2) 'custpo  
                txtItem.Text = drSQL.Item(3) 'cust name
                TxtDescription.Text = drSQL.Item(4) 'ship ad1
                TxtCustItem.Text = drSQL.Item(5) 'ship ad2
                TxtCustDesc.Text = drSQL.Item(6) 'shipcity
                txtLotNo.Text = drSQL.Item(7)   'shipstate
                TxtUoM.Text = drSQL.Item(8) 'shipzip
                txtCustomer.Text = drSQL.Item(9) 'customercontact
                TxtCOLnNo.Text = drSQL.Item(10) 'customercontactphone.


                txtbuyer.Text = drSQL.Item(12) 'buyer
                txtdesig.Text = drSQL.Item(13) 'desig
                txtdept.Text = drSQL.Item(14) 'dept
                txtphone.Text = drSQL.Item(15) 'phone
                txtmobile.Text = drSQL.Item(16) 'mobile
                txtcusto.Text = drSQL.Item(17) 'customer



            Else
                MsgBox("Invalid Invoice Number", vbInformation)
                Exit Sub
            End If

            TxtInfo.Focus()

        End If


        If report <> "C" Then

            'FillDataGridShip()
            ''--
            cnSQL = New SqlConnection(ConnectionString)
            'cnSQL.Open()


            Dim stockDShip As DataSet = New DataSet


            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAS As SqlDataAdapter = New SqlDataAdapter
            'Try
            stockDAS.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAS.TableMappings.Add("Table", "Stock")
            'get data
            stockDAS.Fill(stockDShip)

            stockDShip.Tables(0).Columns(0).ColumnName = "LotNumber"
            stockDShip.Tables(0).Columns(1).ColumnName = "OrderNo."

            stockDShip.Tables(0).Columns(2).ColumnName = "ItemNumber"
            stockDShip.Tables(0).Columns(3).ColumnName = "CoLineNo."
            stockDShip.Tables(0).Columns(4).ColumnName = "ShipNo"

            '  stockDShip.Tables(0).Columns(5).ColumnName = "ShipDate"
            stockDShip.Tables(0).Columns(5).ColumnName = "Qty/Label"
            stockDShip.Tables(0).Columns(6).ColumnName = "ShipQty"


            stockDShip.Tables(0).Columns(7).ColumnName = "CustomerID"
            stockDShip.Tables(0).Columns(8).ColumnName = "InvoiceNum"

            stockDShip.Tables(0).Columns(9).ColumnName = "ShipDate"

            stockDShip.Tables(0).Columns(0).ReadOnly = True
            stockDShip.Tables(0).Columns(1).ReadOnly = True
            stockDShip.Tables(0).Columns(2).ReadOnly = True
            stockDShip.Tables(0).Columns(3).ReadOnly = True
            stockDShip.Tables(0).Columns(4).ReadOnly = True
            stockDShip.Tables(0).Columns(5).ReadOnly = False
            stockDShip.Tables(0).Columns(6).ReadOnly = True
            stockDShip.Tables(0).Columns(7).ReadOnly = True
            stockDShip.Tables(0).Columns(8).ReadOnly = True
            stockDShip.Tables(0).Columns(9).ReadOnly = True


            'dataGridship.Columns("LotNumber").ReadOnly = false;




            DataGridShip.DataSource = stockDShip.Tables(0)
            cnSQL.Close()
            DataGridShip.Expand(-1)
            MsgBox("Select Lot Number for label print", MsgBoxStyle.Information, "SHIP-LABELS")



            '  for each (DataGridship dgvc in dgSearchedResults.Columns)
            '         {
            '            dgvc.ReadOnly = true;
            '       }

            '    dataGridView1.Columns("ColumnName").ReadOnly = false;




        End If

        Exit Sub

        MsgBox(Err.Description)

        cnSQL.Close()
        '   End Try

        'Catch
        '    MsgBox("Wrong data entered! Check the Invoice number! ", MsgBoxStyle.Exclamation, "Error!")
        '    ClearAll()
        'End Try


    End Sub

    Private Sub shipdisable()
        TxtInfo.Enabled = True
        TxtLineNo.Enabled = False
        txtshipLine.Enabled = False
        txtshippartno.Enabled = False
        txtshiplot.Enabled = False
        txtshipqty.Enabled = False
        txtshipqpl.Enabled = False

    End Sub

    Private Sub shipenable()
        TxtInfo.Enabled = False
        TxtLineNo.Enabled = True
        txtshipLine.Enabled = True
        txtshippartno.Enabled = True
        txtshiplot.Enabled = True
        txtshipqty.Enabled = True
        txtshipqpl.Enabled = True
        TxtLineNo.Enabled = True


    End Sub
    Private Sub shipclear()
        txtshipLine.Text = ""
        txtshippartno.Text = ""
        txtshiplot.Text = ""
        txtshipqty.Text = ""
        txtshipqpl.Text = ""
        txtshiplabel.Text = ""

    End Sub
    Private Sub l_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles groupShip.Enter

    End Sub

    Private Sub txtInvoiceNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.Leave

        If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then

            TxtLineNo.Enabled = True

        ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
            TxtLineNo.Enabled = True
        Else
            TxtLineNo.Enabled = False
        End If


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Leave


        If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
            LblInvoice.Text = "Shipment No."
        ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
            LblInvoice.Text = "Invoice No."

        ElseIf ComboBox1.Text = Trim("C)Box Label") Then
            LblInvoice.Text = "Invoice No."

        ElseIf ComboBox1.Text = Trim("D)KIT Child Items List") Then
            LblInvoice.Text = "D"
        End If


    End Sub
    Sub FillDataGridShip()



    End Sub



    Private Sub DataGridShip_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridShip.Navigate

    End Sub

    Private Sub DataGridShip_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridShip.CurrentCellChanged

        Dim a As String

        Dim b As Double

        Dim cnSQL As SqlConnection
        Dim strsql As String

        Dim c As String
        Dim d As Double

        b = DataGridShip.CurrentCell.ColumnNumber

        TextlblqtyPORV.Text = ""
        txtNoofLabels.Text = ""

        If b = 0 Then

            a = DataGridShip.Item(DataGridShip.CurrentCell)
            c = DataGridShip.Item(DataGridShip.CurrentCell.RowNumber, 3)
            d = DataGridShip.Item(DataGridShip.CurrentCell.RowNumber, 6)


            cnSQL = New SqlConnection(ConnectionString)
            Dim cmSQL As SqlCommand
            Dim drSQL As SqlDataReader



            If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
                strsql = "Select LotNumber,ItemNumber,COLineNumber,ShippedQuantity,CONumber,CustomerID,ItemUM " & _
                               "from TSS_History_Shipment " & _
                              "where ShipmentNumber  = '" & txtInvoiceNo.Text & "' and " & _
                              "LotNumber like '" & a & "' and COLineNumber = '" & c & "' and ShippedQuantity = " & d & "  and ShippedQuantity > 0"



            ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then

                strsql = "Select LotNumber,ItemNumber,COLineNumber,ShippedQuantity,CONumber,CustomerID,ItemUM " & _
                                           "from TSS_ItemLabel_Ver2  " & _
                                          "where InvoiceNumber  = '" & txtInvoiceNo.Text & "' and " & _
                                          "LotNumber like '" & a & "' and COLineNumber = '" & c & "' and ShippedQuantity = " & d & " and ShippedQuantity > 0"



            End If



            cnSQL.Open()

            cmSQL = New SqlCommand(strsql, cnSQL)
            drSQL = cmSQL.ExecuteReader()
            If drSQL.Read() Then
                TxtUoM.Text = ""
                txtshipLine.Text = drSQL.Item(2)
                txtshippartno.Text = drSQL.Item(1)
                txtshiplot.Text = drSQL.Item(0)
                txtshipqty.Text = drSQL.Item(3)
                txtconumber.Text = drSQL.Item(4)
                txtcustid.Text = drSQL.Item(5)
                TxtUoM.Text = drSQL.Item(6)
                txtshipqpl.Focus()

            End If

        Else
            MsgBox("Pl click on Lot Number", vbInformation)
            'MsgBox("CurrentCellChanged event")
            Exit Sub


        End If

    End Sub

    Private Sub btnshipprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnshipprint.Click


        If ComboBox1.Text = Trim("C)Box Label") Then

            If txtshiplabel.Text = "" Then
                MsgBox("No of of labels should not be blank", vbInformation)
                Exit Sub
            End If


        End If

        Dim report As String
        Dim customername As String
        Dim custpart As String
        Dim curedate As String
        Dim custpo As String
        'Dim UOM As String


        If ComboBox1.Text <> Trim("C)Box Label") Then

            'get customername and customer part number

            txtshipinvdate.Text = DataGridShip.Item(0, 5)

            Dim cnSQL As SqlConnection
            Dim cmSQL As SqlCommand
            Dim drSQL As SqlDataReader
            Dim strSQL As String

            strSQL = "Select CustomerName,CustomerItemNumber,ItemUM from TSS_Customer_Item where CustomerID = '" & txtcustid.Text & "' and ItemNumber = '" & txtshippartno.Text & "'"

            cnSQL = New SqlConnection(ConnectionString)
            cnSQL.Open()

            cmSQL = New SqlCommand(strSQL, cnSQL)
            drSQL = cmSQL.ExecuteReader()
            If drSQL.Read() Then



                customername = drSQL.Item(0)
                custpart = drSQL.Item(1)

            Else
                customername = "-"
                custpart = "-"

            End If

            'end
            'select customername if blank

            If Len(customername) < 3 Then
                strSQL = "Select CustomerName from TSS_Customer_Item where CustomerID = '" & txtcustid.Text & "' "

                cnSQL = New SqlConnection(ConnectionString)
                cnSQL.Open()

                cmSQL = New SqlCommand(strSQL, cnSQL)
                drSQL = cmSQL.ExecuteReader()
                If drSQL.Read() Then

                    customername = drSQL.Item(0)
                Else
                    customername = "-"

                End If
            End If


            'get curedate

            strSQL = "Select VendorLotNumber from TSS_ItemLabel_CureDt where LotNumber = '" & txtshiplot.Text & "' "

            cnSQL = New SqlConnection(ConnectionString)
            cnSQL.Open()

            cmSQL = New SqlCommand(strSQL, cnSQL)
            drSQL = cmSQL.ExecuteReader()

            If drSQL.Read() Then

                curedate = drSQL.Item(0)
            Else
                curedate = "-"


            End If

            'end

            'get customer ponumber

            strSQL = "Select CustomerPONumber from FS_COHeader where CONumber = '" & txtconumber.Text & "' "

            cnSQL = New SqlConnection(ConnectionString)
            cnSQL.Open()

            cmSQL = New SqlCommand(strSQL, cnSQL)
            drSQL = cmSQL.ExecuteReader()

            If drSQL.Read() Then

                custpo = drSQL.Item(0)

            End If


        End If

        'End


        'report = Microsoft.VisualBasic.Left(TxtSelectReport.Text, 1)

        Dim pdvoption As New CrystalDecisions.Shared.ParameterDiscreteValue
        Dim rptShipLabel As New ReportDocument

        If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
            pdvoption.Value = "Shipment No."
            report = "A"
        ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
            pdvoption.Value = "Invoice No."
            report = "B"
        ElseIf ComboBox1.Text = Trim("C)Box Label") Then
            pdvoption.Value = "C"
            report = "C"
        ElseIf ComboBox1.Text = Trim("D)KIT Child Items List") Then
            pdvoption.Value = "D"
            report = "D"
        End If

        If report = "A" Or report = "B" Then
            Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection2 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection3 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection4 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection5 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection6 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection7 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection8 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection9 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection10 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection11 As New CrystalDecisions.Shared.ParameterValues




            Dim pdvShipNo As New CrystalDecisions.Shared.ParameterDiscreteValue 'SHIPNO OR INVOICE NUMBER  value 
            Dim pdvLineNo As New CrystalDecisions.Shared.ParameterDiscreteValue  'LOT NUMBER
            Dim pdvqty As New CrystalDecisions.Shared.ParameterDiscreteValue      'SHIPQPL
            'Dim pdvoption As New CrystalDecisions.Shared.ParameterDiscreteValue ' option ship nu or invoice num caption
            Dim pdvcustpo As New CrystalDecisions.Shared.ParameterDiscreteValue   'CUST PO
            Dim pdvtsspart As New CrystalDecisions.Shared.ParameterDiscreteValue   'TSS PART
            Dim pdvcustpart As New CrystalDecisions.Shared.ParameterDiscreteValue 'CUST PART
            Dim pdvcustomer As New CrystalDecisions.Shared.ParameterDiscreteValue 'CUSTOMER NMAE
            Dim pdvcuredate As New CrystalDecisions.Shared.ParameterDiscreteValue  'CURE DATE
            Dim pdvuom As New CrystalDecisions.Shared.ParameterDiscreteValue  'txtuom.text
            Dim pdvshipdate As New CrystalDecisions.Shared.ParameterDiscreteValue  'shipdate



            If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
                rptShipLabel.Load("Label2.rpt")

            ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
                ''rptShipLabel.Load("Label2Invoice.rpt")
                rptShipLabel.Load("Label2.rpt")

            End If

            pdvShipNo.Value = Trim(txtInvoiceNo.Text)
            pdvLineNo.Value = Trim(txtshiplot.Text) 'LOT NUMBER ALLOTED TO LINE NUNMBER
            pdvqty.Value = Trim(txtshipqpl.Text)

            pdvcustpo.Value = custpo
            pdvtsspart.Value = Trim(txtshippartno.Text)
            pdvcustpart.Value = custpart
            pdvcustomer.Value = customername
            pdvcuredate.Value = curedate
            pdvuom.Value = Trim(TxtUoM.Text)
            pdvshipdate.Value = Trim(txtshipinvdate.Text)


            If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
                pdvoption.Value = "Shipment No."
            ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
                pdvoption.Value = "Invoice No."

            End If

            pvCollection.Add(pdvShipNo)
            pvCollection2.Add(pdvLineNo)
            pvCollection3.Add(pdvqty)
            pvCollection4.Add(pdvoption)
            pvCollection5.Add(pdvcustpo)
            pvCollection6.Add(pdvtsspart)
            pvCollection7.Add(pdvcustpart)
            pvCollection8.Add(pdvcustomer)
            pvCollection9.Add(pdvcuredate)
            pvCollection10.Add(pdvuom)
            pvCollection11.Add(pdvshipdate)

            ' Apply the current parameter values.
            rptShipLabel.DataDefinition.ParameterFields("Shipno").ApplyCurrentValues(pvCollection)
            rptShipLabel.DataDefinition.ParameterFields("lotNumber").ApplyCurrentValues(pvCollection2)
            rptShipLabel.DataDefinition.ParameterFields("LotQty").ApplyCurrentValues(pvCollection3)
            rptShipLabel.DataDefinition.ParameterFields("option").ApplyCurrentValues(pvCollection4)

            rptShipLabel.DataDefinition.ParameterFields("custpo").ApplyCurrentValues(pvCollection5) 'customerpo
            rptShipLabel.DataDefinition.ParameterFields("tsspart").ApplyCurrentValues(pvCollection6) 'tsspart 
            rptShipLabel.DataDefinition.ParameterFields("custpart").ApplyCurrentValues(pvCollection7) 'custpart
            rptShipLabel.DataDefinition.ParameterFields("customer").ApplyCurrentValues(pvCollection8) 'customername
            rptShipLabel.DataDefinition.ParameterFields("curedate").ApplyCurrentValues(pvCollection9) 'curedate
            rptShipLabel.DataDefinition.ParameterFields("uom").ApplyCurrentValues(pvCollection10) 'uom
            rptShipLabel.DataDefinition.ParameterFields("shipdate").ApplyCurrentValues(pvCollection11) '





            Dim CTR As Integer
            Dim LABELCOUNT As Integer
            LABELCOUNT = Val(txtshiplabel.Text)
            Dim lastlabelqty As Integer

            'lastlabelqty = Val(TXTLOTQTYPORV.Text) - ((Val(txtNoofLabels.Text) - 1) * Val(TextlblqtyPORV.Text))
            lastlabelqty = Val(txtshipqty.Text) - ((Val(txtshiplabel.Text) - 1) * Val(txtshipqpl.Text))


            CTR = 1

            Do While CTR <= LABELCOUNT
                If CTR = LABELCOUNT Then

                    pdvqty.Value = (lastlabelqty)
                    pvCollection3.Add(pdvqty)
                    rptShipLabel.DataDefinition.ParameterFields("LotQty").ApplyCurrentValues(pvCollection3)

                    '    pdvlotnumberbarcode.Value = RTrim(LTrim(txtLot.Text)) + Trim(lastlabelqty)


                End If

                rptShipLabel.PrintToPrinter(1, False, 0, 0)


                'MsgBox("pdvqty.value")

                CTR = CTR + 1
            Loop


            pvCollection.Clear()
            pvCollection2.Clear()
            pvCollection3.Clear()

        End If

        If report = "C" Then

            Dim ctr1 As Integer
            Dim lblcnt As Integer

            lblcnt = Val(txtshiplabel.Text)
            Dim pvCollection5 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection6 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection7 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection8 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection9 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection10 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection11 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection12 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection13 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection14 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection15 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection16 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection17 As New CrystalDecisions.Shared.ParameterValues

            Dim pvCollection18 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection19 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection20 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection21 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection22 As New CrystalDecisions.Shared.ParameterValues
            Dim pvCollection23 As New CrystalDecisions.Shared.ParameterValues


            rptShipLabel.Load("BoxLabel.rpt")

            Dim pdvPackInfo As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvInvoiceNo As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvPartNo As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvPartName As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvCustPO As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvItemNumber As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvItemDescription As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvqty As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvSINo As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvMfgRef As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvUom As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvcustomer As New CrystalDecisions.Shared.ParameterDiscreteValue

            Dim pdvbuyer As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvdesig As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvdept As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvphone As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvmobile As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim pdvcustomername As New CrystalDecisions.Shared.ParameterDiscreteValue



            pdvPackInfo.Value = Trim(TxtInfo.Text) ' Transport details
            pdvInvoiceNo.Value = Trim(txtInvoiceNo.Text) 'ok invoice
            pdvPartNo.Value = Trim(TxtCustOrdNo.Text) ' shipmentnumber 
            pdvPartName.Value = Trim(txtQty.Text) 'invoice date
            pdvCustPO.Value = Trim(TxtCustPO.Text) 'ok customer po
            pdvItemNumber.Value = Trim(txtItem.Text) 'custname, shipto
            pdvItemDescription.Value = Trim(TxtDescription.Text) 'sp ad1
            pdvqty.Value = Trim(TxtCustItem.Text) 'ship ad2
            pdvSINo.Value = Trim(TxtCustDesc.Text) 'shipcity
            pdvMfgRef.Value = Trim(txtLotNo.Text) 'ship state
            pdvUom.Value = Trim(TxtUoM.Text) 'ship zip
            pdvcustomer.Value = Trim(txtCustomer.Text) ' customr contact
            pdvoption.Value = Trim(TxtCOLnNo.Text) ' customr contact

            pdvbuyer.Value = Trim(txtbuyer.Text) 'buyername
            pdvdesig.Value = Trim(txtdesig.Text) 'designation
            pdvdept.Value = Trim(txtdept.Text) 'department
            pdvphone.Value = Trim(txtphone.Text) 'phone
            pdvmobile.Value = Trim(txtmobile.Text) 'mobile
            pdvcustomername.Value = Trim(txtcusto.Text) 'customername


            pvCollection5.Add(pdvPackInfo)
            pvCollection6.Add(pdvInvoiceNo)
            pvCollection7.Add(pdvPartNo)
            pvCollection8.Add(pdvPartName)
            pvCollection9.Add(pdvCustPO)
            pvCollection10.Add(pdvItemNumber)
            pvCollection11.Add(pdvItemDescription)
            pvCollection12.Add(pdvqty)
            pvCollection13.Add(pdvSINo)
            pvCollection14.Add(pdvMfgRef)
            pvCollection15.Add(pdvUom)
            pvCollection16.Add(pdvcustomer)
            pvCollection17.Add(pdvoption)

            pvCollection18.Add(pdvbuyer)
            pvCollection19.Add(pdvdesig)
            pvCollection20.Add(pdvdept)
            pvCollection21.Add(pdvphone)
            pvCollection22.Add(pdvmobile)
            pvCollection23.Add(pdvcustomername)


            'rptShipLabel.DataDefinition.ParameterFields("Shipno").ApplyCurrentValues(pvCollection)

            rptShipLabel.DataDefinition.ParameterFields("PartNo").ApplyCurrentValues(pvCollection7)
            rptShipLabel.DataDefinition.ParameterFields("PartName").ApplyCurrentValues(pvCollection8) 'done
            rptShipLabel.DataDefinition.ParameterFields("Qty").ApplyCurrentValues(pvCollection12)

            rptShipLabel.DataDefinition.ParameterFields("InvoiceNo").ApplyCurrentValues(pvCollection6) 'done
            rptShipLabel.DataDefinition.ParameterFields("SINo").ApplyCurrentValues(pvCollection13)
            rptShipLabel.DataDefinition.ParameterFields("MfgRef").ApplyCurrentValues(pvCollection14) 'done
            rptShipLabel.DataDefinition.ParameterFields("CustPO").ApplyCurrentValues(pvCollection9)
            rptShipLabel.DataDefinition.ParameterFields("UoM").ApplyCurrentValues(pvCollection15)
            rptShipLabel.DataDefinition.ParameterFields("ItemNumber").ApplyCurrentValues(pvCollection10)
            rptShipLabel.DataDefinition.ParameterFields("ItemDescription").ApplyCurrentValues(pvCollection11)
            rptShipLabel.DataDefinition.ParameterFields("PackInfo").ApplyCurrentValues(pvCollection5) 'done
            rptShipLabel.DataDefinition.ParameterFields("customer").ApplyCurrentValues(pvCollection16)
            rptShipLabel.DataDefinition.ParameterFields("option").ApplyCurrentValues(pvCollection17)


            rptShipLabel.DataDefinition.ParameterFields("buyer").ApplyCurrentValues(pvCollection18)
            rptShipLabel.DataDefinition.ParameterFields("desig").ApplyCurrentValues(pvCollection19)
            rptShipLabel.DataDefinition.ParameterFields("dept").ApplyCurrentValues(pvCollection20)
            rptShipLabel.DataDefinition.ParameterFields("phone").ApplyCurrentValues(pvCollection21)
            rptShipLabel.DataDefinition.ParameterFields("mobile").ApplyCurrentValues(pvCollection22)
            rptShipLabel.DataDefinition.ParameterFields("custoname").ApplyCurrentValues(pvCollection23)




            ctr1 = 1

            Do While ctr1 <= lblcnt

                rptShipLabel.PrintToPrinter(1, False, 0, 0)

                ctr1 = ctr1 + 1
            Loop

            pvCollection5.Clear()
            pvCollection6.Clear()
            pvCollection7.Clear()
            pvCollection8.Clear()
            pvCollection9.Clear()
            pvCollection10.Clear()
            pvCollection11.Clear()
            pvCollection12.Clear()
            pvCollection13.Clear()
            pvCollection14.Clear()
            pvCollection15.Clear()
            pvCollection16.Clear()
            pvCollection17.Clear()

            pvCollection18.Clear()
            pvCollection19.Clear()
            pvCollection20.Clear()
            pvCollection21.Clear()
            pvCollection22.Clear()
            pvCollection23.Clear()

            rptShipLabel.Close()

        End If

        rptShipLabel.Close()


    End Sub

    Private Sub txtshipqpl_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtshipqpl.TextChanged

    End Sub

    Private Sub txtshipqpl_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtshipqpl.Leave

        If Val(txtshipqpl.Text) >= 1 Then

            Dim C As Integer
            Dim D As Integer

            C = Val(txtshipqty.Text) Mod Val(txtshipqpl.Text)

            D = Val(txtshipqty.Text) - Val(C)
            txtshiplabel.Text = D / Val(txtshipqpl.Text)

            If C > 0 Then

                txtshiplabel.Text = txtshiplabel.Text + 1
                'txtNoofLabels.Text = txtNoofLabels.Text + 1

            End If

        End If


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Exit Sub
        Me.Close()

    End Sub

    Private Sub btnPorvPrintCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPorvPrintCancel.Click
        'Exit Sub
        Me.Close()

    End Sub

    Private Sub btnclear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclear.Click

        Dim cnSQLL As SqlConnection
        'Dim cmSQLL As SqlCommand
        'Dim drSQLL As SqlDataReader
        Dim strSQLL As String



        strSQLL = "Select LotNumber,CONumber,ItemNumber,COLineNumber,ShippedQuantity,ItemNumber,ItemNumber," & _
                "ItemNumber, ItemNumber, ItemNumber  " & _
                "from TSS_ItemLabel_Ver2 " & _
                "where ShipmentNumber  = 999999999 and " & _
                "COLineNumber like '9999' order by COLineNumber "



        '--dummy
        'AS
        'SELECT     dbo.TSS_History_Shipment.LotNumber, dbo.TSS_History_Shipment.CONumber, dbo.TSS_History_Shipment.ItemNumber, dbo.TSS_History_Shipment.COLineNumber, 
        '                     dbo.TSS_History_Shipment.ShipmentNumber, dbo.TSS_History_Shipment.ShipmentDate, dbo.TSS_History_Shipment.ShippedQuantity, 
        '      dbo.TSS_History_Shipment.CustomerID, dbo._NoLock_FS_ARInvoiceHeader.InvoiceNumber


        'end of dummy




        cnSQLL = New SqlConnection(ConnectionString)
        'cnSQL.Open()


        Dim stockDShip As DataSet = New DataSet


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQLL, cnSQLL)
        Dim stockDAS As SqlDataAdapter = New SqlDataAdapter
        'Try
        stockDAS.SelectCommand = sqlCmd
        cnSQLL.Open()

        stockDAS.TableMappings.Add("Table", "Stock")
        'get data
        stockDAS.Fill(stockDShip)

        stockDShip.Tables(0).Columns(0).ColumnName = "LotNumber"
        stockDShip.Tables(0).Columns(1).ColumnName = "OrderNo."

        stockDShip.Tables(0).Columns(2).ColumnName = "ItemNumber"
        stockDShip.Tables(0).Columns(3).ColumnName = "CoLineNo."
        stockDShip.Tables(0).Columns(4).ColumnName = "ShipQty"

        stockDShip.Tables(0).Columns(5).ColumnName = "Description"
        stockDShip.Tables(0).Columns(6).ColumnName = "CustItemNumber"

        stockDShip.Tables(0).Columns(7).ColumnName = "uom"
        stockDShip.Tables(0).Columns(8).ColumnName = "Customer"
        stockDShip.Tables(0).Columns(9).ColumnName = "PONumber"


        DataGridShip.DataSource = stockDShip.Tables(0)
        cnSQLL.Close()
        DataGridShip.Expand(-1)
        'MsgBox("Select Lot Number for label print", MsgBoxStyle.Information, "SHIP-LABELS")



    End Sub

    Private Sub btnporvok_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnporvok.Click

        If RBPOwise.Checked = True Or RBDateWise.Checked = True Or RBPorvIMTRGoel.Checked = True Then
        Else
            MsgBox("Select the option before clicking ok", vbInformation)
        End If


        If RdbPORV.Checked = True Or RBPORVGoel.Checked = True Or RBPorvIMTRGoel.Checked = True Then

            If RBPOwise.Checked = True Or RBDateWise.Checked = True Then

                datagridStock.Visible = True

                datagridStock.Enabled = True

                FillDataGrid()
            Else
                MsgBox("Select the option before clicking ok", vbInformation, "Label Software")
                Exit Sub
            End If
        Else
            MsgBox("Select the option before clicking ok", vbInformation, "Label Software")
            Exit Sub
        End If

        porvclear()
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtconumber.TextChanged

    End Sub

    Private Sub LblInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LblInvoice.Click

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles dtpFROM.ValueChanged

    End Sub

    Private Sub RBPOwise_CheckedChanged(sender As Object, e As EventArgs) Handles RBPOwise.CheckedChanged



        lblfrom.Enabled = False
        lblto.Enabled = False
        dtpFROM.Enabled = False
        dtpTo.Enabled = False


        lblPONo.Enabled = True
        txtPonumber.Enabled = True

        If RBPOwise.Checked = True Then
            txtPonumber.Enabled = True
            TXTPartSelect.Enabled = False
        End If


    End Sub

    Private Sub RBDateWise_CheckedChanged(sender As Object, e As EventArgs) Handles RBDateWise.CheckedChanged
        lblfrom.Enabled = True
        lblto.Enabled = True
        dtpFROM.Enabled = True
        dtpTo.Enabled = True

        lblPONo.Enabled = False
        txtPonumber.Enabled = False

        If RBDateWise.Checked = True Then
            txtPonumber.Enabled = False
            TXTPartSelect.Enabled = True
        End If


        txtPonumber.Enabled = False


    End Sub

    Private Sub RBPORVGoel_CheckedChanged(sender As Object, e As EventArgs) Handles RBPORVGoel.CheckedChanged
        Groupporv.Enabled = True
        TXTPartSelect.Text = "%"
        groupShip.Enabled = False
        txtPonumber.Text = ""

        RBDateWise.Text = "PORV DT Wise"

        porvclear()

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RBPorvIMTRGoel.CheckedChanged
        Groupporv.Enabled = True
        TXTPartSelect.Text = "%"
        groupShip.Enabled = False
        txtPonumber.Text = ""
        RBDateWise.Text = "IMTR DT Wise"
        porvclear()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        Dim pages As New List(Of Metafile)
        Dim pageIndex As Integer = 0
        Dim doc As New Printing.PrintDocument()
        Dim ReportViewer1 As New ReportViewer

        ' //Get a reference to the default credentials
        Dim credentials As System.Net.ICredentials
        credentials = System.Net.CredentialCache.DefaultCredentials


        Dim rsCredentials As ReportServerCredentials
        rsCredentials = ReportViewer1.ServerReport.ReportServerCredentials

        '  // Set the credentials for the server report
        rsCredentials.NetworkCredentials = credentials

        With ReportViewer1
            .Visible = False
            .ProcessingMode = ProcessingMode.Remote
            .ServerReport.ReportPath = "/Reports/IT-PendingForTesting/Modified Reports/LabelPrint"
            .ServerReport.ReportServerUrl = New  _
         Uri("http://tssblrfsh101/reportserver")
        End With

        Me.ReportViewer1.RefreshReport()
        'ReportViewer1.ServerReport.SetParameters(mReportParams)


        doc = New Printing.PrintDocument()
        '  AddHandler doc.PrintPage, AddressOf PrintPageHandler
        Dim dialog As New PrintDialog()
        dialog.Document = doc
        Dim print As DialogResult
        print = dialog.ShowDialog()

        doc.PrinterSettings = dialog.PrinterSettings

        Dim deviceInfo As String = _
        "<DeviceInfo>" & _
        "<OutputFormat>emf</OutputFormat>" & _
        "  <PageWidth>8.5in</PageWidth>" & _
        "  <PageHeight>11in</PageHeight>" & _
        "  <MarginTop>0.25in</MarginTop>" & _
        "  <MarginLeft>0.25in</MarginLeft>" & _
        "  <MarginRight>0.25in</MarginRight>" & _
        "  <MarginBottom>0.25in</MarginBottom>" & _
        "</DeviceInfo>"

        Dim warnings() As Microsoft.Reporting.WinForms.Warning
        Dim streamids() As String
        Dim mimeType, encoding, filenameExtension, path As String
        mimeType = "" : encoding = "" : filenameExtension = ""

        'Input parameter report
        ' Dim Qty As Integer = 450
        'Dim HSK As Integer = 7181844

        Dim Quantity As New ReportParameter '("Qty", Qty)
        Dim HistoryShipmentKey As New ReportParameter '("HSK", HSK)

        Dim parmSO1(1) As ReportParameter

        ' parmSO1(0) = parmDateFrom
        ' parmSO1(1) = parmDateTo


        parmSO1(0) = Quantity
        parmSO1(1) = HistoryShipmentKey


       



                '   Dim data() As Byte

                '      ReportViewer1.ServerReport.SetParameters(parmSO1(0))
                '     ReportViewer1.ServerReport.SetParameters(parmSO1(1))

                'data = ReportViewer1.ServerReport.Render("Image", _
                '       deviceInfo, mimeType, encoding, filenameExtension, _
                '       streamids, warnings)
                '   pages.Add(New Metafile(New MemoryStream(data)))

                'For Each pageName As String In streamids
                '    data = ReportViewer1.ServerReport.RenderStream("Image", _
                '           pageName, deviceInfo, mimeType, encoding)
                '    pages.Add(New Metafile(New MemoryStream(data)))
                'Next
                doc.Print()
                ' Me.ReportViewer1.RefreshReport()
    End Sub

    'Private Sub PrintPageHandler(ByVal sender As Object, _
    '   ByVal e As PrintPageEventArgs)
    'Dim page As Metafile = pages(pageIndex)
    '   pageIndex += 1
    'Dim pWidth As Integer = 827
    'Dim pHeight As Integer = 1100
    '   e.Graphics.DrawImage(page, 0, 0, pWidth, pHeight)
    '  e.HasMorePages = pageIndex < pages.Count
    'End Sub
    ' Private Sub frmPrintReport_Load(ByVal sender As System.Object, _
    '    ByVal e As System.EventArgs) Handles MyBase.Load

    '   Button1.Text = "Print"
    '  With ReportViewer1
    '     .Visible = False
    '    .ProcessingMode = ProcessingMode.Remote
    '   .ServerReport.ReportPath = "/Reports/Master Reports/User Access"
    '  .ServerReport.ReportServerUrl = New  _
    '          Uri("http://TSSBLRFSH101/reportserver")
    'End With
    'Me.Controls.Add(ReportViewer1)

    'End Sub
    'Create the sales order number report parameter  
    ' Dim salesOrderNumber As New ReportParameter()
    ' salesOrderNumber.Name = "SalesOrderNumber"
    'salesOrderNumber.Values.Add("SO43661")

    'Set the report parameters for the report  
    '  Dim parameters() As ReportParameter = {salesOrderNumber}
    ' serverReport.SetParameters(parameters)

    ' Dim print1 As New printersettings()
    'Dim pd As New ReportDocument()

    ' pd.Print()


    'Refresh the report  
    '    ReportViewer1.RefreshReport()
    'End Sub
End Class