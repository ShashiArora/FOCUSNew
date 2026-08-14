Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
'Imports SoftBrands.FourthShift.Transaction
'Imports CrystalDecisions.CrystalReports.Engine
Imports System.Windows.Forms



Public Class mainmenu




    Inherits System.Windows.Forms.Form
    Public mode1 As String
    Dim granted As String
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem26 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem27 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem31 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem32 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem30 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem35 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem33 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem34 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem37 As System.Windows.Forms.MenuItem
    Dim screenname As String
    'Dim screentype As String
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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem26 = New System.Windows.Forms.MenuItem()
        Me.MenuItem25 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.MenuItem15 = New System.Windows.Forms.MenuItem()
        Me.MenuItem21 = New System.Windows.Forms.MenuItem()
        Me.MenuItem17 = New System.Windows.Forms.MenuItem()
        Me.MenuItem18 = New System.Windows.Forms.MenuItem()
        Me.MenuItem14 = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.MenuItem13 = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.MenuItem19 = New System.Windows.Forms.MenuItem()
        Me.MenuItem20 = New System.Windows.Forms.MenuItem()
        Me.MenuItem22 = New System.Windows.Forms.MenuItem()
        Me.MenuItem23 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.MenuItem27 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.MenuItem16 = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.MenuItem29 = New System.Windows.Forms.MenuItem()
        Me.MenuItem33 = New System.Windows.Forms.MenuItem()
        Me.MenuItem34 = New System.Windows.Forms.MenuItem()
        Me.MenuItem37 = New System.Windows.Forms.MenuItem()
        Me.MenuItem31 = New System.Windows.Forms.MenuItem()
        Me.MenuItem32 = New System.Windows.Forms.MenuItem()
        Me.MenuItem36 = New System.Windows.Forms.MenuItem()
        Me.MenuItem30 = New System.Windows.Forms.MenuItem()
        Me.MenuItem35 = New System.Windows.Forms.MenuItem()
        Me.MenuItem12 = New System.Windows.Forms.MenuItem()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3, Me.MenuItem4, Me.MenuItem12})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem26, Me.MenuItem25, Me.MenuItem5, Me.MenuItem15, Me.MenuItem21, Me.MenuItem17, Me.MenuItem18, Me.MenuItem14, Me.MenuItem7, Me.MenuItem6, Me.MenuItem13, Me.MenuItem8})
        Me.MenuItem1.Text = "&Marketing"
        '
        'MenuItem26
        '
        Me.MenuItem26.Index = 0
        Me.MenuItem26.Text = "Pricing"
        '
        'MenuItem25
        '
        Me.MenuItem25.Index = 1
        Me.MenuItem25.Text = "My Customer"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 2
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem9, Me.MenuItem10})
        Me.MenuItem5.Text = "&Enquiry"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 0
        Me.MenuItem9.Text = "Add"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 1
        Me.MenuItem10.Text = "Edit"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 3
        Me.MenuItem15.Text = "Enquiry Price -Customer-Parts Creation View"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 4
        Me.MenuItem21.Text = "Enquiry  Price Quick View"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 5
        Me.MenuItem17.Text = "Enquiry Status View"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 6
        Me.MenuItem18.Text = "Pending for Approval"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 7
        Me.MenuItem14.Text = "Part Number Creation"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 8
        Me.MenuItem7.Text = "&Customer Creation"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 9
        Me.MenuItem6.Text = "&RFQ Pending"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 10
        Me.MenuItem13.Text = "R&FQ Completed"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 11
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem19, Me.MenuItem20, Me.MenuItem22, Me.MenuItem23})
        Me.MenuItem8.Text = "Pro&jects"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 0
        Me.MenuItem19.Text = "Pending for Approval"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 1
        Me.MenuItem20.Text = "Master List"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 2
        Me.MenuItem22.Text = "Pricing Pending"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 3
        Me.MenuItem23.Text = "Pricing Completed"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem11, Me.MenuItem27})
        Me.MenuItem2.Text = "&Purchase"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 0
        Me.MenuItem11.Text = "Sub Contractor -DC"
        '
        'MenuItem27
        '
        Me.MenuItem27.Index = 1
        Me.MenuItem27.Text = "Sub Contractor - Return DC"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem16})
        Me.MenuItem3.Text = "&SCM"
        Me.MenuItem3.Visible = False
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 0
        Me.MenuItem16.Text = "Domestic Shipment Labels"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 3
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem29, Me.MenuItem31, Me.MenuItem32, Me.MenuItem36, Me.MenuItem30, Me.MenuItem35})
        Me.MenuItem4.Text = "&Stores -Consumable Mgmt"
        '
        'MenuItem29
        '
        Me.MenuItem29.Index = 0
        Me.MenuItem29.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem33, Me.MenuItem34, Me.MenuItem37})
        Me.MenuItem29.Text = "&Masters"
        '
        'MenuItem33
        '
        Me.MenuItem33.Index = 0
        Me.MenuItem33.Text = "&Item Master"
        '
        'MenuItem34
        '
        Me.MenuItem34.Index = 1
        Me.MenuItem34.Text = "&Stock Rooms"
        '
        'MenuItem37
        '
        Me.MenuItem37.Index = 2
        Me.MenuItem37.Text = "&Approval  Work Flow "
        '
        'MenuItem31
        '
        Me.MenuItem31.Index = 1
        Me.MenuItem31.Text = "&Receipts"
        '
        'MenuItem32
        '
        Me.MenuItem32.Index = 2
        Me.MenuItem32.Text = "&Material Request "
        '
        'MenuItem36
        '
        Me.MenuItem36.Index = 3
        Me.MenuItem36.Text = "&Approvals"
        '
        'MenuItem30
        '
        Me.MenuItem30.Index = 4
        Me.MenuItem30.Text = "Material I&ssues"
        '
        'MenuItem35
        '
        Me.MenuItem35.Index = 5
        Me.MenuItem35.Text = "Material Receipt at Mfg &Bldg"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 4
        Me.MenuItem12.Text = "&LogOff"
        '
        'mainmenu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1136, 634)
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMenu1
        Me.Name = "mainmenu"
        Me.Text = "Focus Main Menu - Ver 4.0   Apr-2018"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub mainmenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Set the color in the MDI client.
        For Each ctl As Control In Me.Controls
            If TypeOf ctl Is MdiClient Then
                ctl.BackColor = Me.BackColor
            End If
        Next ctl


        ' Display a child form.
        Dim frm As New mainmenu
        'frm.MdiParent = Me
        'frm.Width = Me.Width \ 2
        'frm.Height = Me.Height \ 2
        'frm.Show()

    End Sub



    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Exit Sub

    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click

        Dim enqadd As New Enquiry
        mode = "Add"
        screenname = "Enquiry"
        granted = ""
        CHECK()

        If granted = "YES" Then
            enqadd.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If


    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Dim rfqadd As New RFQ
        screenname = "RFQ"
        screentype = "PEND"
        granted = ""
        CHECK()

        If granted = "YES" Then
            '  rfqadd.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click

        Dim part As New PartCreation

        screenname = "PartCreation"
        granted = ""
        CHECK()

        If granted = "YES" Then
            part.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Dim enqedit As New Enquiry

        mode = "Edit"

        ' enqedit.Show()

        screenname = "Enquiry"
        granted = ""
        CHECK()

        If granted = "YES" Then
            enqedit.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If









    End Sub

    Public Sub CHECK()

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String


        strSQL = "Select * from dbo.ENQ_User_Rights where  User_Id = '" & username & "' AND Screen = '" & screenname & "'"

        cnSQL.Open()

        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()

        If drSQL.Read() Then

            If IsDBNull(drSQL.Item(0)) Then
                MsgBox("Permission denied", vbInformation)
                granted = "NO"
                Exit Sub

            Else
                granted = "YES"
            End If
        End If


    End Sub


    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click

        Dim rfqadd As New RFQ
        screenname = "RFQ"
        screentype = "COMP"
        granted = ""
        CHECK()

        If granted = "YES" Then
            '    rfqadd.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click

        Dim rfqprice As New RFQPriceViewALL
        'rfqcomp = "sales"
        datatype = ""
        screenname = "RFQPriceView"
        granted = ""
        CHECK()

        If granted = "YES" Then
            RFQPriceViewALL.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If



    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Dim customer As New CustomerCreation
        'rfqcomp = "sales"

        screenname = "CustomerCreation"
        granted = ""
        CHECK()

        If granted = "YES" Then
            customer.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem17.Click
        Dim EnquirySt As New EnquiryStatus
        'rfqcomp = "sales"

        screenname = "EnquiryStatus"
        granted = ""
        CHECK()

        If granted = "YES" Then
            EnquirySt.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If


    End Sub

    Private Sub MenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem18.Click

    End Sub

    Private Sub MenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem19.Click

        Dim Projectupd As New ProjectsUpdate

        screenname = "ProjectsUpdate"
        granted = ""
        CHECK()

        If granted = "YES" Then
            Projectupd.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem20.Click
        Dim ProjectMas As New ProjectMasterList

        screenname = "ProjectMasterList"
        granted = ""
        CHECK()

        If granted = "YES" Then
            ProjectMas.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim EnquirySt As New RFQPriceViewALL
        'rfqcomp = "sales"

        screenname = "RFQPriceView"
        granted = ""
        datatype = "ALL"
        CHECK()

        If granted = "YES" Then
            RFQPriceViewALL.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If


    End Sub

    Private Sub MenuItem21_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem21.Click
        Dim EnquirySt As New EnquiryStatus
        'rfqcomp = "sales"

        screenname = "RFQPriceViewQuick"
        granted = ""
        CHECK()

        If granted = "YES" Then
            RFQPriceViewQuick.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem22.Click
        Dim rfqaddp As New RFQ

        screenname = "RFQP"
        screentype = "PENDP"
        granted = ""
        CHECK()

        If granted = "YES" Then
            '  rfqaddp.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub MenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem23.Click
        Dim rfqaddpc As New RFQ
        screenname = "RFQP"
        screentype = "COMPP"
        granted = ""
        CHECK()

        If granted = "YES" Then
            '  rfqaddpc.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub MenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '   Dim PORVAuto As New Porv
        '  screenname = "Porv"
        ' granted = ""
        'CHECK()

        'If granted = "YES" Then
        ShipLabelDomestic.Show()
        'Else
        'MsgBox("Permission denied", vbInformation)
        'Exit Sub
        'End If
    End Sub

    Private Sub MenuItem25_Click(sender As Object, e As EventArgs) Handles MenuItem25.Click
        Dim mycust As New MyCustomer
        mode = "Add"
        screenname = "MyCustomer"
        granted = ""
        CHECK()

        If granted = "YES" Then
            mycust.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub


    Private Sub MenuItem26_Click(sender As Object, e As EventArgs) Handles MenuItem26.Click

        Dim price As New Pricing
        mode = "Add"
        screenname = "Pricing"
        granted = ""
        CHECK()

        If granted = "YES" Then
            price.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If



    End Sub





    Private Sub MenuItem11_Click(sender As Object, e As EventArgs) Handles MenuItem11.Click
        Dim SDC As New SUBDC
        mode = "Add"
        screenname = "SUBDC"
        granted = ""
        CHECK()

        If granted = "YES" Then
            SDC.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem16_Click(sender As Object, e As EventArgs) Handles MenuItem16.Click
        Dim price As New ShipLabelDomestic
        mode = "Add"
        screenname = "ShipLabelDomestic"
        granted = ""
        CHECK()

        If granted = "YES" Then
            ShipLabelDomestic.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub MenuItem24_Click_1(sender As Object, e As EventArgs)
        Dim Itemm As New WHItemMaster
        mode = "Add"
        screenname = "ItemMaster"
        granted = ""
        '  CHECK()
        Itemm.Show()
        If granted = "YES" Then
            Itemm.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub MenuItem31_Click(sender As Object, e As EventArgs) Handles MenuItem31.Click
        Dim Receiptm As New WHReceipts

        mode = "Add"
        screenname = "WHReceipts"
        granted = ""
        CHECK()
        '  Receiptm.Show()
        If granted = "YES" Then
            Receiptm.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub MenuItem28_Click(sender As Object, e As EventArgs)
        '    Dim Locstk As New StockRoom
        '   mode = "Add"
        '  screenname = "StockRoom"
        ' granted = ""
        '  CHECK()
        'Locstk.Show()
        'If granted = "YES" Then
        'Locstk.Show()
        'Else
        'MsgBox("Permission denied", vbInformation)
        'Exit Sub
        'End If
    End Sub

    Private Sub MenuItem32_Click(sender As Object, e As EventArgs) Handles MenuItem32.Click
        Dim Matreq As New WHMaterialReq
        mode = "Add"
        screenname = "WHMaterialReq"
        granted = ""
        CHECK()
        ' Matreq.Show()
        If granted = "YES" Then
            Matreq.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub MenuItem30_Click(sender As Object, e As EventArgs) Handles MenuItem30.Click
        Dim MatIss As New WHMaterialIssue

        mode = "Add"
        screenname = "WHMaterialIssue"
        granted = ""
        CHECK()
        '  MatIss.Show()
        If granted = "YES" Then
            MatIss.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub MenuItem36_Click(sender As Object, e As EventArgs) Handles MenuItem36.Click
        Dim MatreqAp As New WHMatRequestApproval

        mode = "Add"
        screenname = "WHMatRequestApproval"
        granted = ""
        CHECK()
        ' WHMatRequestApproval.Show()
        If granted = "YES" Then
            MatreqAp.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub MenuItem34_Click(sender As Object, e As EventArgs)

        Dim rettool As New WHReturnToolReceipt

        mode = "Add"
        screenname = "WHReturnToolReceipt"
        granted = ""
        '  CHECK()
        WHReturnToolReceipt.Show()
        If granted = "YES" Then
            rettool.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If


    End Sub

    Private Sub MenuItem35_Click(sender As Object, e As EventArgs) Handles MenuItem35.Click
        Dim Receiptmfg As New WHMatReceiptMfgBldg

        mode = "Add"
        screenname = "WHMatReceiptMfgBldg"
        granted = ""
        CHECK()
        '  WHMatReceiptMfgBldg.Show()
        If granted = "YES" Then
            Receiptmfg.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub MenuItem12_Click(sender As Object, e As EventArgs) Handles MenuItem12.Click
        WHMaterialIssue.Close()
        WHMaterialReq.Close()
        WHMatReceiptMfgBldg.Close()
        WHMatRequestApproval.Close()
        WHReceipts.Close()
        Me.Close()

    End Sub

    Private Sub MenuItem33_Click(sender As Object, e As EventArgs) Handles MenuItem33.Click

        Dim item As New WHItemMaster
        mode = "Add"
        screenname = "WHItemMaster"
        granted = ""
        CHECK()
        '  Receiptm.Show()
        If granted = "YES" Then
            item.Show()
        Else
            MsgBox("Permission denied", vbInformation)
            Exit Sub
        End If
        





    End Sub
End Class
