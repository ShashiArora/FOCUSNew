Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
'Imports SoftionBrands.FourthShift.Transaction
'Imports Microsoft.Office.Interop.Outlook
'Imports CrystalDecisions.CrystalReports.Engine
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection

Public Class ProjectsUpdate

    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Private Sub ProjectsUpdate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load



        RadioButtonPending.Checked = True

        DtpSOP.Format = DateTimePickerFormat.Custom
        DtpSOP.CustomFormat = "MMM yyyy"


        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand

        strSql = "SELECT Int_code,AE_NAME FROM ENQ_Project_Engineers " & _
                 "WHERE Status like 'A%' ORDER BY AE_NAME"
        cmSQL = New SqlCommand(strSql, sqlCon)
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource As SqlDataAdapter = New SqlDataAdapter
        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eSource")
        With ComboBoxEngineers
            .DataSource = source.Tables("eSource")
            .DisplayMember = "AE_NAME"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With

        'load project types


        strSql = "SELECT Proj_Intcode, Proj_Type FROM ENQ_ProjectTypes " & _
                 "WHERE Status like 'A%' ORDER BY Proj_Intcode"
        cmSQL = New SqlCommand(strSql, sqlCon)
        Dim sqlCmd1 As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource1 As SqlDataAdapter = New SqlDataAdapter
        ESource1.SelectCommand = sqlCmd1
        ESource1.Fill(source, "eSource1")
        With ComboBoxProjectType
            .DataSource = source.Tables("eSource1")
            .DisplayMember = "Proj_Type"
            .ValueMember = "Proj_Intcode"
            .SelectedIndex = 0
        End With

        listloaddocuments()
        listloadCertificateDetails()


        barcode()

    End Sub
    Private Sub listloaddocuments()


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        'Dim a As ListView


        cnSQL1.Open()
        strSQL1 = "SELECT Documents, Int_code FROM ENQ_Project_Documents " & _
                 "WHERE  Status = 'A'"
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        Dim ColumnValue As String = Nothing
        Do While drSQL1.Read()

            ColumnValue = (drSQL1.GetValue(0)).ToString
            CheckedListBoxDoc.Items.Add(ColumnValue)
            CheckedListBoxDoc.ValueMember = "Int_code"

        Loop
    End Sub

    Private Sub ButtonRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRefresh.Click

        DataGridViewProjectPending.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        'strSQL = "SELECT RegNo,  [Reg.Date], CustomerID, CustomerName, City, Class,Class1, Cust_Exist_New as Exist_Cust, CSR, TSSISeg, TSSSeg,MarketType, Enq_Ref_no, Enq_Ref_date, " & _
        '         "Enq_Source, Enq_Recd_date,Doc_upload,Doc_Details,Special_instructions, Enq_Int_code as Key1,Cust_IntCode from TSS_Enq_Pending_Project_Aproval "

        '        strSQL = "SELECT a.RegNo,  a.[Reg.Date], a.CustomerID, a.CustomerName, a.City, a.Class,a.Class1, a.Cust_Exist_New as Exist_Cust, a.CSR, a.TSSISeg, a.TSSSeg,a.MarketType, a.Enq_Ref_no, a.Enq_Ref_date, " & _
        '                "a.Enq_Source, a.Enq_Recd_date,a.Doc_upload, a.Doc_Details,a.Special_instructions, a.Enq_Int_code as Key1,a.Cust_IntCode from TSS_Enq_Pending_Project_Aproval a where " & _
        '               "a.RegNo not in (select b.Enq_Reg_No from ENQ_Project_Approval_Status b where (  b.Status in ('NotProject','Rejected')) and a.RegNo = b.Enq_Reg_No) ORDER BY a.RegNo"


        strSQL = "SELECT a.RegNo,  a.[Reg.Date], a.CustomerID, a.CustomerName, a.City, a.Class,a.Class1, a.Cust_Exist_New as Exist_Cust, a.CSR, a.TSSISeg, a.TSSSeg,a.MarketType, a.Enq_Ref_no, a.Enq_Ref_date, " & _
             "a.Enq_Source, a.Enq_Recd_date,a.Doc_upload, a.Doc_Details,a.Special_instructions, a.Enq_Int_code as Key1,a.Cust_IntCode from TSS_Enq_Pending_Project_Aproval a where " & _
            "a.RegNo not in (select b.Enq_Reg_No from ENQ_Project_Approval_Status b where (b.Status <> 'MoreInfoRequired') and a.RegNo = b.Enq_Reg_No) order by a.RegNo"


        ' strSQL = "SELECT a.RegNo,  a.[Reg.Date], a.CustomerID, a.CustomerName, a.City, a.Class,a.Class1, a.Cust_Exist_New as Exist_Cust, a.CSR, a.TSSISeg, a.TSSSeg,a.MarketType, a.Enq_Ref_no, a.Enq_Ref_date, " & _
        '  "a.Enq_Source, a.Enq_Recd_date,a.Doc_upload, a.Doc_Details,a.Special_instructions, a.Enq_Int_code as Key1,a.Cust_IntCode from TSS_Enq_Pending_Project_Aproval a where " & _
        ' "a.RegNo not in (select b.Enq_Reg_No from ENQ_Project_Approval_Status b where (b.Status <> 'MoreInfoRequired' OR b.Status <> 'NotProject' OR b.Status <> 'Rejected') and a.RegNo = b.Enq_Reg_No)"




        'Enq_Int_code as Key1, Cust_IntCode 

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewProjectPending.DataSource = stockDC.Tables(0)
        cnSQL.Close()



    End Sub

    Private Sub DataGridViewProjectPending_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewProjectPending.CellClick
        '      MsgBox("cellclick")

    End Sub


    Private Sub DataGridViewProjectPending_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewProjectPending.CellContentClick


    End Sub

    Private Sub Label28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxenqDetails.Enter

    End Sub

    Private Sub ProjectPending_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridViewProjectPending.RowHeaderMouseClick
        'c = DataGridViewProjectPending.CurrentCell.Value

        'strSQL = "SELECT RegNo,  [Reg.Date], CustomerID, CustomerName, City, Class,Class1, Cust_Exist_New as Exis_Cust, CSR, TSSISeg, TSSSeg,MarketType, 11
        'Enq_Ref_no, Enq_Ref_date, " & _"
        '        "Enq_Source, Enq_Recd_date,Doc_upload,Doc_Details,Special_instructions from TSS_Enq_Pending_Project_Aproval"

        GroupBox2.Enabled = True

        txtRegNo.Text = DataGridViewProjectPending.CurrentRow.Cells(0).Value.ToString
        DtpRegDate.Value = DataGridViewProjectPending.CurrentRow.Cells(1).Value
        txtcustomerid.Text = DataGridViewProjectPending.CurrentRow.Cells(2).Value.ToString
        txtCustomer.Text = DataGridViewProjectPending.CurrentRow.Cells(3).Value.ToString
        txtCustcity.Text = DataGridViewProjectPending.CurrentRow.Cells(4).Value.ToString

        txtEnqRef.Text = DataGridViewProjectPending.CurrentRow.Cells(12).Value.ToString
        dtpEnqDt.Value = DataGridViewProjectPending.CurrentRow.Cells(13).Value.ToString
        txtEnqSource.Text = DataGridViewProjectPending.CurrentRow.Cells(14).Value.ToString
        DTPEnqRecd.Value = DataGridViewProjectPending.CurrentRow.Cells(15).Value.ToString

        If Trim(DataGridViewProjectPending.CurrentRow.Cells(16).Value.ToString()) = "YES" Then
            rbdocyes.Checked = True
            rbDocNo.Checked = False
        Else
            rbDocNo.Checked = True
            rbdocyes.Checked = True


        End If

        txtDocDetails.Text = DataGridViewProjectPending.CurrentRow.Cells(17).Value.ToString
        txtSpecial.Text = DataGridViewProjectPending.CurrentRow.Cells(18).Value.ToString

        txtenqintcode.Text = DataGridViewProjectPending.CurrentRow.Cells(19).Value

        cleardetails()
        fillPartList()


        cleardocdetails()
        filldocdetails()

        ClearCertDetails()
        clearqtydetails()

        callapprovaldetails()



        ButtonSave.Enabled = True

    End Sub
    Private Sub callapprovaldetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String


        strSQL1 = "select * from ENQ_Project_Approval_Status where Enq_Reg_No = '" & txtRegNo.Text & "' "


        'Enq_Reg_No	numeric(18, 0)	Unchecked
        'Status	varchar(50)	Checked
        'Remarks	varchar(500)	Checked
        'Project_Number	varchar(50)	Checked
        'Project_RegDate	datetime	Checked
        'Project_Alotted	varchar(50)	Checked


        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then
            If IsDBNull(drSQL1.Item(0)) Then

            Else
                If drSQL1.Item(1) = "Accepted" Then
                    RadioButtonAcptd.Checked = True
                ElseIf drSQL1.Item(1) = "Rejected" Then
                    RadioButtonRejected.Checked = True
                ElseIf drSQL1.Item(1) = "NotProject" Then
                    RadioButtonNotProject.Checked = True
                ElseIf drSQL1.Item(1) = "MoreInfoRequired" Then
                    RadioButtonMoreInfo.Checked = True
                End If

                txtremarks.Text = drSQL1.Item(2)
                txtProjectNumber.Text = drSQL1.Item(3)
                ComboBoxEngineers.Text = drSQL1.Item(5)


            End If
        End If


    End Sub
    Sub fillPartList()

        'datagridDetail.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        strSql = "SELECT Sl_no,FS_Yes_NO,Part_Source,PartNumber,PartDescription,CustPartNumber,CustPartDescription,uom,Dimension,Material,Special,Req,Enq_Detail_code as DetailKey FROM ENQ_Details where Enq_Int_code = " & txtenqintcode.Text & " order by Sl_no"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd

        sqlCon.Open()
        stockDAC.TableMappings.Add("Table", "Part")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewItemDetail.DataSource = stockDC.Tables(0)
        sqlCon.Close()


    End Sub

    Private Sub DtpRegDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtpRegDate.ValueChanged

    End Sub

    Private Sub ButtonCustDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCustDetails.Click

        'GroupBoxenqDetails.Width = 1240
        'GroupBoxenqDetails.Height = 172

        GroupBoxCustDetails.Visible = True
        GroupBoxCustDetails.Width = 1170 '864
        GroupBoxCustDetails.Height = 163
        custdetails()


    End Sub

    Private Sub Label52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label52.Click

    End Sub

    Private Sub txtcustClose_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBoxCustDetails.Visible = False

        ' GroupBoxenqDetails.Height = 121


    End Sub

    Private Sub BtnCustClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCustClose.Click
        GroupBoxCustDetails.Visible = False
        ' GroupBoxenqDetails.Height = 172
    End Sub

    Private Sub DataGridQty_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridQty.Navigate

    End Sub

    Private Sub RadioButtonAcptd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonAcptd.CheckedChanged
        GroupBoxProjectDetails.Visible = True
        txtProjectNumber.Text = "IRP" & "_" & Trim(DataGridViewProjectPending.CurrentRow.Cells(8).Value.ToString) & "_" & Year(DtpRegDate.Value) & "_" & txtRegNo.Text

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String
        Dim earno As Integer


        strSQL1 = "select max(LastUsed_No)from ENQ_Project_Document_Control WHERE Project_Doc_Name = 'EAR'"

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                earno = 1
            Else
                earno = drSQL1.Item(0) + 1
            End If


        End If


        txtEARNumber.Text = "EIR" & "_" & Year(DtpRegDate.Value) & "_" & txtRegNo.Text & "_" & earno

    End Sub

    Private Sub RadioButtonMoreInfo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonMoreInfo.CheckedChanged
        GroupBoxProjectDetails.Visible = False
        txtProjectNumber.Text = ""
        'ComboBoxEngineers.Text = ""
    End Sub

    Private Sub RadioButtonRejected_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonRejected.CheckedChanged
        GroupBoxProjectDetails.Visible = False
        txtProjectNumber.Text = ""
        ' ComboBoxEngineers.Text = ""
    End Sub
    Public Sub fillqtydetails()

        DataGridQty.Show()
        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet

        'Dim cmSQL As SqlCommand
        ' Dim drSQL As SqlDataReader

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

    Public Sub clearqtydetails()

        DataGridQty.Show()
        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet

        'Dim cmSQL As SqlCommand
        ' Dim drSQL As SqlDataReader

        strSql = "SELECT Qty,Qty_Type, Enq_Qty_IntCode FROM ENQ_Qty_Details " & _
                "WHERE  Enq_Int_code=  '000000000000000000000'"


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


    Private Sub listloadCertificateDetails()

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
            CheckedListBoxCertificate.ValueMember = "Int_code"

        Loop


    End Sub
    Private Sub fillCertDetails()

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

    Private Sub custdetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        If Len(txtcustomerid.Text) > 4 Then

            strSQL1 = "select '', CustomerName, CustomerAddress1, CustomerAddress2,'', CustomerCity, CustomerZip,CustomerState,  CustomerCountry, CustomerContact, " & _
            "'','',CustomerContactPhone,CustomerContactFax,CustomerContactEmail,'','','','', CustomerClass3,CSR,'',CustomerClass7, FOBPoint,CustomerID,'','','','','',CustomerClass1 from FSDBBR.dbo.FS_Customer where CustomerID = '" & txtcustomerid.Text & "'"

            'ist line 9 'custclssss3 -19

            'SELECT     CustomerID, CustomerName, CustomerAddress1, CustomerAddress2, CustomerCity, CustomerState, CustomerZip, CustomerCountry, CustomerContact, 
            'CustomerContactPhone, CSR, CustomerClass1, CustomerClass3, CustomerClass7, FOBPoint
            'FROM(FS_Customer)


        Else
            strSQL1 = "select * from ENQ_New_Customers where Cust_IntCode = " & DataGridViewProjectPending.CurrentRow.Cells(20).Value.ToString & ""


        End If
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

            txtCId.Text = drSQL1.Item(24)
            txtCName.Text = drSQL1.Item(1)

            txtCustAd1.Text = drSQL1.Item(2)
            txtCustAdr2.Text = drSQL1.Item(3)
            txtCustAdr3.Text = drSQL1.Item(4)
            txtCcity.Text = drSQL1.Item(5)
            txtcustzip.Text = drSQL1.Item(6)
            txtCustState.Text = drSQL1.Item(7)
            txtCustCountry.Text = drSQL1.Item(8)
            txtContact.Text = drSQL1.Item(9)
            txtDesignation.Text = drSQL1.Item(10)
            txtDept.Text = drSQL1.Item(11)
            txtMobile.Text = drSQL1.Item(12)
            txtPhone.Text = drSQL1.Item(13)
            txtFax.Text = drSQL1.Item(14)
            txtemail.Text = drSQL1.Item(15)

            txtcustclass.Text = Trim(drSQL1.Item(19))
            txtcustcsr.Text = Trim(drSQL1.Item(20))
            txtcustisr.Text = Trim(drSQL1.Item(21))
            txtcustTSSISeg.Text = Trim(drSQL1.Item(23))
            txtcustTSSSeg.Text = Trim(drSQL1.Item(22))

            txtcustclass1.Text = Trim(drSQL1.Item(30))



        End If


    End Sub

    Private Sub GroupBoxCustDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxCustDetails.Enter

    End Sub

    Private Sub txtTSSSeg_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcustTSSSeg.TextChanged

    End Sub

    Private Sub DataGridViewItemDetail_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewItemDetail.CellContentClick

    End Sub

    Private Sub DataGridViewItemDetail_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridViewItemDetail.RowHeaderMouseClick

        txtdetailintcode.Text = DataGridViewItemDetail.CurrentRow.Cells(12).Value.ToString

        fillqtydetails()

        ClearCertDetails()
        fillCertDetails()


    End Sub
    Private Sub ClearCertDetails()


        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert As String
        'Dim b As Integer

        strsql = "Select Certificates from ENQ_Certificates "

        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        i = CheckedListBoxCertificate.Items.Count

        Do While drSQL1.Read()
            cert = drSQL1.Item(0)
            a = 0
            Do While a < i

                If Trim(cert) = Trim(CheckedListBoxCertificate.Items(a)) Then

                    CheckedListBoxCertificate.SetItemChecked(a, False)

                    a = i
                Else

                    a = a + 1

                End If


            Loop

        Loop


    End Sub
    Private Sub txtProjectNumber_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProjectNumber.Click

    End Sub

    Private Sub txtProjectNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProjectNumber.TextChanged

    End Sub
    Private Sub projectnumber()


    End Sub

    Private Sub checktransmode()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL As String

        strSQL = "select Enq_Reg_No from ENQ_Project_Approval_Status where Enq_Reg_No  = " & txtRegNo.Text & ""


        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                mode = "ADD"
            Else
                mode = "EDIT"
            End If

        End If


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'check this regnumber already existing or not.
        mode = "ADD"

        If RadioButtonAcptd.Checked = False And RadioButtonMoreInfo.Checked = False And RadioButtonNotProject.Checked = False And RadioButtonRejected.Checked = False Then
            MsgBox("Any one status should be selected before saving", vbInformation)
            Exit Sub
        End If


        checktransmode()


        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim st As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        If RadioButtonAcptd.Checked = True Then
            st = "Accepted"
        ElseIf RadioButtonMoreInfo.Checked = True Then
            st = "MoreInfoRequired"
        ElseIf RadioButtonNotProject.Checked = True Then
            st = "NotProject"
        ElseIf RadioButtonRejected.Checked = True Then
            st = "Rejected"
        End If

        curdate = System.DateTime.Now()

        SaveDocuments()

        If mode = "ADD" Then


            strsql = "insert ENQ_Project_Approval_Status values(" & txtRegNo.Text & ",'" & st & "','" & txtremarks.Text & "'," & _
            "'" & txtProjectNumber.Text & "','" & curdate & "','" & ComboBoxEngineers.Text & "','" & curdate & "','" & curdate & "', '" & username & "','" & ComboBoxProjectType.Text & "','" & txtEARNumber.Text & "','" & txtProjectName.Text & "','" & txtApplication.Text & "')"

            cnSQL.Open()

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot save details" & strsql, MsgBoxStyle.Exclamation, "Error!")
                Application.Exit()

            Else
                MsgBox("Details saved.", vbInformation)
                ButtonSave.Enabled = False
                ' listloaddocuments() 'to clear the data
                cleardocdetails()
                ClearCertDetails()
                cleardetails() 'to clear the data

                Exit Sub
            End If
        ElseIf mode = "EDIT" Then

            'updateDocuments() 'delete and save again

            strsql = "update ENQ_Project_Approval_Status  set " & _
            "Status = '" & st & "', Remarks = '" & txtremarks.Text & "', Project_Number = '" & txtProjectNumber.Text & "', " & _
            "Project_RegDate = '" & curdate & "', Project_Alotted	= '" & ComboBoxEngineers.Text & "',Date_Modify	= '" & curdate & "', UserId = '" & username & "',Project_Type = '" & ComboBoxProjectType.Text & "' where " & _
            "Enq_Reg_No = " & txtRegNo.Text & ""


            cnSQL.Open()

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot update details" & strsql, MsgBoxStyle.Exclamation, "Error!")
                Application.Exit()

            Else
                MsgBox("Details saved.", vbInformation)
                ButtonSave.Enabled = False

                cleardocdetails()
                ClearCertDetails()
                cleardetails() 'to clear the data

                Exit Sub
            End If

        End If


    End Sub
    Private Sub cleardetails()
        RadioButtonMoreInfo.Checked = False
        RadioButtonNotProject.Checked = False
        RadioButtonAcptd.Checked = False
        RadioButtonRejected.Checked = False

        txtremarks.Text = ""
        txtProjectNumber.Text = ""
        ComboBoxEngineers.Text = ""
        ComboBoxProjectType.Text = ""
    End Sub


    Private Sub SaveDocuments()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        cnSQL.Open()

        Dim i As Integer
        i = CheckedListBoxDoc.CheckedItems.Count
        Dim a As Integer
        Dim cert As String

        ' a = 1

        'it it is edit mode delete and update again


        If mode = "EDIT" Then

            strsql = " delete from ENQ_Projects_Documents_Status where Enq_Reg_No = " & txtRegNo.Text & ""

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot delete existing document details. " & strsql, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If


        End If

        Do While a < i
            cert = ""
            cert = CheckedListBoxDoc.CheckedItems.Item(a)
            a = a + 1
            curdate = System.DateTime.Now()

            strsql = "insert ENQ_Projects_Documents_Status values (" & txtRegNo.Text & ", '" & cert & "','" & curdate & "','" & curdate & "', '" & username & "' )"

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot Save document Details. " & strsql, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If


        Loop
    End Sub

    Private Sub cleardocdetails()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert As String
        'Dim b As Integer

        strsql = "Select Documents from ENQ_Project_Documents"



        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        i = CheckedListBoxDoc.Items.Count

        Do While drSQL1.Read()
            cert = drSQL1.Item(0)
            a = 0
            Do While a < i

                If cert = CheckedListBoxDoc.Items(a) Then

                    CheckedListBoxDoc.SetItemChecked(a, False)

                    a = i
                Else

                    a = a + 1

                End If


            Loop

        Loop

    End Sub

    Private Sub filldocdetails()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert As String
        'Dim b As Integer

        strsql = "Select Documents from ENQ_Projects_Documents_Status where Enq_Reg_No = " & txtRegNo.Text & " "



        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        i = CheckedListBoxDoc.Items.Count

        Do While drSQL1.Read()
            cert = drSQL1.Item(0)
            a = 0
            Do While a < i

                If cert = CheckedListBoxDoc.Items(a) Then

                    CheckedListBoxDoc.SetItemChecked(a, True)

                    a = i
                Else

                    a = a + 1

                End If


            Loop

        Loop

    End Sub


    Private Sub RadioButtonNotProject_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonNotProject.CheckedChanged
        GroupBoxProjectDetails.Visible = False
        txtProjectNumber.Text = ""
        ' ComboBoxEngineers.Text = ""
    End Sub

    Private Sub GroupBoxSelect_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxSelect.Enter

    End Sub

    Private Sub ComboBoxEngineers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBoxProjectType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CheckedListBoxDoc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBoxDoc.SelectedIndexChanged

    End Sub

    Private Sub txtremarks_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtremarks.TextChanged

    End Sub

    Private Sub ComboBoxProjectType_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxProjectType.SelectedIndexChanged

    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        'check this regnumber already existing or not.
        mode = "ADD"

        If RadioButtonAcptd.Checked = False And RadioButtonMoreInfo.Checked = False And RadioButtonNotProject.Checked = False And RadioButtonRejected.Checked = False Then
            MsgBox("Any one status should be selected before saving", vbInformation)
            Exit Sub
        End If


        checktransmode()


        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim st As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)



        If RadioButtonAcptd.Checked = True Then
            st = "Accepted"
        ElseIf RadioButtonMoreInfo.Checked = True Then
            st = "MoreInfoRequired"
        ElseIf RadioButtonNotProject.Checked = True Then
            st = "NotProject"
        ElseIf RadioButtonRejected.Checked = True Then
            st = "Rejected"
        End If


        curdate = System.DateTime.Now()

        SaveDocuments()

        If mode = "ADD" Then


            strsql = "insert ENQ_Project_Approval_Status values(" & txtRegNo.Text & ",'" & st & "','" & txtremarks.Text & "'," & _
            "'" & txtProjectNumber.Text & "','" & curdate & "','" & ComboBoxEngineers.Text & "','" & curdate & "','" & curdate & "', '" & username & "','" & ComboBoxProjectType.Text & "','" & txtEARNumber.Text & "','" & txtProjectName.Text & "','" & txtApplication.Text & "')"

            cnSQL.Open()

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot save details" & strsql, MsgBoxStyle.Exclamation, "Error!")
                Application.Exit()

            Else
                MsgBox("Details saved.", vbInformation)
                ButtonSave.Enabled = False
                ' listloaddocuments() 'to clear the data

                If RadioButtonAcptd.Checked = True Then
                    UPDATEEARNO()

                End If

                cleardocdetails()
                ClearCertDetails()
                cleardetails() 'to clear the data

                Exit Sub
            End If
        ElseIf mode = "EDIT" Then

            'updateDocuments() 'delete and save again

            strsql = "update ENQ_Project_Approval_Status  set " & _
            "Status = '" & st & "', Remarks = '" & txtremarks.Text & "', Project_Number = '" & txtProjectNumber.Text & "', " & _
            "Project_RegDate = '" & curdate & "', Project_Alotted	= '" & ComboBoxEngineers.Text & "',Date_Modify	= '" & curdate & "', UserId = '" & username & "',Project_Type = '" & ComboBoxProjectType.Text & "' where " & _
            "Enq_Reg_No = " & txtRegNo.Text & ""


            cnSQL.Open()

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot update details" & strsql, MsgBoxStyle.Exclamation, "Error!")
                Application.Exit()

            Else
                MsgBox("Details saved.", vbInformation)
                ButtonSave.Enabled = False

                cleardocdetails()
                ClearCertDetails()
                cleardetails() 'to clear the data

                Exit Sub
            End If

        End If

        'If RadioButtonAcptd.Checked = True Then
        'update ear number
        'End If


    End Sub
    Private Sub UPDATEEARNO()
        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim st As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim a As Integer
        a = Mid(txtEARNumber.Text, 15)

        ' a = Left(txtEARNumber.Text, 3)



        strsql = "update ENQ_Project_Document_Control  set " & _
            "LastUsed_No = " & a & " where Project_Doc_Name = 'EAR'"


            cnSQL.Open()

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot update EARNO details" & strsql, MsgBoxStyle.Exclamation, "Error!")
                Application.Exit()

            Else
           End If


    End Sub


    Private Sub barcode()

        ' Dim i As String = "Line0|Line1|Line2|Line3"
        'Dim a() As String
        'Dim j As Integer
        'a = i.Split("|")
        'For j = 0 To a.GetUpperBound(0)
        'MsgBox(a(j))
        'Next

        Dim i As String '= "Line0|Line1|Line2|Line3"
        i = txtDocDetails.Text
        Dim a() As String
        Dim j As Integer
        a = i.Split("$")
        For j = 0 To a.GetUpperBound(0)

            'MsgBox(a(j))
            If j = 0 Then
                txtDocDetails.Text = (a(j))
            ElseIf j = 1 Then
                txtcustomerid.Text = (a(j))
            ElseIf j = 2 Then
                txtCustomer.Text = (a(j))
            ElseIf j = 3 Then
                txtCustcity.Text = (a(j))


            ElseIf j = 4 Then
                txtCustcity.Text = (a(j))
            End If


        Next
        txtSpecial.Focus()

    End Sub

    Private Sub txtDocDetails_DoubleClick(sender As Object, e As EventArgs) Handles txtDocDetails.DoubleClick
        barcode()
    End Sub

    Private Sub txtDocDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles txtDocDetails.KeyDown
        ' barcode()
    End Sub

    Private Sub txtDocDetails_LostFocus(sender As Object, e As EventArgs) Handles txtDocDetails.LostFocus
        ' barcode()
    End Sub




    Private Sub txtDocDetails_TextChanged(sender As Object, e As EventArgs) Handles txtDocDetails.TextChanged
        ' barcode()

    End Sub
End Class