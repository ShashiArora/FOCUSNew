Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
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





Public Class EnquiryStatus
    Inherits System.Windows.Forms.Form

    Private ConnectionString As String
    Public stockDA As SqlDataAdapter = New SqlDataAdapter
    Public ZCLASS As String
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"


    Private Sub EnquiryStatus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        RadioButtonSummary.Checked = True

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource As SqlDataAdapter = New SqlDataAdapter
        ESource.SelectCommand = sqlCmd

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
        With ComboBoxISR()
            .DataSource = source.Tables("eisr")
            .DisplayMember = "ISR"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With




        ' DataGridViewEnquiryStatus.Enabled = True

        ' Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim strSQL As String

        'Dim stockDC As DataSet = New DataSet

        'strSQL = "Select * from TSS_Enquiry_Status"

        'Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        'Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        'stockDAC.SelectCommand = sqlCmd
        'cnSQL.Open()

        'stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        'stockDAC.Fill(stockDC)

        'DataGridViewEnquiryStatus.DataSource = stockDC.Tables(0)
        'cnSQL.Close()
        
    End Sub

    Private Sub DataGridViewEnquiryStatus_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewEnquiryStatus.CellContentClick

    End Sub

    Private Sub RadioButtonDetail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonDetail.CheckedChanged
        MsgBox("Screen is not ready", MsgBoxStyle.Information)
        Exit Sub
    End Sub

    Private Sub ButtonRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRefresh.Click

        ZCLASS = "NO"

        ZCLASSHANDLING()

        DataGridViewEnquiryStatus.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        Dim stockDC As DataSet = New DataSet

        If IsDBNull(ComboBoxCSR.Text) Then
            ComboBoxCSR.Text = ""
            ComboBoxCSR.Text = "%"
        Else
            ComboBoxCSR.Text = ComboBoxCSR.Text & "%"
        End If

        If IsDBNull(ComboBoxISR.Text) Then
            ComboBoxISR.Text = ""
            ComboBoxISR.Text = "%"
        Else
            ComboBoxISR.Text = ComboBoxISR.Text & "%"
        End If

        If Len(txtcustid.Text) < 2 Then
            txtcustid.Text = ""
            txtcustid.Text = "%"
        Else
            txtcustid.Text = txtcustid.Text & "%"
        End If

        If Len(txtcustname.Text) = 0 Then
            txtcustname.Text = ""
            txtcustname.Text = "%"
        Else
            txtcustname.Text = txtcustname.Text & "%"
        End If


        If ZCLASS = "NO" Then
            If ComboBoxCSR.Text Like "%" And ComboBoxISR.Text Like "%" And txtcustid.Text Like "%" And txtcustname.Text Like "%" Then
                strSQL = "Select * from TSS_Enquiry_Status a where CSR like '" & ComboBoxCSR.Text & "' and CustomerID like '" & txtcustid.Text & "' and CustomerName like '" & txtcustname.Text & "' " & _
                "AND [Reg.Date] >= '" & dtpFromDate.Value & "' and [Reg.Date] <= '" & dtpToDate.Value & "'"
            Else
                strSQL = "Select * from TSS_Enquiry_Status a where CSR like '" & ComboBoxCSR.Text & "' and CustomerID like '" & txtcustid.Text & "' and CustomerName like '" & txtcustname.Text & "' " & _
                "AND [Reg.Date] >= '" & dtpFromDate.Value & "' and [Reg.Date] <= '" & dtpToDate.Value & "' and a.CSR IN (select CSR from ENQ_CSR where CSR = a.CSR and ISR LIKE '" & ComboBoxISR.Text & "') AND a.Class3 IN('K','I')"
            End If
        ElseIf ZCLASS = "YES" Then
            If ComboBoxCSR.Text Like "%" And ComboBoxISR.Text Like "%" And txtcustid.Text Like "%" And txtcustname.Text Like "%" Then
                strSQL = "Select * from TSS_Enquiry_Status a where CSR like '" & ComboBoxCSR.Text & "' and CustomerID like '" & txtcustid.Text & "' and CustomerName like '" & txtcustname.Text & "' " & _
                "AND [Reg.Date] >= '" & dtpFromDate.Value & "' and [Reg.Date] <= '" & dtpToDate.Value & "'"
            Else
                strSQL = "Select * from TSS_Enquiry_Status a where CSR like '" & ComboBoxCSR.Text & "' and CustomerID like '" & txtcustid.Text & "' and CustomerName like '" & txtcustname.Text & "' " & _
            "AND [Reg.Date] >= '" & dtpFromDate.Value & "' and [Reg.Date] <= '" & dtpToDate.Value & "' and a.CSR IN (select CSR from ENQ_CSR where CSR = a.CSR and (ZCLASS1 like '" & ComboBoxISR.Text & "' OR ZCLASS2 like '" & ComboBoxISR.Text & "' )) AND  a.Class3 NOT IN('K','I')"
            End If
        End If

            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAC.TableMappings.Add("Table", "Enq")
            'get data
            stockDAC.Fill(stockDC)

            DataGridViewEnquiryStatus.DataSource = stockDC.Tables(0)
        cnSQL.Close()

        Dim i As Integer


        For i = 0 To DataGridViewEnquiryStatus.Rows.Count - 1

            'MsgBox(Trim(DataGridViewEnquiryStatus.Rows(i).Cells(8).Value))

            If Trim(DataGridViewEnquiryStatus.Rows(i).Cells(8).Value) = "Rejected" Or Trim(DataGridViewEnquiryStatus.Rows(i).Cells(8).Value) = "Closed" Then

                DataGridViewEnquiryStatus.Rows(i).DefaultCellStyle.BackColor = Color.LightBlue

            ElseIf IsDBNull(DataGridViewEnquiryStatus.Rows(i).Cells(10).Value) And (Trim(DataGridViewEnquiryStatus.Rows(i).Cells(8).Value) = "Pending" Or Trim(DataGridViewEnquiryStatus.Rows(i).Cells(8).Value) = "Accepted") Then


                'ElseIf (Trim(DataGridViewEnquiryStatus.Rows(i).Cells(8).Value) = "Pending" Or Trim(DataGridViewEnquiryStatus.Rows(i).Cells(8).Value) <> "Accepted") And (DataGridViewEnquiryStatus.Rows(i).Cells(10).Value) Is Null Then

                DataGridViewEnquiryStatus.Rows(i).DefaultCellStyle.BackColor = Color.Magenta



            End If

        Next

            ' DataGridCustomerCreation.Expand(-1)

    End Sub
    Sub fillcustomerlist()

        
        DataGridCustomer.Show()



        Dim sqlcon As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        txtCustID.Text = txtCustID.Text & "%"

        If Len(Trim(txtcustid.Text)) > 1 Then
            strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.FS_Customer " & _
                 "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRE%') AND (CustomerID NOT LIKE '0000%')AND CustomerID like '" & txtcustid.Text & "' " & _
                    "ORDER BY CustomerID"

        ElseIf Len(Trim(txtcustname.Text)) > 1 Then

            strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.FS_Customer " & _
                     "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRE%') AND (CustomerID NOT LIKE '0000%')AND CustomerName like '" & txtcustname.Text & "' " & _
                        "ORDER BY CustomerName"

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

        DataGridCustomer.DataSource = stockDC.Tables(0)
        sqlcon.Close()
        DataGridCustomer.Expand(-1)


    End Sub

    Private Sub DataGridCustomer_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridCustomer.CurrentCellChanged
        Dim a As Integer
     

        a = DataGridCustomer.CurrentCell.ColumnNumber()

        If a = 0 Then
            txtcustid.Text = DataGridCustomer.Item(DataGridCustomer.CurrentCell)

            txtcustname.Text = DataGridCustomer.Item(DataGridCustomer.CurrentCell.RowNumber, 1)

     
            txtcustid.Enabled = False



        Else
            MsgBox("Click on CustomerID to select the customer", vbInformation)
            Exit Sub
        End If


        DataGridCustomer.Hide()

    End Sub



    Private Sub DataGridCustomer_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridCustomer.Navigate

    End Sub


    Private Sub txtcustid_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtcustid.DoubleClick
        DataGridCustomer.Visible = True
        fillcustomerlist()
    End Sub

    Private Sub txtcustid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcustid.TextChanged

    End Sub

    Private Sub txtcustname_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtcustname.DoubleClick
        DataGridCustomer.Visible = True
        fillcustomerlist()
    End Sub

    Private Sub txtcustname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcustname.TextChanged

    End Sub

    Private Sub ZCLASSHANDLING()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select ZCLASS1,ZCLASS2 from ENQ_CSR where ZCLASS1 LIKE '" & ComboBoxISR.Text & "' or ZCLASS2 LIKE '" & ComboBoxISR.Text & "'"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        'drSQL1 = cmSQL1.ExecuteReader
        'drSQL1 = cmSQL1.ExecuteReader
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                ZCLASS = "NO"
            Else

                ZCLASS = "YES"
            End If

        Else
            ZCLASS = "NO"
        End If


    End Sub

End Class