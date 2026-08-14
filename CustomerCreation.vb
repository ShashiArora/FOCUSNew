Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms

Public Class CustomerCreation
    Inherits System.Windows.Forms.Form

    Private ConnectionString As String
    Public stockDA As SqlDataAdapter = New SqlDataAdapter

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustID.TextChanged

    End Sub

    Private Sub txtCustCountry_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustCountry.TextChanged

    End Sub

    Private Sub CustomerCreation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load



    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxCustomer.Enter

    End Sub

    Private Sub DataGridCustomerCreation_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridCustomerCreation.CurrentCellChanged
        Dim b As Integer
        'Dim custid As String
        b = DataGridCustomerCreation.CurrentCell.ColumnNumber()

        If b = 0 Then
            clearcustomerdata()

            'strSQL = "SELECT  Cust_IntCode , Name, Addr1, Addr2, Addr3, City, Pin, State, Country, ContactPerson, Designation, 10
            'Dept, Mobile, Phone, Fax, Email, Ecc, Vat, " & _" 17
            '"CST, Duns_No,Remarks, Class1,Class3, CSR, ISR, TSSISeg, TSSSeg, Tax_Type,Date_Add, Date_Modify " % _29
            '"FROM ENQ_New_Customers WHERE (Len(CustomerID) < 4) ORDER BY Date_Add"

            txtCustIndCode.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 0)
            txtCustomer.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 1)
            txtCustAd1.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 2)
            txtCustAdr2.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 3)
            txtCustAdr3.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 4)
            txtCustcity.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 5)
            txtCustPin.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 6)
            txtCustState.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 7)
            txtCustCountry.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 8)
            txtContact.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 9)
            txtDesignation.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 10)
            txtDept.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 11)
            txtMobile.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 12)
            txtPhone.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 13)
            txtFax.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 14)
            txtemail.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 15)
            txtEcc.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 16)
            txtVat.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 17)
            txtCst.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 18)
            If IsDBNull(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 19)) Then
                txtDunsno.Text = ""
            Else

                txtDunsno.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 19)
            End If

            txtRemarks.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 20)
            If IsDBNull(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 21)) Then
                TxtCustClass1.Text = ""
            Else
                TxtCustClass1.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 21)
            End If

            txtCustClss3.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 22)

            txtCSR.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 23)
            txtISR.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 24)
            txtTSSISeg.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 25)
            txtTSSSeg.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 26)
            If IsDBNull(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 27)) Then
                txtTax.Text = ""
            Else
                txtTax.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 27)
            End If


            If RadioButtonCustPending.Checked = True Then
                If Trim(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 28)) = "Domestic" Then
                    RadioButtonDomestic.Checked = True
                    RadioButtonExport.Checked = False
                ElseIf Trim(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 28)) = "Export" Then
                    RadioButtonDomestic.Checked = False
                    RadioButtonExport.Checked = True

                End If
            End If

            'End If

            If RadioButtonCustCompleted.Checked = True Or RadioButtonCustAll.Checked = True Then


                If IsDBNull(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 28)) Then
                    txtCustID.Text = "-"
                Else

                    txtCustID.Text = DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 28)
                End If

                If IsDBNull(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 29)) Then
                    RadioButtonDomestic.Checked = True
                ElseIf Trim(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 29)) = "Domestic" Then
                    RadioButtonDomestic.Checked = True
                    RadioButtonExport.Checked = False
                ElseIf Trim(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 29)) = "Export" Then
                    RadioButtonDomestic.Checked = False
                    RadioButtonExport.Checked = True

                End If
                If IsDBNull(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 32)) Then

                    txtCurrency.Text = ""
                Else
                    txtCurrency.Text = Trim(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 32))
                End If


            End If
            If RadioButtonCustPending.Checked = True Then
                If IsDBNull(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 31)) Then

                    txtCurrency.Text = ""
                Else
                    txtCurrency.Text = Trim(DataGridCustomerCreation.Item(DataGridCustomerCreation.CurrentCell.RowNumber, 31))
                End If

            End If


        Else
            MsgBox("Click on Customer Internal Code", vbInformation)
            Exit Sub
        End If


    End Sub

    Private Sub DataGridCustomerCreation_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGridCustomerCreation.Navigate

    End Sub
    '  Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Private Sub ButtonCustRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCustRefresh.Click
        
        DataGridCustomerCreation.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String
        strSQL = ""
        Dim stockDC As DataSet = New DataSet

        If RadioButtonCustPending.Checked = True Then
            ButtonUpdate.Enabled = True

            '  strSQL = "SELECT  Cust_IntCode , Name, Addr1, Addr2, Addr3, City, Pin, State, Country, ContactPerson, Designation, Dept, Mobile, Phone, Fax, Email, Ecc, Vat, " & _
            ' "CST, Duns_No,Remarks, Class1,Class3, CSR, ISR, TSSISeg, TSSSeg, Tax_Type,MarketType,Date_Add, Date_Modify,Currency " & _
            '"FROM ENQ_New_Customers INNER JOIN dbo.ENQ_New_Customers ON dbo.ENQ_Header.Cust_IntCode = dbo.ENQ_New_Customers.Cust_IntCode  WHERE (Len(CustomerID) < 4) ORDER BY Date_Add"

            strSQL = "SELECT  a.Cust_IntCode , a.Name, a.Addr1, a.Addr2, a.Addr3, a.City, a.Pin, a.State, a.Country, a.ContactPerson, a.Designation, a.Dept, a.Mobile, a.Phone, a.Fax, a.Email, a.Ecc, a.Vat, " & _
            "a.CST, a.Duns_No, a.Remarks, a.Class1, a.Class3, a.CSR, a.ISR, a.TSSISeg, a.TSSSeg, a.Tax_Type, a.MarketType, a.Date_Add, a.Date_Modify, a.Currency " & _
            "FROM ENQ_New_Customers a INNER JOIN ENQ_Header b ON a.Cust_IntCode = b.Cust_IntCode  WHERE (Len(a.CustomerID) < 4) and b.Enq_Status = 'Accepted' ORDER BY Date_Add"

        ElseIf RadioButtonCustCompleted.Checked = True Then
            ButtonUpdate.Enabled = False
            strSQL = "SELECT  Cust_IntCode , Name, Addr1, Addr2, Addr3, City, Pin, State, Country, ContactPerson, Designation, Dept, Mobile, Phone, Fax, Email, Ecc, Vat, " & _
           "CST, Duns_No,Remarks, Class1,Class3, CSR, ISR, TSSISeg, TSSSeg, Tax_Type,CustomerID, MarketType,Date_Add, Date_Modify,Currency " & _
           "FROM ENQ_New_Customers WHERE (Len(CustomerID) > 4) ORDER BY Date_Add"

        ElseIf RadioButtonCustAll.Checked = True Then
            ButtonUpdate.Enabled = False
            strSQL = "SELECT  Cust_IntCode , Name, Addr1, Addr2, Addr3, City, Pin, State, Country, ContactPerson, Designation, Dept, Mobile, Phone, Fax, Email, Ecc, Vat, " & _
           "CST, Duns_No,Remarks, Class1,Class3, CSR, ISR, TSSISeg, TSSSeg, Tax_Type,CustomerID,MarketType Date_Add, Date_Modify,Currency " & _
           "FROM ENQ_New_Customers Order by Date_Add"

        End If

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridCustomerCreation.DataSource = stockDC.Tables(0)
        cnSQL.Close()
        DataGridCustomerCreation.Expand(-1)


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
        TxtCustClass1.Text = ""
        txtCustClss3.Text = ""
        txtCSR.Text = ""
        txtISR.Text = ""
        txtTSSISeg.Text = ""
        txtTSSSeg.Text = ""
        txtDunsno.Text = ""
        txtTax.Text = ""


    End Sub

    Private Sub ComboboxClass3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub GroupBox6_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox6.Enter

    End Sub

    Private Sub ButtonUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonUpdate.Click
        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        curdate = System.DateTime.Now()

        txtCustID.Text = UCase(txtCustID.Text)


        'Dim drSQL1 As SqlDataReader


        If Len(txtCustID.Text) >= 5 Then
            strsql = "update ENQ_New_Customers  set CustomerID = '" & txtCustID.Text & "',CustomerIDAdd_Date = '" & curdate & "' WHERE Cust_IntCode = " & txtCustIndCode.Text & ""

        End If

        cnSQL.Open()
        cmSQL = New SqlCommand(strsql, cnSQL)


        If cmSQL.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot Save Customer ID. " & strsql, MsgBoxStyle.Exclamation, "Error!")
            Application.Exit()

        Else
            MsgBox("Customer ID updated ", vbInformation)
            Exit Sub

        End If



    End Sub
End Class