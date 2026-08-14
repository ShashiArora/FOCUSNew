Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
Imports System.Reflection



Public Class MyCustomer
    ' Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"
    Private Sub MyCustomer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridViewMyCustomer.Enabled = True


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet


        strSQL = "SELECT CustomerID,CustomerName,CustomerCity,CSR,INS_SALES_CDE AS ISR,[TSSI-Seg],[TSS-Seg] FROM FSPrograms.dbo.TSS_CUSTOMERID_ISR where INS_SALES_CDE = '" & username & "' "



        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewMyCustomer.DataSource = stockDC.Tables(0)

        Dim col1 As DataGridViewColumn = DataGridViewMyCustomer.Columns(1) 'CUSTOMER NAME
        col1.Width = 300

        Dim col2 As DataGridViewColumn = DataGridViewMyCustomer.Columns(2) 'CUSTOMER CITY
        col2.Width = 150



        cnSQL.Close()



    End Sub
End Class