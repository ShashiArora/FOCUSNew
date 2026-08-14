Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports outlook = Microsoft.Office.Interop.Outlook
'Imports SoftBrands.FourthShift.Transaction
'Imports CrystalDecisions.CrystalReports.Engine
Imports System.Windows.Forms

Public Class RFQHistory
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"



    Private Sub RFQHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        datagridviewHistory.Enabled = True

        'RBToolNo.Checked = True
        'RadioButtonGroup.Checked = True
        'RadioButtonVendorYes.Checked = True


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        strSQL = "SELECT     RegNo, [Reg.Date], CustomerID, CustomerName, CSR, SlNo, Part_Source, PartNumber, PartDescription, Class1, Qty, Price, Qty_Type, Factor, FinalPrice, Source_Mtrl, " & _
            " MOQ, SPU, LeadTime, Type, Stock_Avble, Vendor_Ref, Name, Vendor_Quote, Special_Remarks FROM         TSS_Enquiry_Price_Completed_QtyPrice where PartNumber = '" & parthistory & "' order by RegNo,SlNo"





        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)



        datagridviewHistory.DataSource = stockDC.Tables(0)




    End Sub

    Private Sub datagridviewHistory_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridviewHistory.CellContentClick

    End Sub
End Class