'
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



Public Class Pricing
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Public iPosition As Integer
    Public partnum As String

    Dim mode As String
    Dim quote As String
    Dim quotedate As Date
    Dim checkstatus As String
    Dim checklost As String

    Dim seq As String















    Private Sub ButtonEnquiryDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSalesHistory.Click
        ' If CheckBoxMainPart.Checked = True Then
        'partnum = txtPartNumber.Text
        'ElseIf CheckBoxAltPart.Checked = True Then
        'partnum = txtPartNumberAlt.Text
        'End If

        If GroupBoxPriceSuggestion.Visible = True Then
            MsgBox("Close the Price Calculation window, before clicking on Sales Data", vbInformation)
            Exit Sub

        End If


        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be entered", vbInformation)
            Exit Sub
        End If


        If Val(txtSales.Text) = 0 Then
            MsgBox("No. of months need to be entered", vbInformation)
            Exit Sub

        End If



        If GroupBoxAlternativeMaterial.Visible = True Then

            GroupBoxAlternativeMaterial.Visible = False

        End If

        txtmax.Text = ""
        txtmin.Text = ""
        txtnotional.Text = ""
        txtavg.Text = ""
        txtitmc.Text = ""

        cleardatagridview()



        DataGridViewSalesHistory.Enabled = True


        'part number and description display

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader

        strSQL = "select ItemNumber, ItemDescription from  [FSDBBR].[dbo].FS_Item WHERE ItemNumber like  '" & partnum & "'"



        cnSQL.Open()
        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()



        If drSQL.Read() Then

            If IsDBNull(drSQL.Item(0)) Then
                lblPartDescription.Text = ""
            Else

                '        lblPartDescription.Text = drSQL.Item(0) & " " & drSQL.Item(1)
                lblPartDescription.Text = drSQL.Item(1)

            End If
        End If

        cnSQL.Close()


        Dim periodyearS As Integer
        periodyearS = Val(txtSales.Text)
        ' periodyear = 2 * 365


        Dim d As Date = Date.Today

        ' d = d.AddDays(-periodyear)

        d = d.AddMonths(-periodyearS)
        ' d = Format(d, "mm/dd/yyyy")



        'notional price
        Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL3 As SqlCommand
        Dim drSQL3 As SqlDataReader
        Dim strSQL3 As String


        '  strSQL3 = "Select Notional_Price  FROM [FSDBBR].[dbo].[TSS_Notional_price] where Part_no = '" & txtPartNumber.Text & "' and [Year] = " & Year(Now) & ""

        strSQL3 = "Select Notional_Price,Part_no, max(Year)  FROM [FSDBBR].[dbo].[TSS_Notional_price] where Part_no = '" & partnum & "' group by  Notional_Price,Part_no"


        cnSQL3.Open()
        cmSQL3 = New SqlCommand(strSQL3, cnSQL3)
        drSQL3 = cmSQL3.ExecuteReader()

        If drSQL3.Read() Then

            If IsDBNull(drSQL3.Item(0)) Then
                txtnotional.Text = 0
            Else


                txtnotional.Text = Format(drSQL3.Item(0), "0.00")
            End If


        End If

        cnSQL3.Close()




        'main sales data

        Dim stockDC As DataSet = New DataSet


        If usertype = "S" Then

            If Val(txtnotional.Text) > 0 Then

                strSQL = "SELECT max(TransactionDate) as Trans_Date, type as Type, CustomerID as Cust_ID, max(CustomerName) as Customer_Name, max(CustomerClass)  as Class, sum(ShipQty) as Ship_Qty, round(count(ShipQty),0) as 'Instances' ,UnitPrice,  " & txtnotional.Text & " as 'NotionalCost' FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' group by CustomerID, UnitPrice, type order by Trans_Date Desc"

            Else

                strSQL = "SELECT max(TransactionDate) as Trans_Date, type as Type, CustomerID as Cust_ID, max(CustomerName) as Customer_Name, max(CustomerClass)  as Class, sum(ShipQty) as Ship_Qty, round(count(ShipQty),0) as 'Instances' ,  UnitPrice, Avg(Cost) as AvgCost FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' group by CustomerID, UnitPrice, type order by Trans_Date Desc"
            End If

        ElseIf usertype = "Q" Then


            If Val(txtnotional.Text) > 0 Then

                strSQL = "SELECT max(TransactionDate) as Trans_Date, type as Type, CustomerID as Cust_ID, max(CustomerName) as Customer_Name, max(CustomerClass)  as Class, sum(ShipQty) as Ship_Qty, round(count(ShipQty),0) as 'Instances' ,UnitPrice FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' group by CustomerID, UnitPrice, type order by Trans_Date Desc"

            Else

                strSQL = "SELECT max(TransactionDate) as Trans_Date, type as Type, CustomerID as Cust_ID, max(CustomerName) as Customer_Name, max(CustomerClass)  as Class, sum(ShipQty) as Ship_Qty, round(count(ShipQty),0) as 'Instances' ,  UnitPrice FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' group by CustomerID, UnitPrice, type order by Trans_Date Desc"
            End If







        End If








        '        strSQL = "SELECT a.TransactionDate as Trans_Date, a.type, a.CustomerID as Cust_ID, a.CustomerName as Customer_Name, a.CustomerClass as Class, a.ShipQty as Ship_Qty, count(a.ShipQty) as 'Instances' ,  UnitPrice, Cost from [FSDBBR].[dbo].[TSS_Price_Order_History] a  where a.TransactionDate in " & _
        '           "(select max(TransactionDate) from  FROM  [FSDBBR].[dbo].[TSS_Price_Order_History ] b  where b.ItemNumber = '" & partnum & "' and  b.TransactionDate >= '" & d & "' and b.CustomerID = a.CustomerID  AND [b.type] = [a.type] ) and " & _
        '          " a.ItemNumber = '" & partnum & "' and  b.TransactionDate >= '" & d & "'"
        'as Cost FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' group by CustomerID, UnitPrice order by Trans_Date Desc"






        '   strSQL = "SELECT [SALES_AMT],[COST_AMT],[SALES_QTY],MarPer, INV_DATE FROM [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where MarPer IN " & _
        '           " (SELECT MAX(MarPer) from [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where ORDER_TYPE = 'INVOICE' AND "
        'INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ADDR_NBR <> '" & txtCustID.Text & "' and ItemNumber LIKE '" & partnum & "' " & _
        '                " and ([INV_DATE] >= '" & d & "' )) and ORDER_TYPE = 'INVOICE' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and  
        'ADDR_NBR <> '" & txtCustID.Text & "' AND   ItemNumber LIKE '" & partnum & "' and ([INV_DATE] >=  '" & d & "') ORDER BY MarPer desc "




        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewSalesHistory.DataSource = stockDC.Tables(0)

        'sales history column setting

        Dim InvoiceDate As DataGridViewColumn = DataGridViewSalesHistory.Columns(0) ' Trsndate
        InvoiceDate.Width = 80
        DataGridViewSalesHistory.Columns.Item(0).DefaultCellStyle.Format = "MM/dd/yyyy" ' "dd/MM/yyyy"

        Dim Type As DataGridViewColumn = DataGridViewSalesHistory.Columns(1) 'Type
        Type.Width = 50


        Dim CustomerID As DataGridViewColumn = DataGridViewSalesHistory.Columns(2) 'custid
        CustomerID.Width = 70

        Dim CustomerName As DataGridViewColumn = DataGridViewSalesHistory.Columns(3) 'custname
        CustomerName.Width = 310

        Dim CustomerClass1 As DataGridViewColumn = DataGridViewSalesHistory.Columns(4) 'custclass
        CustomerClass1.Width = 60

        Dim ShipQty As DataGridViewColumn = DataGridViewSalesHistory.Columns(5) 'shipqty
        ShipQty.Width = 100

        DataGridViewSalesHistory.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewSalesHistory.Columns.Item(5).ValueType = GetType(Double)
        'DataGridViewSalesHistory.Columns.Item(4).DefaultCellStyle.Format = "n2"

        Dim Instaces As DataGridViewColumn = DataGridViewSalesHistory.Columns(6) 'Instances
        Instaces.Width = 100

        DataGridViewSalesHistory.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        ' DataGridViewSalesHistory.Columns.Item(5).ValueType = GetType(Double)
        'DataGridViewSalesHistory.Columns.Item(5).DefaultCellStyle.Format = "n2"

        Dim InvoiceLocalUnitPrice As DataGridViewColumn = DataGridViewSalesHistory.Columns(7) 'unit price
        InvoiceLocalUnitPrice.Width = 130

        DataGridViewSalesHistory.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewSalesHistory.Columns.Item(7).ValueType = GetType(Double)
        DataGridViewSalesHistory.Columns.Item(7).DefaultCellStyle.Format = "n2"

        If usertype = "S" Then


            If Val(txtnotional.Text) > 0 Then
                Dim NotionalCost As DataGridViewColumn = DataGridViewSalesHistory.Columns(8) 'cost
                NotionalCost.Width = 110

            Else

                Dim cost As DataGridViewColumn = DataGridViewSalesHistory.Columns(8) 'cost
                cost.Width = 110
            End If

            DataGridViewSalesHistory.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridViewSalesHistory.Columns.Item(8).ValueType = GetType(Double)
            DataGridViewSalesHistory.Columns.Item(8).DefaultCellStyle.Format = "n2"

        End If


        'end of col setting

        cnSQL.Close()

        'minimum price
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "SELECT UnitPrice, sum(ShipQty) as ShipQty, COUNT(*) as Instances FROM [FSDBBR].[dbo].[TSS_Price_Order_History] where UnitPrice in (select min(UnitPrice) from [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "') and ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "' group by CustomerID ,UnitPrice "

        ' SELECT UnitPrice, sum(ShipQty)ShipQty,COUNT(*)soldtimes
        'FROM [FSDBBR].[dbo].[TSS_Price_Order_History] where UnitPrice in (select min(UnitPrice) from [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = 'ORAR00015-N90' and [TransactionDate] >= '2/12/2014')
        'and ItemNumber = 'ORAR00015-N90' and [TransactionDate] >= '10-12-2005'
        'group by UnitPrice 

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtmin.Text = 0
            Else
                txtmin.Text = Format(drSQL1.Item(0), "0.00")
            End If


            If IsDBNull(drSQL1.Item(1)) Then
                txtminqty.Text = 0
            Else
                txtminqty.Text = drSQL1.Item(1)
            End If

            If IsDBNull(drSQL1.Item(2)) Then
                txtmininstances.Text = 0
            Else
                txtmininstances.Text = drSQL1.Item(2)
            End If

        End If

        cnSQL1.Close()

        'maximum price


        strSQL1 = "SELECT UnitPrice, sum(ShipQty) as ShipQty, COUNT(*) as Instances FROM [FSDBBR].[dbo].[TSS_Price_Order_History] where UnitPrice in (select max(UnitPrice) from [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "') and ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "' group by CustomerID ,UnitPrice "


        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtmax.Text = 0
            Else
                txtmax.Text = Format(drSQL1.Item(0), "0.00")
            End If


            If IsDBNull(drSQL1.Item(1)) Then
                txtmaxqty.Text = 0
            Else
                txtmaxqty.Text = drSQL1.Item(1)
            End If

            If IsDBNull(drSQL1.Item(2)) Then
                txtmaxinstances.Text = 0
            Else
                txtmaxinstances.Text = drSQL1.Item(2)
            End If

        End If


        'end of maximum price


        'cost

        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL2 As SqlCommand
        Dim drSQL2 As SqlDataReader
        Dim strSQL2 As String


        strSQL2 = "SELECT TotalRolledCost FROM  [FSDBBR].[dbo].[TSS_ITEM_COST] where ItemNumber = '" & partnum & "'"
        cnSQL2.Open()
        cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
        drSQL2 = cmSQL2.ExecuteReader()

        If drSQL2.Read() Then

            If IsDBNull(drSQL2.Item(0)) Then
                txtitmc.Text = 0
            Else

                txtitmc.Text = Format(drSQL2.Item(0), "0.00")
            End If


        End If

        cnSQL2.Close()


        ''notional price
        'Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cmSQL3 As SqlCommand
        'Dim drSQL3 As SqlDataReader
        'Dim strSQL3 As String


        ''  strSQL3 = "Select Notional_Price  FROM [FSDBBR].[dbo].[TSS_Notional_price] where Part_no = '" & txtPartNumber.Text & "' and [Year] = " & Year(Now) & ""

        'strSQL3 = "Select Notional_Price,Part_no, max(Year)  FROM [FSDBBR].[dbo].[TSS_Notional_price] where Part_no = '" & partnum & "' group by  Notional_Price,Part_no"


        'cnSQL3.Open()
        'cmSQL3 = New SqlCommand(strSQL3, cnSQL3)
        'drSQL3 = cmSQL3.ExecuteReader()

        'If drSQL3.Read() Then

        '    If IsDBNull(drSQL3.Item(0)) Then
        '        txtnotional.Text = 0
        '    Else


        '        txtnotional.Text = Format(drSQL3.Item(0), "0.00")
        '    End If


        'End If

        'cnSQL3.Close()


        'call stock 

        Dim cnSQL4 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL4 As SqlCommand
        Dim drSQL4 As SqlDataReader
        Dim strSQL4 As String

        strSQL4 = "SELECT  SUM([InventoryQuantity])  FROM [FSDBBR].[dbo].[TSS_Price_MarketingInventory] where ItemNumber = '" & partnum & "'"


        cnSQL4.Open()
        cmSQL4 = New SqlCommand(strSQL4, cnSQL4)
        drSQL4 = cmSQL4.ExecuteReader()

        If drSQL4.Read() Then

            If IsDBNull(drSQL4.Item(0)) Then
                txtcurstock.Text = 0

            Else

                txtcurstock.Text = Format(drSQL4.Item(0), "0.00")
                txtcurstock.Text = Val(txtcurstock.Text)
            End If


        End If

        cnSQL4.Close()

        'end of stock


        ' available to promise

        Dim cnSQL5 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL5 As SqlCommand
        Dim drSQL5 As SqlDataReader
        Dim strSQL5 As String

        strSQL5 = "SELECT  SUM([Pending Qty]) FROM [FSDBBR].[dbo].[TSS_PendingSalesOrders_ver5] where COLineStatus in (3,4) and (CustomerName not like 'TSS%' AND CustomerName not like 'TRE%')  AND ItemNumber = '" & partnum & "'"

        'Required for kit to be done.

        cnSQL5.Open()
        cmSQL5 = New SqlCommand(strSQL5, cnSQL5)
        drSQL5 = cmSQL5.ExecuteReader()

        If drSQL5.Read() Then

            If IsDBNull(drSQL5.Item(0)) Then
                txtAvblePromise.Text = 0

                txtAvblePromise.Text = Val(txtcurstock.Text) - Val(txtAvblePromise.Text)


            Else

                txtAvblePromise.Text = Format(drSQL5.Item(0), "0.00")

                txtAvblePromise.Text = Val(txtcurstock.Text) - Val(txtAvblePromise.Text)

            End If


        End If

        cnSQL5.Close()

        'end of available to promised


        '     cnSQL1.Close()
        '    cnSQL2.Close()
        '   cnSQL3.Close()


    End Sub
    '    Private Sub EnqDetails()

    '        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
    '        Dim cmSQL1 As SqlCommand
    '        Dim drSQL1 As SqlDataReader
    '        Dim strSQL1 As String

    '        strSQL1 = "SELECT RegNo,  [Reg.Date], CustomerID, CustomerName, City, Class,Class1, Cust_Exist_New as Exis_Cust, CSR, TSSISeg, TSSSeg,MarketType, " & _
    '        " Enq_Ref_no, Enq_Ref_date, " & _
    '              "Enq_Source, Enq_Recd_date,Doc_upload,Doc_Details,Special_instructions from TSS_Enq_Pending_Project_Aproval where RegNo = " & txtEnqRegNo.Text & ""

    '        cnSQL1.Open()
    '        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
    '        drSQL1 = cmSQL1.ExecuteReader()

    '        If drSQL1.Read() Then

    '            txtRegNo1.Text = drSQL1.Item(0)
    '            DtpRegDate1.Value = drSQL1.Item(1)
    '            txtcustomerid1.Text = drSQL1.Item(2)
    '            txtcustomer1.Text = drSQL1.Item(3)
    '            txtcustcity1.Text = drSQL1.Item(4)
    '            txtEnqRef1.Text = drSQL1.Item(12)

    '            dtpActionStartDate1.Value = drSQL1.Item(13)

    '            dtpActionStartDate1.Format = DateTimePickerFormat.Custom
    '            dtpActionStartDate1.CustomFormat = "MMM yyyy"



    '            txtEnqSource1.Text = drSQL1.Item(14)
    '            DTPEnqRecd1.Value = drSQL1.Item(15)


    '            If Trim(drSQL1.Item(16)) = "YES" Then
    '                rbdocyes1.Checked = True
    '                rbdocno1.Checked = False
    '            Else
    '                rbdocno1.Checked = True
    '                rbdocyes1.Checked = True
    '            End If

    '            txtdocdetails1.Text = drSQL1.Item(17)
    '            txtspecial1.Text = drSQL1.Item(18)

    '        End If


    '    End Sub

    '    Private Sub dtpEnqDt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    End Sub

    '    Private Sub BtnCustClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        GroupBoxEnqDetails1.Visible = False
    '    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonQuoteHistory.Click



        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be enteted", vbInformation)
            Exit Sub
        End If

        If Val(txtQuote.Text) = 0 Then
            MsgBox("No. of months need to be entered", vbInformation)
            Exit Sub

        End If

        DataGridViewQuoteHistory.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        '  mode = "ADD"

        Dim periodyearQ As Integer
        periodyearQ = Val(txtQuote.Text)

        Dim d As Date = Date.Today

        d = d.AddMonths(-periodyearQ)



        '        Dim periodyear As Integer
        '       periodyear = 2
        '      periodyear = 2 * 365


        'Dim d As Date = Date.Today
        'd = d.AddDays(-365)


        Dim stockDC As DataSet = New DataSet
        'strSQL = "SELECT CustomerID,CustomerName,  InvoiceDate, ShipQuantity,InvoiceLocalUnitPrice,CustomerClass1 FROM [FSDBBR].[dbo].[TSS_IVIE_IBS_SALES_COMBINE_2015Onwards] where LineItemNumber = '" & txtPartNumber.Text & "' and  InvoiceDate >= '" & d & "' order by InvoiceDate Desc"

        strSQL = " SELECT  COCreatedDate as QuoteDate,  CustomerID, CustomerName, CustomerClass1  as CustomerClass, ItemOrderedQuantity as QuoteQty, Rate as UnitPrice FROM  [FSDBBR].[dbo].[TSS_Quotes]  where ItemNumber = '" & partnum & "' and  COCreatedDate >= '" & d & "' order by COCreatedDate Desc"


        ' strSQL = "SELECT InvoiceDate, CustomerID, CustomerName, CustomerClass1  as CustomerClass, ShipQuantity, InvoiceLocalUnitPrice AS UnitPrice FROM [FSDBBR].[dbo].[TSS_IVIE_IBS_SALES_COMBINE_2015Onwards] where LineItemNumber = '" & txtPartNumber.Text & "' and  InvoiceDate >= '" & d & "' order by InvoiceDate Desc"


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewQuoteHistory.DataSource = stockDC.Tables(0)
        'cnSQL.Close()

        'txtmin.Text = DataGridViewQuoteHistory.Rows.Count

        'Quote history column setting

        Dim COCreatedDate As DataGridViewColumn = DataGridViewQuoteHistory.Columns(0)
        COCreatedDate.Width = 120

        Dim CustomerID As DataGridViewColumn = DataGridViewQuoteHistory.Columns(1)
        CustomerID.Width = 120

        Dim CustomerName As DataGridViewColumn = DataGridViewQuoteHistory.Columns(2)
        CustomerName.Width = 360

        Dim CustomerClass1 As DataGridViewColumn = DataGridViewQuoteHistory.Columns(3)
        CustomerClass1.Width = 120

        Dim ItemOrderedQuantity As DataGridViewColumn = DataGridViewQuoteHistory.Columns(4)
        ItemOrderedQuantity.Width = 130

        DataGridViewQuoteHistory.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewQuoteHistory.Columns.Item(4).ValueType = GetType(Double)
        DataGridViewQuoteHistory.Columns.Item(4).DefaultCellStyle.Format = "n2"

        Dim Rate As DataGridViewColumn = DataGridViewQuoteHistory.Columns(5)
        Rate.Width = 130
        DataGridViewQuoteHistory.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewQuoteHistory.Columns.Item(5).ValueType = GetType(Double)
        DataGridViewQuoteHistory.Columns.Item(5).DefaultCellStyle.Format = "n2"







        'end of col setting








        cnSQL.Close()




    End Sub







    Private Sub ButtonPurHistory_Click(sender As Object, e As EventArgs) Handles ButtonPurHistory.Click

        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be enteted", vbInformation)
            Exit Sub
        End If



        If Val(txtPur.Text) = 0 Then
            MsgBox("No. of months need to be entered", vbInformation)
            Exit Sub

        End If

        DataGridViewPurchaseHistory.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String



        Dim periodyearP As Integer
        periodyearP = Val(txtPur.Text)


        Dim d As Date = Date.Today

        d = d.AddMonths(-periodyearP)

        Dim stockDC As DataSet = New DataSet

        strSQL = "SELECT  [TransactionDate], Type, [VendorID], [VendorName],  [RecQty] as ReceiptQty , [UnitPrice] as PurPrice, ([UnitPrice]+  [LCM]) as LandingCost FROM [FSDBBR].[dbo].[TSS_Price_PurchaseDetails] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' order by TransactionDate Desc"


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewPurchaseHistory.DataSource = stockDC.Tables(0)
        'cnSQL.Close()

        'purchase history column setting

        Dim TransDate As DataGridViewColumn = DataGridViewPurchaseHistory.Columns(0)
        TransDate.Width = 120

        Dim Type As DataGridViewColumn = DataGridViewPurchaseHistory.Columns(1)
        Type.Width = 80


        Dim VendorID As DataGridViewColumn = DataGridViewPurchaseHistory.Columns(2)
        VendorID.Width = 100

        Dim VendorName As DataGridViewColumn = DataGridViewPurchaseHistory.Columns(3)
        VendorName.Width = 300

        Dim Recpt As DataGridViewColumn = DataGridViewPurchaseHistory.Columns(4)
        Recpt.Width = 120

        DataGridViewPurchaseHistory.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewPurchaseHistory.Columns.Item(4).ValueType = GetType(Double)
        DataGridViewPurchaseHistory.Columns.Item(4).DefaultCellStyle.Format = "n2"



        Dim PurPrice As DataGridViewColumn = DataGridViewPurchaseHistory.Columns(5)
        PurPrice.Width = 130
        DataGridViewPurchaseHistory.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewPurchaseHistory.Columns.Item(5).ValueType = GetType(Double)
        DataGridViewPurchaseHistory.Columns.Item(5).DefaultCellStyle.Format = "n2"

        Dim LandingCost As DataGridViewColumn = DataGridViewPurchaseHistory.Columns(6)
        LandingCost.Width = 130

        DataGridViewPurchaseHistory.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewPurchaseHistory.Columns.Item(6).ValueType = GetType(Double)
        DataGridViewPurchaseHistory.Columns.Item(6).DefaultCellStyle.Format = "n2"


        'end of col setting


        cnSQL.Close()

        'lowest qty

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "SELECT ([UnitPrice]+  [LCM]) as LandingCost, RecQty, TransactionDate FROM [FSDBBR].[dbo].[TSS_Price_PurchaseDetails] where RecQty in (select min(RecQty) from [FSDBBR].[dbo].[TSS_Price_PurchaseDetails] where ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "') and ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "'"

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtpurlowcost.Text = 0
            Else
                txtpurlowcost.Text = Format(drSQL1.Item(0), "0.00")
            End If


            If IsDBNull(drSQL1.Item(1)) Then
                txtpurlowqty.Text = 0
            Else
                txtpurlowqty.Text = drSQL1.Item(1)
            End If

            If IsDBNull(drSQL1.Item(2)) Then
                txtpurlowdate.Text = 0
            Else
                txtpurlowdate.Text = (drSQL1.Item(2))
            End If

        End If

        cnSQL1.Close()

        'end of lowest qty

        'highest qty

        strSQL1 = "SELECT ([UnitPrice]+  [LCM]) as LandingCost, RecQty, TransactionDate FROM [FSDBBR].[dbo].[TSS_Price_PurchaseDetails] where RecQty in (select max(RecQty) from [FSDBBR].[dbo].[TSS_Price_PurchaseDetails] where ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "') and ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "'"

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtpurhighcost.Text = 0
            Else
                txtpurhighcost.Text = Format(drSQL1.Item(0), "0.00")
            End If


            If IsDBNull(drSQL1.Item(1)) Then
                txtpurhighqty.Text = 0
            Else
                txtpurhighqty.Text = drSQL1.Item(1)
            End If

            If IsDBNull(drSQL1.Item(2)) Then
                txtpurhighdate.Text = ""
            Else
                txtpurhighdate.Text = (drSQL1.Item(2))
            End If

        End If

        cnSQL1.Close()

        'end of highest qty

        'latest qty
        strSQL1 = "SELECT ([UnitPrice]+  [LCM]) as LandingCost, RecQty as ReceiptQty, TransactionDate  FROM [FSDBBR].[dbo].[TSS_Price_PurchaseDetails] where TransactionDate in (select max(TransactionDate) from [FSDBBR].[dbo].[TSS_Price_PurchaseDetails] where ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "') and ItemNumber =  '" & partnum & "' and [TransactionDate] >= '" & d & "' "

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtpurlatestCost.Text = 0
            Else
                txtpurlatestCost.Text = Format(drSQL1.Item(0), "0.00")
            End If


            If IsDBNull(drSQL1.Item(1)) Then
                txtpurlatestqty.Text = 0
            Else
                txtpurlatestqty.Text = drSQL1.Item(1)
            End If

            If IsDBNull(drSQL1.Item(2)) Then
                txtpurlatestDate.Text = ""
            Else
                txtpurlatestDate.Text = (drSQL1.Item(2))
            End If

        End If

        If usertype = "QP" Then
            itmcdetails()

        End If

        'end of latest qty

    End Sub

    Private Sub GroupBoxSelect_Enter(sender As Object, e As EventArgs) Handles GroupBoxSelect.Enter

    End Sub

    Private Sub Pricing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtSales.Text = 12
        txtQuote.Text = 12
        txtPur.Text = 12

        CheckBoxMainPart.Checked = True
        CheckBoxAltPart.Checked = False


        If Len(txtPartNumberAlt.Text) > 5 Then
            CheckBoxAltPart.Checked = True
            CheckBoxMainPart.Checked = False
        End If

        If usertype = "Q" Then

            GroupBox2.Visible = False
            GroupBoxITMC.Visible = False
            DataGridViewPurchaseHistory.Visible = False
            Label60.Visible = False
            TextBox63.Visible = False
            txtnotional.Visible = False
            txtPur.Visible = False
            ButtonPurHistory.Visible = False
            Buttonsugprice.Visible = False

        ElseIf usertype = "QP" Then

            GroupBox10.Enabled = False
            GroupBox11.Enabled = False
            TextBox63.Enabled = False
            txtnotional.Enabled = False
            ButtonSalesHistory.Enabled = False
            ButtonQuoteHistory.Enabled = False
            txtSales.Enabled = False
            txtQuote.Enabled = False
        End If


        'End If


    End Sub

    Private Sub ButtonAlternativeDetails_Click(sender As Object, e As EventArgs) Handles ButtonAlternativeDetails.Click

        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be enteted", vbInformation)
            Exit Sub
        End If


        ' select ItemNumber, ItemDescription from FS_Item WHERE ItemNumber like 'ORAR%'

        'ALTERNATIVE MATERIAL

        If GroupBoxAlternativeMaterial.Visible = False Then

            GroupBoxAlternativeMaterial.Visible = True
            ' GroupBoxAlternativeMaterial.BringToFront()

            If GroupBoxPriceSuggestion.Visible = True Then
                GroupBoxPriceSuggestion.Visible = False
                GroupBoxSugSummary.Visible = False
            End If

            ' GroupBoxEnqDetails1.Location = New Point(6, 7)
            GroupBoxAlternativeMaterial.Location = New Point(725, 16)

            GroupBoxAlternativeMaterial.Width = 557
            GroupBoxAlternativeMaterial.Height = 481
            GroupBoxAlternativeMaterial.BringToFront()

            ' DataGridViewAltParts1.Location = New Point(16, 24)
            DataGridViewAltParts1.Location = New Point(9, 37)

            DataGridViewAltParts1.Width = 510
            DataGridViewAltParts1.Height = 195


            'DataGridViewAltParts2.Location = New Point(16, 247)
            DataGridViewAltParts2.Location = New Point(9, 247)

            DataGridViewAltParts2.Width = 510
            DataGridViewAltParts2.Height = 195

            'GroupBoxAlternativeMaterial.BringToFront()

        End If

        'checKing for custom parts

        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL2 As SqlCommand
        Dim drSQL2 As SqlDataReader
        Dim strSQL2 As String

        strSQL2 = "SELECT ItemNumber FROM [FSDBBR].[dbo].[TSS_Price_Item_Part_Select_Custom] where ItemNumber = '" & partnum & "' "


        cnSQL2.Open()
        cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
        drSQL2 = cmSQL2.ExecuteReader()

        If drSQL2.Read() Then

            MsgBox("This is a custom part, so alternative part is not available !", vbInformation)
            GroupBoxAlternativeMaterial.Visible = False
            Exit Sub

        End If

        cnSQL2.Close()


        'end of checking for custom parts

        'checking for kits

        strSQL2 = "SELECT ItemNumber FROM [FSDBBR].[dbo].[TSS_Price_Kit_Items] where ItemNumber = '" & partnum & "' "


        cnSQL2.Open()
        cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
        drSQL2 = cmSQL2.ExecuteReader()

        If drSQL2.Read() Then

            MsgBox("This is a kit, so alternative part is not available !", vbInformation)
            GroupBoxAlternativeMaterial.Visible = False
            Exit Sub

        End If

        cnSQL2.Close()


        'end of checking



        If Mid(partnum, 1, 2) = "OR" Or Mid(partnum, 1, 4) = "ORAR" Then


            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim strSQL As String

            'Dim PART As String

            'PART = Right(txtPartNumber.Text, 9)

            Dim s As String
            s = Mid(partnum, 1, 9)

            's = Left(s, length)
            s = Mid(s, 1, 9)
            's = Right(s, length)
            s = s & "%"

            Dim stockDC As DataSet = New DataSet


            strSQL = "select ItemNumber, ItemDescription from  [FSDBBR].[dbo].FS_Item WHERE ItemNumber like  '" & s & "'  and ItemNumber  not like '" & partnum & "'  and " & _
                     "ItemNumber in ( select ItemNumber from  [FSDBBR].[dbo].TSS_Price_Order_History)"



            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAC.TableMappings.Add("Table", "Enq")
            'get data
            stockDAC.Fill(stockDC)

            DataGridViewAltParts1.DataSource = stockDC.Tables(0)


            Dim itemnum As DataGridViewColumn = DataGridViewAltParts1.Columns(0)
            itemnum.Width = 180

            Dim itemdesc As DataGridViewColumn = DataGridViewAltParts1.Columns(1)
            itemdesc.Width = 290

            cnSQL.Close()


            If Mid(txtPartNumber.Text, 1, 4) = "ORAR" Then
                logic1()
            ElseIf Mid(txtPartNumber.Text, 1, 2) = "OR" And Mid(txtPartNumber.Text, 1, 4) <> "ORAR" Then
                logic2()

            End If

        Else
            'check for logic table data

            strSQL2 = "SELECT ItemNumber FROM [FSPrograms].[dbo].[TSS_Price_PartNumbers] where ItemNumber = '" & Mid(partnum, 1, 4) & "' "


            cnSQL2.Open()
            cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
            drSQL2 = cmSQL2.ExecuteReader()

            If drSQL2.Read() Then

                logic3()
            Else
                MsgBox("This is a custom part, so alternative part is not available !", vbInformation)
                GroupBoxAlternativeMaterial.Visible = False
                Exit Sub

            End If

            cnSQL2.Close()



        End If



    End Sub

    Private Sub logic1()

        'alternative size
        '  Dim iPosition As Integer
        'iPosition = GetPositionOfFirstNumericCharacter("ololo123")
        'Dim lenp As Integer = 9
        'Dim nump As Integer

        'iPosition = GetPositionOfFirstNumericCharacter(s)

        'nump = lenp - iPosition +1 
        'Dim partnum As String

        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be entered", vbInformation)
            Exit Sub
        End If

        Dim part As String
        part = partnum


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        'Dim PART As String

        'PART = Right(txtPartNumber.Text, 9)

        Dim s As Integer
        Dim e As Integer
        Dim p As Integer

        Dim spart As String
        Dim epart As String

        p = Mid(part, 5, 5)

        s = p - 1
        e = p + 1

        Dim partnum1 As String

        partnum1 = "ORAR"


        Dim SS As String
        Dim ES As String

        ' SS = String.Format("{00000}", s)
        'ES = String.Format("{00000}", e)

        SS = Format(s, "00000")
        ES = Format(e, "00000")


        spart = partnum1 & SS & "%"
        epart = partnum1 & ES & "%"


        'String.Format("{0:00000}", 15)


        's = Left(s, length)
        's = Mid(s, 1, 9)
        's = Right(s, length)
        's = s + "%"

        Dim stockDC As DataSet = New DataSet

        'strSQL = "select ItemNumber, ItemDescription from  [FSDBBR].[dbo].FS_Item WHERE ItemNumber like  '" & Mid(txtPartNumber.Text, 5, 5) & "' + '%'"

        'select ItemNumber,CONVERT(Decimal(10,0),RIGHT(LEFT(ItemNumber,9),2))  from  [FSDBBR].[dbo].FS_Item WHERE ItemNumber like  'ORAR%'  and ItemNumber not like 'ORAR00015%' AND (CONVERT(Decimal(10,0),RIGHT(LEFT(ItemNumber,9),5)) > 12 AND CONVERT(Decimal(10,0),RIGHT(LEFT(ItemNumber,9),5)) < 15)

        strSQL = "select ItemNumber,ItemDescription  from  [FSDBBR].[dbo].FS_Item WHERE ItemNumber like  'ORAR%'  and left(ItemNumber,9) not like '" & Mid(partnum, 1, 9) & "' And  (ItemNumber like '" & spart & "' or ItemNumber like '" & epart & "') AND ItemNumber in ( select ItemNumber from  [FSDBBR].[dbo].TSS_Price_Order_History)"



        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewAltParts2.DataSource = stockDC.Tables(0)

        Dim itemnum As DataGridViewColumn = DataGridViewAltParts2.Columns(0)
        itemnum.Width = 180

        Dim itemdesc As DataGridViewColumn = DataGridViewAltParts2.Columns(1)
        itemdesc.Width = 290

        cnSQL.Close()

    End Sub
    Private Sub logic2()

        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be enteted", vbInformation)
            Exit Sub
        End If

        'alternative size
        Dim iPosition As Integer
        'iPosition = GetPositionOfFirstNumericCharacter("ololo123")
        'Dim lenp As Integer = 9
        'Dim nump As Integer

        Dim part As String
        part = partnum 'txtPartNumber.Text

        iPosition = GetPositionOfFirstNumericCharacter(part)

        'nump = lenp - iPosition + 1

        If iPosition = 3 Then

            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim strSQL As String

            'Dim PART As String

            'PART = Right(txtPartNumber.Text, 9)

            Dim s As Integer
            Dim e As Integer
            Dim p As Integer

            'Dim spart As String
            'Dim epart As String

            p = Mid(part, 6, 4)
            p = p / 100

            Dim CS As Integer
            CS = Mid(part, 3, 3)





            s = (p - 2) ' * 100
            e = (p + 2) '* 100

            Dim partnum1 As String

            partnum1 = Mid(part, 1, 5)




            Dim stockDC As DataSet = New DataSet


            strSQL = "select ItemNumber,ItemDescription  from  [FSDBBR].[dbo].TSS_FS_ITEM_Part_Select WHERE ItemNumber like  'OR%'  and num = 1 and num1 = 1 and left(ItemNumber,9) not like '" & Mid(partnum, 1, 9) & "' And CS = " & CS & " and   ID >= " & s & " and ID <= " & e & "  AND ItemNumber in ( select ItemNumber from  [FSDBBR].[dbo].TSS_Price_Order_History)"

            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAC.TableMappings.Add("Table", "Enq")
            'get data                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
            stockDAC.Fill(stockDC)

            DataGridViewAltParts2.DataSource = stockDC.Tables(0)


            Dim itemnum As DataGridViewColumn = DataGridViewAltParts2.Columns(0)
            itemnum.Width = 150

            Dim itemdesc As DataGridViewColumn = DataGridViewAltParts2.Columns(1)
            itemdesc.Width = 250

            cnSQL.Close()
        End If

    End Sub

    Private Sub logic3()


        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be entered", vbInformation)
            Exit Sub
        End If


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String
        Dim PART As String

        PART = Mid(partnum, 1, 4)

        Dim s As String
        s = Mid(partnum, 1, 9)
        s = Mid(s, 1, 9)
        s = s & "%"

        Dim stockDC As DataSet = New DataSet


        strSQL = "select ItemNumber, ItemDescription from  [FSDBBR].[dbo].FS_Item WHERE ItemNumber like  '" & s & "'  and ItemNumber  not like '" & partnum & "'  AND ItemNumber in ( select ItemNumber from  [FSDBBR].[dbo].TSS_Price_Order_History)"

        ' strSQL = "Select ItemNumber, ItemDescription from [FSDBBR].[dbo].TSS_FS_ITEM_Part_Select WHERE ItemNumber = '" & PART & "' AND ItemNumber not like '" & Mid(partnum, 1, 9) & "' "

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
        stockDAC.Fill(stockDC)

        DataGridViewAltParts1.DataSource = stockDC.Tables(0)


        Dim itemnum As DataGridViewColumn = DataGridViewAltParts1.Columns(0)
        itemnum.Width = 150

        Dim itemdesc As DataGridViewColumn = DataGridViewAltParts1.Columns(1)
        itemdesc.Width = 250

        cnSQL.Close()

        'nearest sizes

        Dim k As String
        k = Mid(partnum, 1, 4)
        '  k = Mid(s, 1, 9)
        k = k & "%"

        Dim stockDC1 As DataSet = New DataSet
        strSQL = "Select ItemNumber, ItemDescription from [FSDBBR].[dbo].FS_Item WHERE ItemNumber  like '" & k & "' and ItemNumber  not like '" & partnum & "' AND ItemNumber in ( select ItemNumber from  [FSDBBR].[dbo].TSS_Price_Order_History) order by ItemNumber "

        Dim sqlCmd1 As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC1 As SqlDataAdapter = New SqlDataAdapter
        stockDAC1.SelectCommand = sqlCmd1
        cnSQL.Open()

        stockDAC1.TableMappings.Add("Table1", "Enq1")
        'get data                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
        stockDAC1.Fill(stockDC1)

        DataGridViewAltParts2.DataSource = stockDC1.Tables(0)

        Dim itemnum1 As DataGridViewColumn = DataGridViewAltParts2.Columns(0)
        itemnum1.Width = 150

        Dim itemdesc1 As DataGridViewColumn = DataGridViewAltParts2.Columns(1)
        itemdesc1.Width = 250

        cnSQL.Close()


    End Sub





    Public Function GetPositionOfFirstNumericCharacter(ByVal s As String) As Integer
        Dim i As Integer
        ' For i = 1 To Len(s)

        For i = 1 To Len(s)
            Dim currentCharacter As String
            currentCharacter = Mid(s, i, 1)
            If IsNumeric(currentCharacter) = True Then
                GetPositionOfFirstNumericCharacter = i
                Exit Function
            End If
        Next i
    End Function

    Private Sub DataGridViewAltParts1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewAltParts1.CellContentClick

    End Sub

    Private Sub DataGridViewSalesHistory_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewSalesHistory.CellClick

        GroupBoxDetailView.Visible = True
        DataGridViewSalesDetailView.Visible = True

        GroupBoxDetailView.Location = New Point(694, 244)

        GroupBoxDetailView.Width = 582
        GroupBoxDetailView.Height = 216
        GroupBoxDetailView.BringToFront()

        DataGridViewSalesDetailView.Location = New Point(6, 14)

        DataGridViewSalesDetailView.Width = 530
        DataGridViewSalesDetailView.Height = 159
        DataGridViewSalesDetailView.BringToFront()


        'important coding to findout the row clicked. 
        'Dim iRowIndex As Integer
        'For i As Integer = 0 To DataGridViewSalesHistory.SelectedCells.Count - 1

        '        iRowIndex = DataGridViewSalesHistory.SelectedCells.Item(i).RowIndex
        '       MsgBox("Row index " & Format(iRowIndex))
        '      Next


        Dim unitprice As Decimal
        Dim customerid As String
        Dim ttype As String

        customerid = DataGridViewSalesHistory.CurrentRow.Cells(2).Value.ToString
        unitprice = DataGridViewSalesHistory.CurrentRow.Cells(7).Value
        ttype = DataGridViewSalesHistory.CurrentRow.Cells(1).Value

        Lblsalesdetailview.Text = DataGridViewSalesHistory.CurrentRow.Cells(3).Value.ToString

        Dim periodyearS As Integer
        periodyearS = Val(txtSales.Text)
        ' periodyear = 2 * 365


        Dim d As Date = Date.Today

        ' d = d.AddDays(-periodyear)

        d = d.AddMonths(-periodyearS)

        Dim STRSQL As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim stockDC As DataSet = New DataSet

        If usertype = "S" Then
            If Val(txtnotional.Text) > 0 Then

                STRSQL = "SELECT TransactionDate as Trans_Date, ShipQty, UnitPrice, " & txtnotional.Text & " as 'Notional Price' FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' and  UnitPrice = " & unitprice & " and CustomerID = '" & customerid & "' and type = '" & ttype & "' order by TransactionDate desc"


            Else
                STRSQL = "SELECT TransactionDate as Trans_Date, ShipQty, UnitPrice, Cost FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' and  UnitPrice = " & unitprice & " and CustomerID = '" & customerid & "' and type = '" & ttype & "' order by TransactionDate desc"
            End If


        ElseIf usertype = "Q" Then

            If Val(txtnotional.Text) > 0 Then

                STRSQL = "SELECT TransactionDate as Trans_Date, ShipQty, UnitPrice FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' and  UnitPrice = " & unitprice & " and CustomerID = '" & customerid & "' and type = '" & ttype & "' order by TransactionDate desc"


            Else
                STRSQL = "SELECT TransactionDate as Trans_Date, ShipQty, UnitPrice FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' and  UnitPrice = " & unitprice & " and CustomerID = '" & customerid & "' and type = '" & ttype & "' order by TransactionDate desc"
            End If



        End If








        Dim sqlCmd As SqlCommand = New SqlCommand(STRSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewSalesDetailView.DataSource = stockDC.Tables(0)

        'sales history column setting

        'Dim InvoiceDate As DataGridViewColumn = DataGridViewSalesHistory.Columns(0) ' Trsndate
        'InvoiceDate.Width = 80

        'Dim CustomerID As DataGridViewColumn = DataGridViewSalesHistory.Columns(1) 'custid
        'CustomerID.Width = 70

        'Dim CustomerName As DataGridViewColumn = DataGridViewSalesHistory.Columns(2) 'custname
        'CustomerName.Width = 360

        'Dim CustomerClass1 As DataGridViewColumn = DataGridViewSalesHistory.Columns(3) 'custclass
        'CustomerClass1.Width = 80

        Dim ShipQty As DataGridViewColumn = DataGridViewSalesDetailView.Columns(1) 'shipqty
        'ShipQty.Width = 130
        DataGridViewSalesDetailView.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewSalesDetailView.Columns.Item(1).ValueType = GetType(Double)
        DataGridViewSalesDetailView.Columns.Item(1).DefaultCellStyle.Format = "n2"

        Dim unitpirce As DataGridViewColumn = DataGridViewSalesDetailView.Columns(2) 'unitprice
        'ShipQty.Width = 130
        DataGridViewSalesDetailView.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewSalesDetailView.Columns.Item(2).ValueType = GetType(Double)
        DataGridViewSalesDetailView.Columns.Item(2).DefaultCellStyle.Format = "n2"

        If usertype = "S" Then

            Dim cost As DataGridViewColumn = DataGridViewSalesDetailView.Columns(3) 'cost
            'ShipQty.Width = 130
            DataGridViewSalesDetailView.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridViewSalesDetailView.Columns.Item(3).ValueType = GetType(Double)
            DataGridViewSalesDetailView.Columns.Item(3).DefaultCellStyle.Format = "n2"
        End If




    End Sub

    Private Sub DataGridViewSalesHistory_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewSalesHistory.CellContentClick

    End Sub

    Private Sub lblAltClose_Click(sender As Object, e As EventArgs) Handles lblAltClose.Click
        GroupBoxAlternativeMaterial.Visible = False

    End Sub

    Private Sub txtPartNumber_LostFocus(sender As Object, e As EventArgs) Handles txtPartNumber.LostFocus
        'If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
        '    partnum = txtPartNumber.Text
        'ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
        '    partnum = txtPartNumberAlt.Text

        'Else
        '    MsgBox("Part number to be enteted", vbInformation)
        '    Exit Sub
        'End If





        'Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim strSQL As String
        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader

        'strSQL = "select ItemNumber, ItemDescription from  [FSDBBR].[dbo].FS_Item WHERE ItemNumber like  '" & (txtPartNumber.Text) & "'"



        'cnSQL.Open()
        'cmSQL = New SqlCommand(strSQL, cnSQL)
        'drSQL = cmSQL.ExecuteReader()



        'If drSQL.Read() Then

        '    If IsDBNull(drSQL.Item(0)) Then
        '        lblPartDescription.Text = ""
        '    Else

        '        lblPartDescription.Text = drSQL.Item(0) & " " & drSQL.Item(1)

        '    End If
        'End If

        'cnSQL.Close()

    End Sub

    Private Sub txtPartNumber_MarginChanged(sender As Object, e As EventArgs) Handles txtPartNumber.MarginChanged

    End Sub

    Private Sub txtPartNumber_ModifiedChanged(sender As Object, e As EventArgs) Handles txtPartNumber.ModifiedChanged

    End Sub

    Private Sub txtPartNumber_TextChanged(sender As Object, e As EventArgs) Handles txtPartNumber.TextChanged

        ' txtPartNumber.Text = UpperCase(txtPartNumber.Text)

        '  DataGridViewProjectMasterEdit.Columns.Clear()

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub lblPartDescription_Click(sender As Object, e As EventArgs) Handles lblPartDescription.Click

    End Sub

    Private Sub txtCustID_DockChanged(sender As Object, e As EventArgs) Handles txtCustID.DockChanged

    End Sub

    Private Sub TextBox1_DoubleClick(sender As Object, e As EventArgs) Handles txtCustID.DoubleClick

        '    If RadioButtonId.Checked = False And RadioButtonName.Checked = False Then
        '        MsgBox(" Select ID or Name ", vbInformation)
        '        Exit Sub
        '    End If

        '    DataGridViewCustomer.Visible = True

        '    fillcustomerlist()
    End Sub
    Sub fillcustomerlist()

        DataGridViewCustomer.Visible = True

        DataGridViewCustomer.Show()

        DataGridViewCustomer.Location = New Point(9, 93)
        DataGridViewCustomer.Width = 655
        DataGridViewCustomer.Height = 228


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim strSql As String
        Dim stockDC As DataSet = New DataSet


        ComboBoxCust.Text = ComboBoxCust.Text.ToUpper()

        txtCustID.Text = ComboBoxCust.Text & "%"

        If RadioButtonName.Checked = True Then
            '   strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR, [Cut_Off_Margin_Perc], [Target_Margin_Perc] FROM FSDBBR.dbo.[TSS_Customer_Mar] " & _
            '       "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRELLEBORG%') AND (CustomerID NOT LIKE '0000%')AND CustomerName like '" & txtCustID.Text & "' " & _
            '         "ORDER BY CustomerID"


            strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.[TSS_Customer_Mar] " & _
                            "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRELLEBORG%') AND (CustomerID NOT LIKE '0000%')AND CustomerName like '" & txtCustID.Text & "' " & _
                               "ORDER BY CustomerID"


        ElseIf RadioButtonId.Checked = True Then

            'strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR, [Cut_Off_Margin_Perc] , [Target_Margin_Perc] FROM FSDBBR.dbo.[TSS_Customer_Mar] " & _
            '         "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRELLEBORG%') AND (CustomerID NOT LIKE '0000%')AND CustomerID like '" & txtCustID.Text & "' " & _
            '            "ORDER BY CustomerID"



            strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.[TSS_Customer_Mar] " & _
                  "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRELLEBORG%') AND (CustomerID NOT LIKE '0000%')AND CustomerID like '" & txtCustID.Text & "' " & _
                     "ORDER BY CustomerID"


        End If





        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewCustomer.DataSource = stockDC.Tables(0)
        'cnSQL.Close()

        ' column setting

        Dim CustomerId As DataGridViewColumn = DataGridViewCustomer.Columns(0)
        CustomerId.Width = 100

        Dim CustomerName As DataGridViewColumn = DataGridViewCustomer.Columns(1)
        CustomerName.Width = 250

        Dim city As DataGridViewColumn = DataGridViewCustomer.Columns(2)
        city.Width = 120

        Dim CSR As DataGridViewColumn = DataGridViewCustomer.Columns(3)
        CSR.Width = 60
        'end of col setting
        cnSQL.Close()


    End Sub

    Private Sub txtCustID_Enter(sender As Object, e As EventArgs) Handles txtCustID.Enter

        ' custdetails()




    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtCustID.TextChanged

        txtcutoffp.Text = ""
        txtcutoffprice.Text = ""
        txttargetp.Text = ""
        txttargetprice.Text = ""


    End Sub

    Private Sub GroupBox5_Enter(sender As Object, e As EventArgs) Handles GroupBox5.Enter

    End Sub

    Private Sub DataGridCustomer_CurrentCellChanged(sender As Object, e As EventArgs)
        'Dim a As Integer
        'Dim custid As String



        'a = DataGridViewCustomer.CurrentCell.ColumnNumber()

        'If a = 0 Then
        '    txtCustID.Text = DataGridCustomer.Item(DataGridCustomer.CurrentCell)

        '    txtCustom.Text = DataGridCustomer.Item(DataGridCustomer.CurrentCell.RowNumber, 1)

        '    txtCity.Text = DataGridCustomer.Item(DataGridCustomer.CurrentCell.RowNumber, 2)

        '    txtCustID.Enabled = False


        '    'txtCustomer.Text = DataGridCustomer.Item(


        'Else
        '    MsgBox("Click on CustomerID to select the customer", vbInformation)
        '    Exit Sub
        'End If

        'DataGridViewCustomer.Hide()

        ''2.datagrid1.item(0,0)<-----it gets the first column/row data of your datagrid
        ''3. 4.'if you want selected
        ''5.datagrid1.item(datagrid1.currentcell.rownumber,0)<---it gets the selected row and the first column 


    End Sub

    Private Sub DataGridCustomer_Navigate(sender As Object, ne As NavigateEventArgs)

    End Sub

    Private Sub GroupBoxPriceSuggestion_Enter(sender As Object, e As EventArgs) Handles GroupBoxPriceSuggestion.Enter

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Buttonsugprice.Click

        Dim periodyearS As Integer
        periodyearS = Val(txtSales.Text)
        ' periodyear = 2 * 365


        If Val(txtmin.Text) > 0 Then

        Else
            MsgBox("Sales data need to be viewed before clicking on price calculation", vbInformation)
            Exit Sub

        End If


        Dim d As Date = Date.Today

        ' d = d.AddDays(-periodyear)

        d = d.AddMonths(-periodyearS)


        If GroupBoxPriceSuggestion.Visible = False Then

            GroupBoxPriceSuggestion.Visible = True


            If GroupBoxAlternativeMaterial.Visible = True Then
                GroupBoxAlternativeMaterial.Visible = False
            End If

            ' GroupBoxEnqDetails1.Location = New Point(6, 7)
            GroupBoxPriceSuggestion.Location = New Point(202, 21)

            GroupBoxPriceSuggestion.Width = 1094
            GroupBoxPriceSuggestion.Height = 416


            RadioButtonId.Checked = True

            priceclear()

            'customer loading

            Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim strSql As String
            Dim source As DataSet = New DataSet
            Dim cmSQL As SqlCommand

            strSql = "SELECT  CustomerID, (CustomerID + ' - ' + CustomerName)AS CUST  FROM  [FSDBBR].[dbo].[TSS_Price_Order_History] where ItemNumber = '" & partnum & "' and  TransactionDate >= '" & d & "' GROUP BY CustomerID,(CustomerID + ' - ' + CustomerName) "

            cmSQL = New SqlCommand(strSql, sqlCon)
            Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
            Dim ESource As SqlDataAdapter = New SqlDataAdapter
            ESource.SelectCommand = sqlCmd
            ESource.Fill(source, "eSource")
            With ComboBoxCust
                .DataSource = source.Tables("eSource")
                .DisplayMember = "CUST"
                .ValueMember = "CustomerID"
                .SelectedIndex = 0
            End With


            'end of customer loading


        End If

    End Sub


    '    If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
    '        partnum = txtPartNumber.Text
    '    ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
    '        partnum = txtPartNumberAlt.Text

    '    Else
    '        MsgBox("Part number to be enteted", vbInformation)
    '        Exit Sub
    '    End If


    'Dim periodyearS As Integer
    '    periodyearS = Val(txtSales.Text)
    '' periodyear = 2 * 365


    'Dim d As Date = Date.Today

    '' d = d.AddDays(-periodyear)

    '    d = d.AddMonths(-periodyearS)

    '    If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
    '        partnum = txtPartNumber.Text
    '    ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
    '        partnum = txtPartNumberAlt.Text

    '    Else
    '        MsgBox("Part number to be enteted", vbInformation)
    '        Exit Sub
    '    End If
    'If GroupBoxPriceSuggestion.Visible = False Then

    '    GroupBoxPriceSuggestion.Visible = True


    '    If GroupBoxAlternativeMaterial.Visible = True Then
    '        GroupBoxAlternativeMaterial.Visible = False
    '    End If

    '    ' GroupBoxEnqDetails1.Location = New Point(6, 7)
    '    GroupBoxPriceSuggestion.Location = New Point(202, 21)

    '    GroupBoxPriceSuggestion.Width = 1094
    '    GroupBoxPriceSuggestion.Height = 416


    '    RadioButtonId.Checked = True

    'End If



    ''particular customer, max percentage
    'strSQL = "SELECT [SALES_AMT],[COST_AMT],[SALES_QTY],MarPer, INV_DATE FROM [FSDBBR].[dbo].[TSS_Price_Sales_Cogs] where MarPer IN " & _
    '      " (SELECT max(MarPer) from [FSDBBR].[dbo].[TSS_Price_Sales_Cogs] where ORDER_TYPE = 'INVOICE' AND  ADDR_NBR = '" & txtCustID.Text & "' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber LIKE '" & partnum & "' " & _
    '       " and ([INV_DATE] >= '" & d & "' ) ) and ORDER_TYPE = 'INVOICE' AND  ADDR_NBR = '" & txtCustID.Text & "' AND  INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber LIKE '" & partnum & "' and ([INV_DATE] >=  '" & d & "') ORDER BY MarPer desc "


    'cnSQL.Open()
    'cmSQL = New SqlCommand(strSQL, cnSQL)
    'drSQL = cmSQL.ExecuteReader()



    'If drSQL.Read() Then

    '    If IsDBNull(drSQL.Item(0)) Then
    '        txtcustmaxperc.Text = ""
    '        txtcustmaxprice.Text = ""
    '        txtcustmaxcost.Text = ""
    '        txtcustmaxqty.Text = ""
    '        txtcustmaxdate.Text = ""

    '    Else
    '        txtcustmaxperc.Text = drSQL.Item(3)
    '        txtcustmaxprice.Text = drSQL.Item(0)
    '        txtcustmaxcost.Text = drSQL.Item(1)
    '        txtcustmaxqty.Text = drSQL.Item(2)
    '        txtcustmaxdate.Text = drSQL.Item(4)

    '    End If
    'End If

    'cnSQL.Close()

    '' PARTICULAR CUSTOMER 'MIN PERC
    ''particular customer, max percentage
    'strSQL = "SELECT [SALES_AMT],[COST_AMT],[SALES_QTY],MarPer, INV_DATE FROM [FSDBBR].[dbo].[TSS_Price_Sales_Cogs] where MarPer IN " & _
    '      " (SELECT MIN(MarPer) from [FSDBBR].[dbo].[TSS_Price_Sales_Cogs] where ORDER_TYPE = 'INVOICE' AND  ADDR_NBR = '" & txtCustID.Text & "' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber LIKE '" & partnum & "' " & _
    '       " and ([INV_DATE] >= '" & d & "' ) ) and ORDER_TYPE = 'INVOICE' AND  ADDR_NBR = '" & txtCustID.Text & "' AND  INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber LIKE '" & partnum & "' and ([INV_DATE] >=  '" & d & "') ORDER BY MarPer desc "


    'cnSQL.Open()
    'cmSQL = New SqlCommand(strSQL, cnSQL)
    'drSQL = cmSQL.ExecuteReader()



    'If drSQL.Read() Then

    '    If IsDBNull(drSQL.Item(0)) Then
    '        txtcustminperc.Text = ""
    '        txtcustminprice.Text = ""
    '        txtcustmincost.Text = ""
    '        txtcustminqty.Text = ""
    '        txtcustmindate.Text = ""

    '    Else
    '        txtcustminperc.Text = drSQL.Item(3)
    '        txtcustminprice.Text = drSQL.Item(0)
    '        txtcustmincost.Text = drSQL.Item(1)
    '        txtcustminqty.Text = drSQL.Item(2)
    '        txtcustmindate.Text = drSQL.Item(4)

    '    End If
    'End If

    'cnSQL.Close()






    '        Sql = "SELECT [INV_DATE], ADDR_NBR, [SALES_AMT],[COST_AMT],[SALES_QTY],[ItemNumber],[MarPer] FROM [FSDBBR].[dbo].[TSS_Price_Sales_Cogs] where  ORDER_TYPE = 'INVOICE' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber = '" & partnum & "'  and [INV_DATE] >= '" & d & "') and ItemNumber =  '" & partnum & "' and [INV_DATE] >= '" & d & "'"


    '  End Sub

    Private Sub DataGridViewCustomer_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewCustomer.CellContentClick

    End Sub

    Private Sub DataGridViewCustomer_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewCustomer.RowHeaderMouseClick


        'If GroupBoxSugSummary.Visible = False Then
        '    GroupBoxSugSummary.Visible = True
        'End If



        '  txtCustID.Text = DataGridViewCustomer.CurrentRow.Cells(0).Value.ToString

        ComboBoxCust.Text = DataGridViewCustomer.CurrentRow.Cells(0).Value.ToString


        txtCustom.Text = DataGridViewCustomer.CurrentRow.Cells(1).Value.ToString
        txtCity.Text = DataGridViewCustomer.CurrentRow.Cells(2).Value.ToString
        'txtcutoffp.Text = DataGridViewCustomer.CurrentRow.Cells(4).Value.ToString
        'txttargetp.Text = DataGridViewCustomer.CurrentRow.Cells(5).Value.ToString



        'If Val(txtnotional.Text) > 0 Then
        '    If Val(txtcutoffp.Text) > 0 Then
        '        txtcutoffprice.Text = Format((Val(txtnotional.Text) / ((100 - Val(txtcutoffp.Text)) / 100)), "0.00")
        '        '  txtcutoffprice.Text = Format(txtcutoffprice.Text, "0.00")

        '    End If
        '    If Val(txttargetp.Text) > 0 Then
        '        txttargetprice.Text = Format((Val(txtnotional.Text) / ((100 - Val(txttargetp.Text)) / 100)), "0.00")
        '        '  txttargetprice.Text = Format(txttargetprice.Text, "0.00")
        '    End If

        'Else


        '    If Val(txtcutoffp.Text) > 0 Then
        '        txtcutoffprice.Text = Format((Val(txtitmc.Text) / ((100 - Val(txtcutoffp.Text)) / 100)), "0.00")
        '        '  txtcutoffprice.Text = Format(txtcutoffprice.Text, "0.00")
        '    End If
        '    If Val(txttargetp.Text) > 0 Then
        '        txttargetprice.Text = Format((Val(txtitmc.Text) / ((100 - Val(txttargetp.Text)) / 100)), "0.00")
        '        ' txttargetprice.Text = Format(txttargetprice.Text, "0.00")

        '    End If


        'End If

        DataGridViewCustomer.Visible = False



    End Sub

    Private Sub txttargetp_TextChanged(sender As Object, e As EventArgs) Handles txttargetp.TextChanged

    End Sub

    Private Sub custdetails()

        Dim CUST As String

        CUST = Mid(ComboBoxCust.Text, 1, 6)


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String


        'strSQL1 = "SELECT CustomerID, CustomerName, CustomerCity, CSR, [Cut_Off_Margin_Perc] , [Target_Margin_Perc] FROM FSDBBR.dbo.[TSS_Customer_Mar] " & _
        '        "WHERE  CustomerID = '" & txtCustID.Text & "' "


        '   strSQL1 = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.[TSS_Customer_Mar] " & _
        '       "WHERE  CustomerID = '" & txtCustID.Text & "' "

        strSQL1 = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.[TSS_Customer_Mar] " & _
            "WHERE  CustomerID = '" & CUST & "' "


        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then



            If IsDBNull(drSQL1.Item(0)) Then
                MsgBox("This Customer ID is not existing,  click  on search to get the list", vbInformation)
                Exit Sub
            Else
                'txtmax.Text = Format(drSQL1.Item(0), "0.00")
                txtCustID.Text = drSQL1.Item(0)
                txtCustom.Text = drSQL1.Item(1)
                txtCity.Text = drSQL1.Item(2)

                '  If IsDBNull(drSQL1.Item(4)) Then

                'txtcutoffp.Text = 0
                'Else

                'txtcutoffp.Text = drSQL1.Item(4)
                'End If

                '    If IsDBNull(drSQL1.Item(5)) Then
                'txttargetp.Text = 0
                'Else
                'txttargetp.Text = drSQL1.Item(5)
                'End If





                'If Val(txtnotional.Text) > 0 Then
                '    If Val(txtcutoffp.Text) > 0 Then
                '        txtcutoffprice.Text = Format((Val(txtnotional.Text) / ((100 - Val(txtcutoffp.Text)) / 100)), "0.00")
                '        ' txtcutoffprice.Text = Format(txtcutoffprice.Text, "0.00")
                '    End If
                '    If Val(txttargetp.Text) > 0 Then
                '        txttargetprice.Text = Format((Val(txtnotional.Text) / ((100 - Val(txttargetp.Text)) / 100)), "0.00")
                '        'txttargetprice.Text = Format(txttargetprice.Text, "0.00")
                '    End If

                'Else


                '    If Val(txtcutoffp.Text) > 0 Then
                '        txtcutoffprice.Text = Format((Val(txtitmc.Text) / ((100 - Val(txtcutoffp.Text)) / 100)), "0.00")
                '        'txtcutoffprice.Text = Format(txtcutoffprice.Text, "0.00")
                '    End If
                '    If Val(txttargetp.Text) > 0 Then
                '        txttargetprice.Text = Format((Val(txtitmc.Text) / ((100 - Val(txttargetp.Text)) / 100)), "0.00")
                '        'txttargetprice.Text = Format(txttargetprice.Text, "0.00")
                '    End If


                'End If



            End If

        Else
            MsgBox("This Customer ID is not existing,  click  on search to get the list", vbInformation)
            Exit Sub

        End If

        cnSQL1.Close()


    End Sub

    Private Sub txtCustom_TextChanged(sender As Object, e As EventArgs) Handles txtCustom.TextChanged

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        GroupBoxSugSummary.Visible = False

    End Sub

    Private Sub lblCustClose_Click(sender As Object, e As EventArgs) Handles lblCustClose.Click
        GroupBoxPriceSuggestion.Visible = False

    End Sub

    Private Sub DataGridViewAltParts2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewAltParts2.CellContentClick
        Dim msgb As String

        msgb = MsgBox(" Do you want to know the pricing  of this item ?", vbYesNo)



        ' msgb = MsgBox("Are you sure of deleting this line ?", vbYesNo)

        If msgb = vbNo Then
            Exit Sub
        Else
            'CheckBoxAltPart.Checked = True

            'CheckBoxMainPart.Checked = False

            txtPartNumberAlt.Text = DataGridViewAltParts2.CurrentRow.Cells(0).Value.ToString
            lblAltPartDescription.Text = DataGridViewAltParts2.CurrentRow.Cells(1).Value.ToString

            GroupBoxAlternativeMaterial.Visible = False
            cleardatagridview()

            CheckBoxAltPart.Checked = True
            CheckBoxMainPart.Checked = False

        End If




    End Sub

    Private Sub txtcutoffprice_TextChanged(sender As Object, e As EventArgs) Handles txtcutoffprice.TextChanged

    End Sub

    Private Sub ButtonOK_Click(sender As Object, e As EventArgs) Handles ButtonOK.Click


        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be enteted", vbInformation)
            Exit Sub
        End If

        priceclear()

        Dim periodyearS As Integer
        periodyearS = Val(txtSales.Text)
        ' periodyear = 2 * 365


        Dim d As Date = Date.Today

        ' d = d.AddDays(-periodyear)

        d = d.AddMonths(-periodyearS)



        If RadioButtonId.Checked = False And RadioButtonName.Checked = False Then
            MsgBox("Search by option need to be selected", vbInformation)
            Exit Sub
        End If


        custdetails()

        'all customer
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader

        ' strSQL = "SELECT  [SALES_AMT],[COST_AMT],[SALES_QTY],max([MarPer]) as MarPerc FROM [FSDBBR].[dbo].[TSS_Price_Sales_Cogs] where  ORDER_TYPE = 'INVOICE' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber = '" & partnum & "' and ([INV_DATE] >= '" & d & "'  and [INV_DATE] >= '" & d & "'"

        strSQL = "SELECT [SALES_AMT],[COST_AMT],[SALES_QTY],MarPer, INV_DATE FROM [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where MarPer IN " & _
                 " (SELECT MAX(MarPer) from [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where ORDER_TYPE = 'INVOICE' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ADDR_NBR <> '" & txtCustID.Text & "' and ItemNumber LIKE '" & partnum & "' " & _
                 " and ([INV_DATE] >= '" & d & "' )) and ORDER_TYPE = 'INVOICE' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and  ADDR_NBR <> '" & txtCustID.Text & "' AND   ItemNumber LIKE '" & partnum & "' and ([INV_DATE] >=  '" & d & "') ORDER BY MarPer desc "


        cnSQL.Open()
        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()



        If drSQL.Read() Then

            If IsDBNull(drSQL.Item(0)) Then
                txtallcustmaxp.Text = ""
                txtallcustmaxqty.Text = ""
                txtallcustmaxprice.Text = ""
                txtallcustmaxcost.Text = ""
                txtallcustmaxInvDate.Text = ""
            Else
                txtallcustmaxp.Text = Format(drSQL.Item(3), "0.0")
                txtallcustmaxqty.Text = drSQL.Item(2)
                txtallcustmaxprice.Text = Format(drSQL.Item(0), "0.00")
                txtallcustmaxprice.Text = Format(Val(txtallcustmaxprice.Text) / Val(txtallcustmaxqty.Text), "0.00")

                txtallcustmaxcost.Text = drSQL.Item(1)
                txtallcustmaxcost.Text = Format(txtallcustmaxcost.Text / Val(txtallcustmaxqty.Text), "0.00")
                txtallcustmaxInvDate.Text = drSQL.Item(4)

            End If
        End If

        cnSQL.Close()

        'minimum all customers

        strSQL = "SELECT [SALES_AMT],[COST_AMT],[SALES_QTY],MarPer, INV_DATE FROM [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where MarPer IN " & _
               " (SELECT min(MarPer) from [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where ORDER_TYPE = 'INVOICE' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ADDR_NBR <> '" & txtCustID.Text & "' and  ItemNumber LIKE '" & partnum & "' " & _
               " and ([INV_DATE] >= '" & d & "' )) and ORDER_TYPE = 'INVOICE' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and  ADDR_NBR <> '" & txtCustID.Text & "' and ItemNumber LIKE '" & partnum & "' and ([INV_DATE] >=  '" & d & "') ORDER BY MarPer desc "


        cnSQL.Open()
        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()



        If drSQL.Read() Then

            If IsDBNull(drSQL.Item(0)) Then
                txtallcustminp.Text = ""
                txtallcustminqty.Text = ""
                txtallcustminprice.Text = ""
                txtallcustmincost.Text = ""
                txtallcustmindate.Text = ""
            Else
                txtallcustminp.Text = Format(drSQL.Item(3), "0.0")
                txtallcustminqty.Text = drSQL.Item(2)

                txtallcustminprice.Text = drSQL.Item(0)
                txtallcustminprice.Text = Format(Val(txtallcustminprice.Text) / Val(txtallcustminqty.Text), "0.00")

                txtallcustmincost.Text = drSQL.Item(1)
                txtallcustmincost.Text = Format(Val(txtallcustmincost.Text) / Val(txtallcustminqty.Text), "0.00")

                txtallcustmindate.Text = drSQL.Item(4)

            End If
        End If

        cnSQL.Close()



        'end of all customer


        '        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        '       Dim strSQL As String
        '      Dim cmSQL As SqlCommand
        '     Dim drSQL As SqlDataReader


        'particular customer, max percentage

        strSQL = "SELECT [SALES_AMT],[COST_AMT],[SALES_QTY],MarPer, INV_DATE FROM [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where MarPer IN " & _
              " (SELECT max(MarPer) from [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where ORDER_TYPE = 'INVOICE' AND  ADDR_NBR = '" & txtCustID.Text & "' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber LIKE '" & partnum & "' " & _
               " and ([INV_DATE] >= '" & d & "' ) ) and ORDER_TYPE = 'INVOICE' AND  ADDR_NBR = '" & txtCustID.Text & "' AND  INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber LIKE '" & partnum & "' and ([INV_DATE] >=  '" & d & "') ORDER BY MarPer desc "


        cnSQL.Open()
        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()



        If drSQL.Read() Then

            If IsDBNull(drSQL.Item(0)) Then
                txtcustmaxperc.Text = ""
                txtcustmaxprice.Text = ""
                txtcustmaxcost.Text = ""
                txtcustmaxqty.Text = ""
                txtcustmaxdate.Text = ""

            Else
                txtcustmaxqty.Text = drSQL.Item(2)
                txtcustmaxperc.Text = Format(drSQL.Item(3), "0.00")

                txtcustmaxprice.Text = drSQL.Item(0)
                txtcustmaxprice.Text = Format((Val(txtcustmaxprice.Text) / Val(txtcustmaxqty.Text)), "0.00")

                txtcustmaxcost.Text = drSQL.Item(1)
                txtcustmaxcost.Text = Format((Val(txtcustmaxcost.Text) / Val(txtcustmaxqty.Text)), "0.00")

                txtcustmaxdate.Text = drSQL.Item(4)

            End If
        End If

        cnSQL.Close()

        ' PARTICULAR CUSTOMER 'MIN PERC
        'particular customer, max percentage
        strSQL = "SELECT [SALES_AMT],[COST_AMT],[SALES_QTY],MarPer, INV_DATE FROM [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where MarPer IN " & _
              " (SELECT MIN(MarPer) from [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where ORDER_TYPE = 'INVOICE' AND  ADDR_NBR = '" & txtCustID.Text & "' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber LIKE '" & partnum & "' " & _
               " and ([INV_DATE] >= '" & d & "' ) ) and ORDER_TYPE = 'INVOICE' AND  ADDR_NBR = '" & txtCustID.Text & "' AND  INTER_COMP_FLG = 0 AND SALES_AMT > 0 and ItemNumber LIKE '" & partnum & "' and ([INV_DATE] >=  '" & d & "') ORDER BY MarPer desc "


        cnSQL.Open()
        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()



        If drSQL.Read() Then

            If IsDBNull(drSQL.Item(0)) Then
                txtcustminperc.Text = ""
                txtcustminprice.Text = ""
                txtcustmincost.Text = ""
                txtcustminqty.Text = ""
                txtcustmindate.Text = ""

            Else

                txtcustminqty.Text = drSQL.Item(2)
                txtcustminperc.Text = Format(drSQL.Item(3), "0.00")

                txtcustminprice.Text = drSQL.Item(0)
                txtcustminprice.Text = Format((Val(txtcustminprice.Text) / Val(txtcustminqty.Text)), "0.00")


                txtcustmincost.Text = drSQL.Item(1)
                txtcustmincost.Text = Format((Val(txtcustmincost.Text) / Val(txtcustminqty.Text)), "0.00")

                txtcustmindate.Text = drSQL.Item(4)

            End If
        End If

        cnSQL.Close()

        'TOTAL MARGIN


        strSQL = "SELECT SUM([SALES_AMT]) AS Sales,SUM([COST_AMT])as Cogs FROM [FSPrograms].[dbo].[TSS_Price_Sales_COGS] where  ORDER_TYPE = 'INVOICE' and  ADDR_NBR = '" & txtCustID.Text & "' AND INTER_COMP_FLG = 0 AND SALES_AMT > 0 AND " & _
               "  [INV_DATE] >= '" & d & "' and ORDER_NBR not like 'Z%'  "


        cnSQL.Open()
        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()



        If drSQL.Read() Then

            If IsDBNull(drSQL.Item(0)) Then
                txtCustPer.Text = ""

            Else

                txtCustPer.Text = Val(drSQL.Item(0)) - Val(drSQL.Item(1))


                txtCustPer.Text = Val(txtCustPer.Text) / Val(drSQL.Item(0))

                txtCustPer.Text = Format((Val(txtCustPer.Text) * 100), "0.00")

                ' txtCustPer.Text = Format(txtCustPer.Text, "0.00")

                '          Format(((Val(drSQL.Item(0)) - Val((drSQL.Item(1))) / Val(drSQL.Item(0))) * 100), "0.00")

            End If
        End If

        cnSQL.Close()

    End Sub

    Private Sub ButtonSearch_Click(sender As Object, e As EventArgs) Handles ButtonSearch.Click
        If RadioButtonId.Checked = False And RadioButtonName.Checked = False Then
            MsgBox("Select ID or Name ", vbInformation)
            Exit Sub
        End If

        DataGridViewCustomer.Visible = True

        fillcustomerlist()

    End Sub

    Private Sub Label36_Click(sender As Object, e As EventArgs)

    End Sub



    Private Sub GroupBox6_Enter(sender As Object, e As EventArgs) Handles GroupBox6.Enter

    End Sub

    Private Sub txtPur_TextChanged(sender As Object, e As EventArgs) Handles txtPur.TextChanged

    End Sub

    Private Sub DataGridViewSalesHistory_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewSalesHistory.RowHeaderMouseClick
        'codings need to be written

        'Dim row As String
        'Dim col As String


        'col = DataGridViewSalesDetailView.CurrentCellAddress.X() 'Column  
        'row = DataGridViewSalesDetailView.CurrentCellAddress.Y() 'Row


        'If col = 5 Then
        '    MsgBox("Coding need to be written here", vbInformation)
        '    Exit Sub
        'Else
        '    MsgBox("Click on instances  value to expand the sales history", vbInformation)
        '    Exit Sub
        'End If



        'str = DataGridViewSalesDetailView.Rows.(datagridviewsalesdetailview.SelectedRows(0).Index).cells






        'str = DataGridView1.Rows[DataGridView.SelectedRows[0].Index].Cells[X].Value.ToString();





    End Sub

    Private Sub DataGridViewSalesDetailView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridViewQuoteHistory_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewQuoteHistory.CellContentClick

    End Sub

    Private Sub Label57_Click(sender As Object, e As EventArgs) Handles LabelDetailViewClose.Click
        GroupBoxDetailView.Visible = False

    End Sub

    Private Sub DataGridViewAltParts1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewAltParts1.CellContentDoubleClick

    End Sub

    Private Sub DataGridViewAltParts1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewAltParts1.CellDoubleClick
        Dim msgb As String

        msgb = MsgBox(" Do you want to know the pricing  of this item ?", vbYesNo)



        ' msgb = MsgBox("Are you sure of deleting this line ?", vbYesNo)

        If msgb = vbNo Then
            Exit Sub
        Else
            ' CheckBoxAltPart.Checked = True
            'CheckBoxMainPart.Checked = False

            txtPartNumberAlt.Text = DataGridViewAltParts1.CurrentRow.Cells(0).Value.ToString
            lblAltPartDescription.Text = DataGridViewAltParts1.CurrentRow.Cells(1).Value.ToString

            CheckBoxAltPart.Checked = True
            CheckBoxMainPart.Checked = False

            GroupBoxAlternativeMaterial.Visible = False
            cleardatagridview()
        End If



    End Sub

    Private Sub txtPartNumberAlt_LostFocus(sender As Object, e As EventArgs) Handles txtPartNumberAlt.LostFocus
        'If CheckBoxMainPart.Checked = True Then
        'partnum = txtPartNumber.Text
        'ElseIf CheckBoxAltPart.Checked = True Then
        'partnum = txtPartNumberAlt.Text
        'End If
    End Sub

    Private Sub txtPartNumberAlt_TextChanged(sender As Object, e As EventArgs) Handles txtPartNumberAlt.TextChanged




    End Sub
    Private Sub cleardatagridview()
        DataGridViewSalesHistory.Columns.Clear()
        DataGridViewQuoteHistory.Columns.Clear()
        DataGridViewPurchaseHistory.Columns.Clear()
        DataGridViewAltParts1.Columns.Clear()
        DataGridViewAltParts2.Columns.Clear()
        DataGridViewCustomer.Columns.Clear()

        txtmin.Text = ""
        txtmininstances.Text = ""
        txtminqty.Text = ""

        txtmax.Text = ""
        txtmaxinstances.Text = ""
        txtmaxqty.Text = ""

        txtpurhighcost.Text = ""
        txtpurhighdate.Text = ""
        txtpurhighqty.Text = ""


        txtpurlowcost.Text = ""
        txtpurlowqty.Text = ""
        txtpurlowdate.Text = ""

        txtpurlatestCost.Text = ""
        txtpurlatestDate.Text = ""
        txtpurlatestqty.Text = ""

        txtnotional.Text = ""

    End Sub


    Private Sub BtnKit_Click(sender As Object, e As EventArgs) Handles BtnKit.Click


        If CheckBoxMainPart.Checked = True And Len(txtPartNumber.Text) > 5 Then
            partnum = txtPartNumber.Text
        ElseIf CheckBoxAltPart.Checked = True And Len(txtPartNumberAlt.Text) > 5 Then
            partnum = txtPartNumberAlt.Text

        Else
            MsgBox("Part number to be enteted", vbInformation)
            Exit Sub
        End If


        'grid

        GroupBoxkitdetails.Visible = True
        DataGridViewkit.Visible = True

        GroupBoxkitdetails.Location = New Point(614, 71)

        GroupBoxkitdetails.Width = 662
        GroupBoxkitdetails.Height = 162
        GroupBoxkitdetails.BringToFront()

        DataGridViewkit.Location = New Point(16, 21)

        DataGridViewkit.Width = 611
        DataGridViewkit.Height = 132
        DataGridViewkit.BringToFront()


        Dim STRSQL As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim stockDC As DataSet = New DataSet

        STRSQL = "Select [ItemNumber],[ItemDescription] AS Description,[ItemUM] as UM,[ItemStatus] as Status ,[RequiredQuantity] as ReqdQty  FROM [FSDBBR].[dbo].[TSS_BOM]  where ComponentItemNumber = '" & partnum & "'"


        Dim sqlCmd As SqlCommand = New SqlCommand(STRSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewkit.DataSource = stockDC.Tables(0)
        'end of grid
    End Sub

    Private Sub DataGridViewkit_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridViewkit_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        Dim msgb As String

        msgb = MsgBox(" Do you want to know the pricing  of this KIT ?", vbYesNo)



        ' msgb = MsgBox("Are you sure of deleting this line ?", vbYesNo)

        If msgb = vbNo Then
            Exit Sub
        Else

            txtPartNumberAlt.Text = DataGridViewkit.CurrentRow.Cells(0).Value.ToString
            GroupBoxkitdetails.Visible = False
            cleardatagridview()

            CheckBoxAltPart.Checked = True
            CheckBoxMainPart.Checked = False

        End If

    End Sub

    Private Sub Label58_Click(sender As Object, e As EventArgs) Handles Labelkitclose.Click
        GroupBoxkitdetails.Visible = False

    End Sub

    Private Sub DataGridViewkit_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewkit.CellContentClick

    End Sub

    Private Sub DataGridViewkit_CellDoubleClick1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewkit.CellDoubleClick
        Dim msgb As String

        msgb = MsgBox(" Do you want to know the pricing  of this item ?", vbYesNo)



        ' msgb = MsgBox("Are you sure of deleting this line ?", vbYesNo)

        If msgb = vbNo Then
            Exit Sub
        Else

            txtPartNumberAlt.Text = DataGridViewkit.CurrentRow.Cells(0).Value.ToString
            GroupBoxkitdetails.Visible = False
            cleardatagridview()

            CheckBoxAltPart.Checked = True
            CheckBoxMainPart.Checked = False


        End If

    End Sub

    Private Sub RadioButtonId_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonId.CheckedChanged

    End Sub

    Private Sub RadioButtonId_Click(sender As Object, e As EventArgs) Handles RadioButtonId.Click
        If RadioButtonId.Checked = True Then
            lblcustomer.Text = "Customer ID"
        Else
            lblcustomer.Text = "Customer Name"
        End If
    End Sub

    Private Sub RadioButtonName_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonName.CheckedChanged

    End Sub

    Private Sub RadioButtonName_Click(sender As Object, e As EventArgs) Handles RadioButtonName.Click
        If RadioButtonName.Checked = True Then
            lblcustomer.Text = "Customer Name"
        Else
            lblcustomer.Text = "Customer ID"
        End If
    End Sub

    Private Sub txtTargetSalPrice_DoubleClick(sender As Object, e As EventArgs) Handles txtTargetSalPrice.DoubleClick


        If Val(txtTargetSalPer.Text) > 0 Then

            If Val(txtnotional.Text) > 0 Then

                txtTargetSalPrice.Text = Format((Val(txtnotional.Text) / ((100 - Val(txtTargetSalPer.Text)) / 100)), "0.00")
            Else

                txtTargetSalPrice.Text = Format((Val(txtitmc.Text) / ((100 - Val(txtTargetSalPer.Text)) / 100)), "0.00")
            End If
        End If

        'txtcutoffprice.Text = Format((Val(txtnotional.Text) / ((100 - Val(txtcutoffp.Text)) / 100)), "0.00")


    End Sub

    Private Sub txtTargetSalPrice_TextChanged(sender As Object, e As EventArgs) Handles txtTargetSalPrice.TextChanged

    End Sub

    Private Sub txtcustmaxperc_TextChanged(sender As Object, e As EventArgs) Handles txtcustmaxperc.TextChanged

    End Sub

    Private Sub txtTargetSalPer_DoubleClick(sender As Object, e As EventArgs) Handles txtTargetSalPer.DoubleClick
        If Val(txtTargetSalPrice.Text) > 0 Then

            If Val(txtnotional.Text) > 0 Then
                txtTargetSalPer.Text = Format(((Val(txtTargetSalPrice.Text) - Val(txtnotional.Text)) / Val((txtTargetSalPrice.Text)) * 100), "0.00")
            Else
                txtTargetSalPer.Text = Format(((Val(txtTargetSalPrice.Text) - Val(txtitmc.Text)) / Val((txtTargetSalPrice.Text)) * 100), "0.00")

            End If

        End If






    End Sub

    Private Sub txtTargetSalPer_TextChanged(sender As Object, e As EventArgs) Handles txtTargetSalPer.TextChanged

    End Sub

    Private Sub DataGridViewAltParts2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridViewAltParts2.DoubleClick

    End Sub

    Private Sub priceclear()

        txtCustID.Text = ""
        txtCustom.Text = ""
        txtCity.Text = ""
        txtcustmaxperc.Text = ""
        txtcustmaxcost.Text = ""
        txtcustmaxprice.Text = ""
        txtcustmaxqty.Text = ""
        txtcustmaxdate.Text = ""

        txtcustminperc.Text = ""
        txtcustmincost.Text = ""
        txtcustminprice.Text = ""
        txtcustminqty.Text = ""
        txtcustmindate.Text = ""

        txtallcustmaxcost.Text = ""
        txtallcustmaxprice.Text = ""
        txtallcustmaxqty.Text = ""
        txtallcustmaxInvDate.Text = ""
        txtallcustmaxp.Text = ""

        txtallcustmincost.Text = ""
        txtallcustmindate.Text = ""
        txtallcustminp.Text = ""
        txtallcustminqty.Text = ""
        txtallcustminprice.Text = ""


        txtCustPer.Text = ""
        txtTargetSalPer.Text = ""
        txtTargetSalPrice.Text = ""

    End Sub

    Private Sub CheckBoxMainPart_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxMainPart.CheckedChanged
        '   CheckBoxAltPart.Checked = False
        '  CheckBoxMainPart.Checked = True



    End Sub

    Private Sub CheckBoxAltPart_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAltPart.CheckedChanged
        'CheckBoxMainPart.Checked = False
        'CheckBoxAltPart.Checked = True


    End Sub

    Private Sub txtCustPer_TextChanged(sender As Object, e As EventArgs) Handles txtCustPer.TextChanged

    End Sub

    Private Sub GroupBoxMailing_Enter(sender As Object, e As EventArgs) Handles GroupBoxMailing.Enter

    End Sub

    Private Sub BtnMailing_Click(sender As Object, e As EventArgs) Handles BtnMailing.Click
        GroupBoxMailing.Visible = True


        GroupBoxMailing.Location = New Point(9, 144)

        GroupBoxMailing.Height = 262
        GroupBoxMailing.Width = 1075



        MsgBox("This feature will be provided in ver2.0 release", vbInformation)
        Exit Sub


    End Sub

    Private Sub Label45_Click(sender As Object, e As EventArgs) Handles Label45.Click
        GroupBoxMailing.Visible = False
    End Sub

    Private Sub txtMailNotes_TextChanged(sender As Object, e As EventArgs) Handles txtMailNotes.TextChanged

    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox("This feature will be provided in ver2.0 release", vbInformation)
        Exit Sub

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("This feature will be provided in ver2.0 release", vbInformation)
        Exit Sub

    End Sub

    Private Sub DataGridViewPurchaseHistory_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewPurchaseHistory.CellContentClick

    End Sub

    Private Sub DataGridViewSalesDetailView_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewSalesDetailView.CellContentClick

    End Sub
    Private Sub itmcdetails()

        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL2 As SqlCommand
        Dim drSQL2 As SqlDataReader
        Dim strSQL2 As String


        strSQL2 = "SELECT TotalRolledCost FROM  [FSDBBR].[dbo].[TSS_ITEM_COST] where ItemNumber = '" & partnum & "'"
        cnSQL2.Open()
        cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
        drSQL2 = cmSQL2.ExecuteReader()

        If drSQL2.Read() Then

            If IsDBNull(drSQL2.Item(0)) Then
                txtitmc.Text = 0
            Else

                txtitmc.Text = Format(drSQL2.Item(0), "0.00")
            End If


        End If

        cnSQL2.Close()



        Dim cnSQL4 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL4 As SqlCommand
        Dim drSQL4 As SqlDataReader
        Dim strSQL4 As String

        strSQL4 = "SELECT  SUM([InventoryQuantity])  FROM [FSDBBR].[dbo].[TSS_Price_MarketingInventory] where ItemNumber = '" & partnum & "'"


        cnSQL4.Open()
        cmSQL4 = New SqlCommand(strSQL4, cnSQL4)
        drSQL4 = cmSQL4.ExecuteReader()

        If drSQL4.Read() Then

            If IsDBNull(drSQL4.Item(0)) Then
                txtcurstock.Text = 0

            Else

                txtcurstock.Text = Format(drSQL4.Item(0), "0.00")
                txtcurstock.Text = Val(txtcurstock.Text)
            End If


        End If

        cnSQL4.Close()

        'end of stock


        ' available to promise

        Dim cnSQL5 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL5 As SqlCommand
        Dim drSQL5 As SqlDataReader
        Dim strSQL5 As String

        strSQL5 = "SELECT  SUM([Pending Qty]) FROM [FSDBBR].[dbo].[TSS_PendingSalesOrders_ver5] where COLineStatus in (3,4) and (CustomerName not like 'TSS%' AND CustomerName not like 'TRE%')  AND ItemNumber = '" & partnum & "'"

        'Required for kit to be done.

        cnSQL5.Open()
        cmSQL5 = New SqlCommand(strSQL5, cnSQL5)
        drSQL5 = cmSQL5.ExecuteReader()

        If drSQL5.Read() Then

            If IsDBNull(drSQL5.Item(0)) Then
                txtAvblePromise.Text = 0

                txtAvblePromise.Text = Val(txtcurstock.Text) - Val(txtAvblePromise.Text)


            Else

                txtAvblePromise.Text = Format(drSQL5.Item(0), "0.00")

                txtAvblePromise.Text = Val(txtcurstock.Text) - Val(txtAvblePromise.Text)

            End If


        End If

        cnSQL5.Close()



    End Sub

End Class
