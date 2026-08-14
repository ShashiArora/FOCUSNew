Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
'Imports SoftBrands.FourthShift.Transaction
'Imports CrystalDecisions.CrystalReports.Engine
Imports System.Windows.Forms



Public Class RFQPriceViewALL

    Inherits System.Windows.Forms.Form
    Public selection As String
    Private ConnectionString As String
    Public stockDA As SqlDataAdapter = New SqlDataAdapter
    Public ZCLASS As String

    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"



    Private Sub RFQPriceView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RadioButtonunread.Checked = True

        'datagridRFQView.Enabled = True
        'Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim strSQL As String


        'Dim stockDC As DataSet = New DataSet

        'strSQL = "SELECT Enq_Detail_code as DetailCode,Enq_Reg_NO, Enq_Reg_date, CustomerID, CustomerName, CSR, Sl_no, PartNumber, PartDescription,  MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
        '        " Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Prod_TotalCost, Prod_CustCost, Prod_LeadTime,  " & _
        '       "  Remarks as SpecialInstructions FROM TSS_Enquiry_RFQView order by Enq_Reg_NO,Sl_no"


        '        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        '       Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        '     stockDAC.SelectCommand = sqlCmd
        '    cnSQL.Open()

        '   stockDAC.TableMappings.Add("Table", "Enq")
        '  stockDAC.Fill(stockDC)

        ' datagridRFQView.DataSource = stockDC.Tables(0)
        'cnSQL.Close()
        'datagridRFQView.Expand(-1)





    End Sub

    Private Sub datagridRFQView_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles datagridRFQView.CurrentCellChanged
        If selection = "" Then
            'If (RadioButtonread.Checked = True Or RadioButtonunread.Checked = True Or RadioButtonall.Checked = True) Then
            Dim b As Integer
            'Dim custid As String
            b = datagridRFQView.CurrentCell.ColumnNumber()

            If b = 0 Then

                txtenqdetailcode.Text = datagridRFQView.Item(datagridRFQView.CurrentCell)
                txtregno.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 1)

                txtspecial.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 21)

                ProtoTotal.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 15)
                ProtoCustShare.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 16)
                ProtoLeadTime.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 17)
                'ProtoQty.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 19)
                'ProtoLifeofTool.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 20)

                ProdTotalCost.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 18)
                ProdCustShare.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 19)
                ProdLeadTime.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 20)
                'ProdQty.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 24)
                'ProdLifeofTool.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 25)


            Else
                MsgBox("Click on Detailcode ", vbInformation)
                Exit Sub
            End If

            fillqtyprice()
            fillcertificatecharges()

        ElseIf selection = "part" Then
            txtpartbyApl.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 12)
            txtAplSpecial.Text = datagridRFQView.Item(datagridRFQView.CurrentCell.RowNumber, 15)


        End If


    End Sub

    Private Sub datagridRFQView_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles datagridRFQView.Navigate

    End Sub

    Private Sub DataGridQty_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub
    Private Sub fillqtyprice()
        DataGridQtyview.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet
        strSql = ""

        'If multiple = "YES" Then

        'not used this as on 12 april 2013


        strSql = "SELECT   FinalPrice as RatePerEach,Qty,Qty_Type as Type FROM ENQ_RFQ_Qty_PriceDetails " & _
                            "WHERE  Enq_Detail_code  = " & txtenqdetailcode.Text & " and Enq_Reg_NO = " & txtregno.Text & ""


        'ElseIf multiple = "NO" Then



        'If (txtitemstatus.Text = "H" Or txtitemstatus.Text = "C" Or txtitemstatus.Text = "F") And rfqmode = "" Then

        'strSql = "SELECT   FinalPrice, Qty,Qty_Type as Type FROM ENQ_RFQ_Qty_PriceDetails " & _
        '       "WHERE RFQ_Int_code = '" & txtRFQIntcode.Text & "' " & _
        '          "ORDER BY Qty"

        'ElseIf txtitemstatus.Text = "P" And rfqmode = "" Then

        'strSql = "SELECT  0.00 as Price,0.00 as Fact, 0.00 as FPrice,Qty,Qty_Type as Type,Enq_Qty_IntCode as IntCode FROM ENQ_Qty_Details " & _
        '                "WHERE  Enq_Detail_code = '" & txtenqdetailintcode.Text & "' " & _
        '                  "ORDER BY Qty"


        'ElseIf (txtitemstatus.Text = "P" Or txtitemstatus.Text = "H") And rfqmode = "addprice" Then

        'strSql = "SELECT  0.00 as Price,0.00 as Fact, 0.00 as FPrice,Qty,Qty_Type as Type,Enq_Qty_IntCode as IntCode FROM ENQ_Qty_Details " & _
        '                            "WHERE  Enq_Detail_code = '" & txtenqdetailintcode.Text & "' " & _
        '                              "ORDER BY Qty"


        'End If
        'End If


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDQC As SqlDataAdapter = New SqlDataAdapter

        stockDQC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDQC.TableMappings.Add("Table", "Part")
        'get data
        stockDQC.Fill(stockDCQ)


        DataGridQtyview.DataSource = stockDCQ.Tables(0)
        sqlCon.Close()
        DataGridQtyview.Expand(-1)


    End Sub
    Private Sub fillcertificatecharges()

        DataGridCertificateChargesview.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet
        strSql = ""
        'If txtitemstatus.Text = "P" And rfqmode = "" Then


        'strSql = "SELECT  0.00 as Protoprice,0.00 as ProdPrice,Certificates from ENQ_EnqWise_Certificates " & _
        '                            "WHERE  Enq_Detail_code = " & txtenqdetailintcode.Text & " and Enq_Reg_NO = " & txtregno.Text & " "
        'ElseIf (txtitemstatus.Text = "H" Or txtitemstatus.Text = "C" Or txtitemstatus.Text = "F") And rfqmode = "" Then

        strSql = "Select Proto_Price, Prod_Price,Certificates from ENQ_EnqWise_Certificates_Charges " & _
         "WHERE  Enq_Detail_code = " & txtenqdetailcode.Text & " and Enq_Reg_NO = " & txtregno.Text & " "


        'End If

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDQC As SqlDataAdapter = New SqlDataAdapter

        stockDQC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDQC.TableMappings.Add("Table", "Part")
        'get data
        stockDQC.Fill(stockDCQ)


        DataGridCertificateChargesview.DataSource = stockDCQ.Tables(0)
        sqlCon.Close()
        DataGridCertificateChargesview.Expand(-1)

        '        countqty = stockDCQ.Tables(0).Rows.Count



    End Sub



    Private Sub txtspecial_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CheckBoxRead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim check As String

        If CheckBoxRead.Checked = True Then
            check = "Cust"
        Else
            check = "N"
        End If

        Dim strsql As String
        Dim strsql1 As String
        Dim cmSQL As SqlCommand
        Dim cmsql1 As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drsql1 As SqlDataReader
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        strsql = "SELECT Read_Status FROM ENQ_PriceView_ReadStatus where Read_Status = 'Cust' and  Enq_Reg_NO = " & txtregno.Text & " and Enq_Detail_code = " & txtenqdetailcode.Text & ""
        cnSQL.Open()
        cmSQL = New SqlCommand(strsql, cnSQL)


        drsql1 = cmSQL.ExecuteReader()

        If drsql1.Read() Then

            If IsDBNull(drsql1.Item(0)) Then
                strsql1 = "insert ENQ_PriceView_ReadStatus values(" & txtregno.Text & "," & txtenqdetailcode.Text & ",'" & check & "')"
            Else
                '
                strsql1 = "update ENQ_PriceView_ReadStatus set Read_Status = '" & check & "' WHERE Enq_Reg_NO = " & txtregno.Text & " and Enq_Detail_code = " & txtenqdetailcode.Text & " "
            End If
        Else
            strsql1 = "insert ENQ_PriceView_ReadStatus values(" & txtregno.Text & "," & txtenqdetailcode.Text & ",'" & check & "')"

        End If

        cnSQL1.Open()
        cmsql1 = New SqlCommand(strsql1, cnSQL1)

        If cmsql1.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot Market it as read " & strsql1, MsgBoxStyle.Exclamation, "Error!")
            Exit Sub
        End If

    End Sub

    Private Sub ButtonRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRefresh.Click
        selection = ""
        ZCLASS = "NO"


        datagridRFQView.Enabled = True
        GroupBoxPrice.Visible = True
        GroupBoxPrice.Enabled = True

        lblpart.Visible = False
        txtpartbyApl.Visible = False
        txtAplSpecial.Visible = False
        GroupBoxPart.Visible = False

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        Dim stockDC As DataSet = New DataSet


        If usertype = "S" Then

            strSQL = "SELECT Enq_Detail_code as DetailCode,Enq_Reg_NO, Enq_Reg_date, CustomerID, CustomerName, CSR, Sl_no, PartNumber, PartDescription,  MOQ, SPU, LeadTime, Type, Stock_Avble," & _
                      "Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, " & _
                      "Remarks as SpecialInstructions FROM TSS_Enquiry_RFQView_New a " & _
                      " WHERE a.Read_Status ='N'  ORDER BY a.Enq_Reg_NO,a.Sl_no "


        Else


            'ZCLASSHANDLING()



            'If RadioButtonunread.Checked = True And ZCLASS = "NO" Then
            If RadioButtonread.Checked = True Then
                strSQL = "SELECT Enq_Detail_code as DetailCode,Enq_Reg_NO, Enq_Reg_date, CustomerID, CustomerName, CSR, Sl_no, PartNumber, PartDescription,  MOQ, SPU, LeadTime, Type, Stock_Avble," & _
                         "Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, " & _
                         "Remarks as SpecialInstructions FROM TSS_Enquiry_RFQView_New a " & _
                         " WHERE a.INS_SALES_CDE = '" & username & "' AND  a.Read_Status ='Cust'  ORDER BY a.Enq_Reg_NO,a.Sl_no "
                'ElseIf RadioButtonunread.Checked = True And ZCLASS = "YES" Then

            ElseIf RadioButtonunread.Checked = True Then
                strSQL = "SELECT Enq_Detail_code as DetailCode,Enq_Reg_NO, Enq_Reg_date, CustomerID, CustomerName, CSR, Sl_no, PartNumber, PartDescription,  MOQ, SPU, LeadTime, Type, Stock_Avble," & _
                                 "Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, " & _
                                 "Remarks as SpecialInstructions FROM TSS_Enquiry_RFQView_New a " & _
                                 " WHERE a.INS_SALES_CDE = '" & username & "' AND  a.Read_Status ='N' ORDER BY a.Enq_Reg_NO,a.Sl_no "


                'ElseIf RadioButtonread.Checked = True And ZCLASS = "NO" Then

                '    strSQL = "SELECT Enq_Detail_code as DetailCode,Enq_Reg_NO, Enq_Reg_date, CustomerID, CustomerName, CSR, Sl_no, PartNumber, PartDescription,  MOQ, SPU, LeadTime, Type, Stock_Avble," & _
                '                        "Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, " & _
                '                        "Remarks as SpecialInstructions FROM TSS_Enquiry_RFQView a " & _
                '                        " WHERE a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and ISR = '" & username & "')AND  a.Read_Status ='Cust' AND a.Class3 IN('K','I') ORDER BY a.Enq_Reg_NO,a.Sl_no "

                'ElseIf RadioButtonread.Checked = True And ZCLASS = "YES" Then

                '    strSQL = "SELECT Enq_Detail_code as DetailCode,Enq_Reg_NO, Enq_Reg_date, CustomerID, CustomerName, CSR, Sl_no, PartNumber, PartDescription,  MOQ, SPU, LeadTime, Type, Stock_Avble," & _
                '                        "Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, " & _
                '                        "Remarks as SpecialInstructions FROM TSS_Enquiry_RFQView a " & _
                '                        " WHERE a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and (ZCLASS1 = '" & username & "' OR ZCLASS2 = '" & username & "' )) AND  a.Read_Status ='Cust' and a.Class3 NOT IN('K','I') ORDER BY a.Enq_Reg_NO,a.Sl_no "


                '  ElseIf RadioButtonall.Checked = True And ZCLASS = "NO" Then

            ElseIf RadioButtonall.Checked = True Then

                strSQL = "SELECT Enq_Detail_code as DetailCode,Enq_Reg_NO, Enq_Reg_date, CustomerID, CustomerName, CSR, Sl_no, PartNumber, PartDescription,  MOQ, SPU, LeadTime, Type, Stock_Avble," & _
                                    "Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, " & _
                                    "Remarks as SpecialInstructions FROM TSS_Enquiry_RFQView_New a " & _
                                    " WHERE a.INS_SALES_CDE = '" & username & "' ORDER BY a.Enq_Reg_NO,a.Sl_no "

                'ElseIf RadioButtonall.Checked = True And ZCLASS = "YES" Then

                '    strSQL = "SELECT Enq_Detail_code as DetailCode,Enq_Reg_NO, Enq_Reg_date, CustomerID, CustomerName, CSR, Sl_no, PartNumber, PartDescription,  MOQ, SPU, LeadTime, Type, Stock_Avble," & _
                '                        "Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, " & _
                '                        "Remarks as SpecialInstructions FROM TSS_Enquiry_RFQView a " & _
                '                        " WHERE a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and (ZCLASS1 = '" & username & "' OR ZCLASS2 = '" & username & "')) AND a.Class3 NOT IN('K','I') ORDER BY a.Enq_Reg_NO,a.Sl_no "

            End If
        End If

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        datagridRFQView.DataSource = stockDC.Tables(0)
        cnSQL.Close()
        datagridRFQView.Expand(-1)

    End Sub

    Private Sub ToolDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCustomerRefresh.Click

        selection = "cust"
        ZCLASS = "NO"
        GroupBoxPrice.Enabled = False
        GroupBoxPrice.Visible = False

        lblpart.Visible = False
        txtpartbyApl.Visible = False
        txtAplSpecial.Visible = False
        GroupBoxPart.Visible = False


        Dim strSQL As String
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL As New SqlCommand
        Dim stockDC As DataSet = New DataSet


        If usertype = "S" Then

            strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, Name, Addr1, Addr2, Addr3, City, CustomerID, CSR FROM ENQ_CustomerCreation_Q a " 


        Else


            'ZCLASSHANDLING()

            'If ZCLASS = "NO" Then
            '    strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, Name, Addr1, Addr2, Addr3, City, CustomerID, CSR FROM ENQ_CustomerCreation_Q a " & _
            '    "WHERE (a.CSR IN (SELECT CSR FROM dbo.ENQ_CSR WHERE ISR = '" & username & "')) AND a.Class3 in ('K','I')"
            'ElseIf ZCLASS = "YES" Then

            strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, Name, Addr1, Addr2, Addr3, City, CustomerID, CSR,Class1 FROM ENQ_CustomerCreation_Q a " & _
                "WHERE a.UserId = '" & username & "' "



            'End If


        End If

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL1)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL1.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        datagridRFQView.DataSource = stockDC.Tables(0)
        cnSQL1.Close()
        datagridRFQView.Expand(-1)

    End Sub

    Private Sub ButtonPartsRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPartsRefresh.Click
        GroupBoxPrice.Enabled = False
        GroupBoxPrice.Visible = False

        lblpart.Visible = True
        txtpartbyApl.Visible = True
        txtAplSpecial.Visible = True
        GroupBoxPart.Visible = True


        selection = "part"
        ZCLASS = "NO"


        Dim strSQL As String
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL As New SqlCommand
        Dim stockDC As DataSet = New DataSet


        If usertype = "S" Then

            strSQL = "SELECT RegNo,  [Reg.Date], CustomerID, CustomerName, City,  CSR,  " & _
        "SlNo, PartNumber, PartDescription, CustPartNumber, CustPartDescription, uom, Part_No_Appl_Sug,Item_Created_Date, Item_Created_By,Special_Inst_Apl " & _
        "FROM ENQ_Parts_Created_Q a"



        Else

            ' ZCLASSHANDLING()


            'If ZCLASS = "NO" Then
            '    strSQL = "SELECT RegNo,  [Reg.Date], CustomerID, CustomerName, City,  CSR,  " & _
            '    "SlNo, PartNumber, PartDescription, CustPartNumber, CustPartDescription, uom, Part_No_Appl_Sug,Item_Created_Date, Item_Created_By,Special_Inst_Apl " & _
            '    "FROM ENQ_Parts_Created_Q a Where a.INS_SALES_CDE  = '" & username & "' ORDER BY RegNo, SlNo"

            'ElseIf ZCLASS = "YES" Then
            strSQL = "SELECT RegNo,  [Reg.Date], CustomerID, CustomerName, City,  CSR,  " & _
         "SlNo, PartNumber, PartDescription, CustPartNumber, CustPartDescription, uom, Part_No_Appl_Sug,Item_Created_Date, Item_Created_By,Special_Inst_Apl " & _
         "FROM ENQ_Parts_Created_Q a Where a.INS_SALES_CDE = '" & username & "' ORDER BY RegNo, SlNo"



            ' End If

        End If

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL1)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL1.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        datagridRFQView.DataSource = stockDC.Tables(0)
        cnSQL1.Close()
        datagridRFQView.Expand(-1)




    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxPart.Enter

    End Sub
    Private Sub ZCLASSHANDLING()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select ZCLASS1,ZCLASS2 from ENQ_CSR where ZCLASS1 = '" & username & "' or ZCLASS2 = '" & username & "'"
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

    Private Sub CheckBoxRead_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxRead.CheckedChanged

    End Sub
End Class