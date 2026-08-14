Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
'Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection




Public Class RFQPriceViewQuick
    Dim ZCLASS As String

    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Private Sub ButtonRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRefresh.Click
        '-copy

        DataGridViewPriceView.ClearSelection()


        'ZCLASS = "NO"

        If RadioButtonAll.Checked = False And RadioButtonPendingPrice.Checked = False And RadioButtonQuotePending.Checked = False And RadioButtonSingle.Checked = False And RadioButtonPurNeed.Checked = False And RadioButtonId.Checked = False And RadioButtonName.Checked = False And RadioButtonPartNumber.Checked = False Then

            MsgBox("you have to select any one option", vbInformation)
            Exit Sub


        End If

        DataGridViewPriceView.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        'ZCLASSHANDLINGQ()

        If usertype = "S" Then
            If RadioButtonAll.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                         "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                         " FROM         TSS_RFQPriceViewQuickNew where Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY Enq_Reg_NO, Sl_no"

            ElseIf RadioButtonPendingPrice.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                 "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                " FROM  TSS_RFQPriceViewQuickNew where (ItemStatus NOT IN ('J', 'T', 'R', 'U')) AND (Req = 'Price' OR Req = 'Both') AND (Enq_Status = 'Accepted') and Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY Enq_Reg_NO, Sl_no"





            ElseIf RadioButtonQuotePending.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                " FROM  TSS_RFQPriceViewQuickNew where FinalPrice > 0 and (len(CONumber) < 3 or CONumber is null) and Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY Enq_Reg_NO, Sl_no"

            ElseIf RadioButtonSingle.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                    "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl  " & _
                    " FROM  TSS_RFQPriceViewQuickNew where Enq_Reg_NO = " & txtRegNo.Text & "  ORDER BY Enq_Reg_NO, Sl_no"
            ElseIf RadioButtonPurNeed.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, AlternateMtrl,Qty, Qty_Type,  " & _
                    "Remarks as Special_Instruction_from_Purchase  FROM  TSS_RFQPriceViewQuickNew where (FinalPrice = 0 or FinalPrice is null) and Len(Remarks) > 1 and Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY Enq_Reg_NO, Sl_no"


            ElseIf RadioButtonPartNumber.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                         "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                         " FROM         TSS_RFQPriceViewQuickNew where PartNumber  = '" & txtPartNumber.Text & "' ORDER BY Enq_Reg_NO, Sl_no"

            ElseIf RadioButtonName.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                        "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                       " FROM         TSS_RFQPriceViewQuickNew where   (CustomerID  = '" & lblCustName.Text & "') ORDER BY Enq_Reg_NO, Sl_no"

            ElseIf RadioButtonId.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                         "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                         " FROM   TSS_RFQPriceViewQuickNew where (CustomerID  = '" & txtCustID.Text & "') ORDER BY Enq_Reg_NO, Sl_no"



            End If

        ElseIf usertype <> "S" Then
            'And ZCLASS = "YES" Then

            If RadioButtonAll.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                         "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl  " & _
                         " FROM         TSS_RFQPriceViewQuickNew a " & _
                         "WHERE a.INS_SALES_CDE = '" & username & "' AND a.Enq_Reg_date >= '" & dtpfrdate.Value & "' and a.Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"

                'WHERE a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and (ZCLASS1 = '" & username & "' OR ZCLASS2 = '" & username & "' )) and a.Class3 NOT IN('K','I')" & _
                '             " AND a.Enq_Reg_date >= '" & dtpfrdate.Value & "' and a.Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"

            ElseIf RadioButtonPendingPrice.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                 "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl  " & _
                " FROM  TSS_RFQPriceViewQuickNew a " & _
                "WHERE a.INS_SALES_CDE = '" & username & "' and (ItemStatus NOT IN ('J', 'T', 'R', 'U')) AND (Req = 'Price' OR Req = 'Both') AND (Enq_Status = 'Accepted') and a.Enq_Reg_date >= '" & dtpfrdate.Value & "' and a.Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"

                'where " & _"
                ' "a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and (ZCLASS1 = '" & username & "' OR ZCLASS2 = '" & username & "' )) and a.Class3 NOT IN('K','I')" & _
                '"and FinalPrice = 0 and  Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"


            ElseIf RadioButtonQuotePending.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl  " & _
                " FROM  TSS_RFQPriceViewQuickNew a " & _
                  "WHERE a.INS_SALES_CDE = '" & username & "' and  FinalPrice > 0 and (len(CONumber) < 3 or CONumber is null) and  a.Enq_Reg_date >= '" & dtpfrdate.Value & "' and a.Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"

                'INDIRA MODIFY


                ' "a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and (ZCLASS1 = '" & username & "' OR ZCLASS2 = '" & username & "' )) and a.Class3 NOT IN('K','I')" & _
                ' " and  FinalPrice > 0 and (len(CONumber) < 3 or CONumber is null) and Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"

            ElseIf RadioButtonSingle.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, AlternateMtrl,Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                    "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl   " & _
                    " FROM  TSS_RFQPriceViewQuickNew a " & _
                   "WHERE a.INS_SALES_CDE = '" & username & "' and a.Enq_Reg_NO = " & txtRegNo.Text & "  ORDER BY a.Enq_Reg_NO, a.Sl_no"

                'where " & _"
                '    "a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and (ZCLASS1 = '" & username & "' OR ZCLASS2 = '" & username & "' )) and a.Class3 NOT IN('K','I')" & _
                '   " and a.Enq_Reg_NO = " & txtRegNo.Text & "  ORDER BY a.Enq_Reg_NO, a.Sl_no"


            ElseIf RadioButtonPurNeed.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, AlternateMtrl, Qty, Qty_Type,  " & _
                    "Remarks as Special_Instruction_from_Purchase  FROM  TSS_RFQPriceViewQuickNew a  " & _
                "WHERE a.INS_SALES_CDE = '" & username & "' and  (FinalPrice = 0 or FinalPrice is null) and Len(Remarks) > 1 and Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"


            ElseIf RadioButtonPartNumber.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                         "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                         " FROM         TSS_RFQPriceViewQuickNew where TSS_RFQPriceViewQuickNew.INS_SALES_CDE = '" & username & "'  AND PartNumber  = '" & txtPartNumber.Text & "' ORDER BY Enq_Reg_NO, Sl_no"

            ElseIf RadioButtonName.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                         "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                         " FROM         TSS_RFQPriceViewQuickNew where TSS_RFQPriceViewQuickNew.INS_SALES_CDE = '" & username & "' and CustomerID  = '" & lblCustName.Text & "' ORDER BY Enq_Reg_NO, Sl_no"

            ElseIf RadioButtonId.Checked = True Then

                strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc,AlternateMtrl, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
                         "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty,Source_Mtrl " & _
                         " FROM         TSS_RFQPriceViewQuickNew where TSS_RFQPriceViewQuickNew.INS_SALES_CDE = '" & username & "' and CustomerID  = '" & txtCustID.Text & "' ORDER BY Enq_Reg_NO, Sl_no"



            End If


            'ElseIf usertype <> "S" And ZCLASS = "NO" Then


            '    If RadioButtonAll.Checked = True Then

            '        strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
            '                 "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty " & _
            '                 "FROM         TSS_RFQPriceViewQuick a " & _
            '                 "WHERE a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and ISR = '" & username & "')AND   a.Class3 IN('K','I')" & _
            '                 " AND a.Enq_Reg_date >= '" & dtpfrdate.Value & "' and a.Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"

            '    ElseIf RadioButtonPendingPrice.Checked = True Then

            '        strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
            '         "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty " & _
            '        " CSR FROM  TSS_RFQPriceViewQuick a where " & _
            '        " a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and ISR = '" & username & "')AND   a.Class3 IN('K','I')" & _
            '        "and FinalPrice = 0 and  Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"


            '    ElseIf RadioButtonQuotePending.Checked = True Then

            '        strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
            '        "Remarks as Special_Instruction_from_Purchase, Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty " & _
            '        " FROM  TSS_RFQPriceViewQuick a  where " & _
            '        "a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and ISR = '" & username & "')AND   a.Class3 IN('K','I')" & _
            '        " and  FinalPrice > 0 and (len(CONumber) < 3 or CONumber is null) and Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"

            '    ElseIf RadioButtonSingle.Checked = True Then

            '        strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, Qty, Qty_Type, FinalPrice as Price, MOQ, SPU, LeadTime, Type, Stock_Avble, " & _
            '            "Remarks as Special_Instruction_from_Purchase,Tools_YesNo, Proto_TotalCost, Proto_CustCost, Proto_LeadTime, Proto_Qty, Proto_LifeTool, Prod_TotalCost, Prod_CustCost, Prod_LeadTime, Prod_Qty " & _
            '            " FROM  TSS_RFQPriceViewQuick a where " & _
            '            "a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and ISR = '" & username & "')AND   a.Class3 IN('K','I')" & _
            '            " and a.Enq_Reg_NO = " & txtRegNo.Text & "  ORDER BY a.Enq_Reg_NO, a.Sl_no"

            '    ElseIf RadioButtonPurNeed.Checked = True Then

            '        strSQL = "SELECT Enq_Reg_NO, Enq_Reg_date, CONumber as Qtn_No,CustomerName, CSR, Sl_no, PartNumber, PartDescription as PartDesc, Qty, Qty_Type,  " & _
            '            "Remarks as Special_Instruction_from_Purchase  FROM  TSS_RFQPriceViewQuick a where " & _
            '            "a.CSR in (select CSR from ENQ_CSR where CSR = a.CSR and ISR = '" & username & "')AND   a.Class3 IN('K','I')" & _
            '            " and (FinalPrice = 0 or FinalPrice is null) and Len(Remarks) > 1 and Enq_Reg_date >= '" & dtpfrdate.Value & "' and Enq_Reg_date <= '" & dtptodate.Value & "' ORDER BY a.Enq_Reg_NO, a.Sl_no"


            '    End If

        End If


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewPriceView.DataSource = stockDC.Tables(0)
        cnSQL.Close()


    End Sub

    Private Sub DataGridViewPriceView_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewPriceView.CellContentClick

    End Sub

    Private Sub RFQPriceViewQuick_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Note_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Note.Click

    End Sub

    Private Sub RadioButtonAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonAll.CheckedChanged
        txtRegNo.Visible = False

    End Sub

    Private Sub RadioButtonPendingPrice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonPendingPrice.CheckedChanged
        txtRegNo.Visible = False

    End Sub

    Private Sub RadioButtonQuotePending_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonQuotePending.CheckedChanged
        txtRegNo.Visible = False
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonSingle.CheckedChanged
        txtRegNo.Visible = True
    End Sub
    Private Sub ZCLASSHANDLINGQ()

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


    Private Sub RadioButtonPurNeed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonPurNeed.CheckedChanged

    End Sub

    Private Sub txtCustID_DoubleClick(sender As Object, e As EventArgs)
        fillcustomer()
    End Sub

    Private Sub txtCustID_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub fillcustomer()
        DataGridCustomer1.Show()


        Dim sqlcon As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        txtCustID.Text = txtCustID.Text & "%"

        If usertype <> "S" Then



            If RadioButtonName.Checked = True Then
                strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.TSS_CustomerWise_CSR_ISR_EmailIDS " & _
                     "WHERE INS_SALES_CDE = '" & username & "' AND CustomerName like '" & txtCustID.Text & "' "

            Else

                strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.TSS_CustomerWise_CSR_ISR_EmailIDS " & _
                     "WHERE INS_SALES_CDE = '" & username & "' AND CustomerId like '" & txtCustID.Text & "' "

            End If


        ElseIf usertype = "S" Then


            If RadioButtonName.Checked = True Then
                strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.FS_Customer " & _
                     "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRELLEBORG%') AND (CustomerID NOT LIKE '0000%')AND CustomerName like '" & txtCustID.Text & "' " & _
                        "ORDER BY CustomerID"

            Else

                strSql = "SELECT CustomerID, CustomerName, CustomerCity, CSR FROM FSDBBR.dbo.FS_Customer " & _
                         "WHERE (CustomerName NOT LIKE 'TSS%') AND (CustomerName NOT LIKE 'TRELLEBORG%') AND (CustomerID NOT LIKE '0000%')AND CustomerID like '" & txtCustID.Text & "' " & _
                            "ORDER BY CustomerID"

            End If

        End If


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlcon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        stockDAC.TableMappings.Add("Table", "Customer")
        'get data
        stockDAC.Fill(stockDC)

        DataGridCustomer1.Visible = True

        DataGridCustomer1.Width = 650 '1150
        DataGridCustomer1.Height = 320 '800



        stockDC.Tables(0).Columns(0).ColumnName = "CustomerID"
        stockDC.Tables(0).Columns(1).ColumnName = "CustomerName"
        stockDC.Tables(0).Columns(2).ColumnName = "CustomerCity"
        stockDC.Tables(0).Columns(3).ColumnName = "CSR"


        'Dim col1 As DataGridViewColumn = DataGridCustomer1.Columns(1) 'CUSTOMER id
        'col1.Width = 100

        'Dim col2 As DataGridViewColumn = DataGridCustomer.Columns(2) 'CUSTOMER name
        'col2.Width = 300


        DataGridCustomer1.DataSource = stockDC.Tables(0)
        sqlcon.Close()
        DataGridCustomer1.Expand(-1)


    End Sub

    Private Sub txtPartNumber_DoubleClick(sender As Object, e As EventArgs)
        fillpart()
    End Sub

    Private Sub txtPartNumber_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub fillpart()
        DataGridPartNumbers.Show()


        Dim sqlcon As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        txtPartNumber.Text = txtPartNumber.Text & "%"


        strSql = "SELECT  ItemNumber, ItemDescription, ItemUM, MakeBuyCode, ItemStatus, ItemKey FROM FSDBBR.dbo.FS_Item WHERE ItemNumber like '" & txtPartNumber.Text & "' "


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlcon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        sqlcon.Open()

        stockDAC.TableMappings.Add("Table", "Part")
        'get data
        stockDAC.Fill(stockDC)

        DataGridPartNumbers.Width = 650 '1150
        DataGridPartNumbers.Height = 320 '800



        stockDC.Tables(0).Columns(0).ColumnName = "PartNumber"
        stockDC.Tables(0).Columns(1).ColumnName = "PartDescription"
        stockDC.Tables(0).Columns(2).ColumnName = "ItemUM"
        stockDC.Tables(0).Columns(3).ColumnName = "MakeBuyCode"
        stockDC.Tables(0).Columns(4).ColumnName = "ItemStatus"

        DataGridPartNumbers.DataSource = stockDC.Tables(0)
        sqlcon.Close()
        DataGridPartNumbers.Expand(-1)

    End Sub

    Private Sub DataGridCustomer_CurrentCellChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridCustomer_DoubleClick(sender As Object, e As EventArgs)

    End Sub


    Private Sub DataGridCustomer_Navigate(sender As Object, ne As NavigateEventArgs)

    End Sub

    Private Sub DataGridPartNumbers_CurrentCellChanged(sender As Object, e As EventArgs) Handles DataGridPartNumbers.CurrentCellChanged
        Dim a As Integer
        'Dim custid As String



        a = DataGridPartNumbers.CurrentCell.ColumnNumber()

        If a = 0 Then
            txtPartNumber.Text = DataGridPartNumbers.Item(DataGridPartNumbers.CurrentCell)

            txtPartNumber.Enabled = False


            'txtCustomer.Text = DataGridCustomer.Item(


        Else
            MsgBox("Click on Partnumber to select ", vbInformation)
            Exit Sub
        End If

        DataGridPartNumbers.Hide()

    End Sub

    Private Sub DataGridPartNumbers_Navigate(sender As Object, ne As NavigateEventArgs) Handles DataGridPartNumbers.Navigate

    End Sub

    Private Sub txtCustID_DoubleClick1(sender As Object, e As EventArgs) Handles txtCustID.DoubleClick
        fillcustomer()
    End Sub

    Private Sub txtCustID_TextChanged_1(sender As Object, e As EventArgs) Handles txtCustID.TextChanged

    End Sub

    Private Sub txtPartNumber_DoubleClick1(sender As Object, e As EventArgs) Handles txtPartNumber.DoubleClick
        fillpart()
    End Sub

    Private Sub txtPartNumber_TextChanged_1(sender As Object, e As EventArgs) Handles txtPartNumber.TextChanged

    End Sub

    Private Sub DataGridCustomer_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridCustomer_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)

        ' txtCustID.Text = DataGridCustomer.CurrentRow.Cells(0).Value.ToString
        'lblCustName.Text = DataGridCustomer.CurrentRow.Cells(1).Value.ToString



    End Sub

    Private Sub DataGridCustomer1_CurrentCellChanged(sender As Object, e As EventArgs) Handles DataGridCustomer1.CurrentCellChanged
        Dim a As Integer
        'Dim custid As String



        a = DataGridCustomer1.CurrentCell.ColumnNumber()

        If a = 0 Then

            If RadioButtonName.Checked = True Then

                txtCustID.Text = DataGridCustomer1.Item(DataGridCustomer1.CurrentCell.RowNumber, 1)

                lblCustName.Text = DataGridCustomer1.Item(DataGridCustomer1.CurrentCell)

            ElseIf RadioButtonId.Checked = True Then

                txtCustID.Text = DataGridCustomer1.Item(DataGridCustomer1.CurrentCell)

                lblCustName.Text = DataGridCustomer1.Item(DataGridCustomer1.CurrentCell.RowNumber, 1)



            End If


            txtCustID.Enabled = False


            'txtCustomer.Text = DataGridCustomer.Item(


        Else
            MsgBox("Click on CustomerID to select the customer", vbInformation)
            Exit Sub
        End If

        DataGridCustomer1.Hide()

        '2.datagrid1.item(0,0)<-----it gets the first column/row data of your datagrid
        '3. 4.'if you want selected
        '5.datagrid1.item(datagrid1.currentcell.rownumber,0)<---it gets the selected row and the first column 

    End Sub

    Private Sub DataGridCustomer1_Navigate(sender As Object, ne As NavigateEventArgs) Handles DataGridCustomer1.Navigate

    End Sub

    Private Sub RadioButtonId_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonId.CheckedChanged

        txtCustID.Enabled = True
        txtCustID.Text = ""
        lblCustName.Text = ""
        txtPartNumber.Text = ""


    End Sub

    Private Sub RadioButtonName_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonName.CheckedChanged
        txtCustID.Enabled = True

        txtCustID.Text = ""
        lblCustName.Text = ""
        txtPartNumber.Text = ""
    End Sub

    Private Sub RadioButtonPartNumber_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonPartNumber.CheckedChanged
        txtCustID.Text = ""
        lblCustName.Text = ""
    End Sub
End Class