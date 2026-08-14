
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection

Public Class WHMaterialIssue
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DTPIssDt.Format = DateTimePickerFormat.Custom
        DTPIssDt.CustomFormat = "dd/MM/yyyy"
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    ' Sub Show(p1 As Integer)
    '    Throw New NotImplementedException
    'End Sub

    Private Sub txtReqNo_DoubleClick(sender As Object, e As EventArgs) Handles txtMatIssNo.DoubleClick

        If mode = "Cancel" Then

            ShowDCNumbers()

        End If

        ' ApprovedPendingReq()



    End Sub

    Private Sub txtReqNo_MouseEnter(sender As Object, e As EventArgs) Handles txtMatIssNo.MouseEnter

    End Sub

    Private Sub txtReqNo_TextChanged(sender As Object, e As EventArgs) Handles txtMatIssNo.TextChanged

    End Sub
    Private Sub ApprovedPendingReq()

        datagridReqPending.Visible = True
        datagridReqPending.BringToFront()
        ' datagridReqPending.Location.X.MaxValue = 574

        datagridReqPending.Location = New System.Drawing.Point(468, 10)

        datagridReqPending.Width = 542
        datagridReqPending.Height = 149

        datagridReqPending.Enabled = True

        'RBToolNo.Checked = True
        'RadioButtonGroup.Checked = True
        'RadioButtonVendorYes.Checked = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        Dim stockDC As DataSet = New DataSet

        'not in TSS_WH_MaterialIssueDetail to be included - already issued material
        'ONLY HEADER NEED TO BE CALLED - DISTINCT MAT REQ NO

        strSQL = "SELECT distinct [MatReq_no],[MatReq_Date],[FS_NonFS],[Type_Dept],[Sub_Div] ,[Cell] ,[MONumber], [Remarks],[User_Id] " & _
                 "FROM [FSPrograms].[dbo].[TSS_WH_MatReqAppComp_P]  where  [1st_AppStatus] = 'YES' AND  [MatReq_no] NOT IN ( SELECT [MatReq_no] FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueHeader] WITH (NOLOCK))"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        datagridReqPending.DataSource = stockDC.Tables(0)

        cnSQL.Close()
        '  TSS_WH_MatReqPendingAppDet_P()
    End Sub

    Private Sub datagridReqPending_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridReqPending.CellContentClick

    End Sub

    Private Sub datagridReqPending_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles datagridReqPending.RowHeaderMouseClick

        txtMatReqNo.Text = datagridReqPending.CurrentRow.Cells(0).Value
        '    strSQL = "SELECT [MatReq_no],[MatReq_Date],[Type_Dept],[Sub_Div] ,[Cell] ,[MONumber] ,[Remarks], [FS_NonFS], [Slno] ,[Part_Number] ,[Part_Desc] ,[Qty] , DetailRemarks, " & _
        '              "FROM [FSPrograms].[dbo].[TSS_WH_MatReqPendingAppDet_P] where ([RejLines])IS NULL"


        ComboBoxdept.Text = datagridReqPending.CurrentRow.Cells(3).Value
        ComboBoxSD.Text = datagridReqPending.CurrentRow.Cells(4).Value
        ComboBoxCell.Text = datagridReqPending.CurrentRow.Cells(5).Value
        txtMO.Text = datagridReqPending.CurrentRow.Cells(6).Value
        txtRemarks.Text = datagridReqPending.CurrentRow.Cells(7).Value

        If datagridReqPending.CurrentRow.Cells(2).Value = "N" Then
            RBNonFSItem.Checked = True
            RBFSItem.Checked = False
        ElseIf datagridReqPending.CurrentRow.Cells(2).Value = "F" Then
            RBNonFSItem.Checked = False
            RBFSItem.Checked = True
        End If
        GroupBox4.Enabled = False


        LoadMatReqDetails()

        'If IsDBNull(datagridEnquiryPending1.Item(datagridEnquiryPending1.CurrentCell.RowNumber, 10)) Then

        '     If IsDBNull(datagridEnquiryPending.CurrentRow.Cells(10).Value.ToString) Then
        'RadioButtonDomestic.Checked = True
        'RadioButtonExport.Checked = False
        'ElseIf Trim(datagridEnquiryPending.CurrentRow.Cells(10).Value.ToString) = "Domestic" Then
        'RadioButtonDomestic.Checked = True
        'RadioButtonExport.Checked = False
        'ElseIf Trim(datagridEnquiryPending.CurrentRow.Cells(10).Value.ToString) = "Export" Then
        'RadioButtonDomestic.Checked = False
        'RadioButtonExport.Checked = True
        'End If


    End Sub

    Private Sub LoadMatReqDetails()

        DataGridViewMaterialReq.Visible = True
        DataGridViewMaterialReq.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        Dim stockDC As DataSet = New DataSet

        'issued should not come again

        '    strSQL = "SELECT [Slno] ,[Part_Number] AS ItemNumber ,[Part_Desc] as Description ,[Qty] ,DetailRemarks, '' as IssuedQty, '' as Notes " & _
        '               "FROM [FSPrograms].[dbo].[TSS_WH_MatReqAppCompDetail_P] where ([RejLines]) IS NULL and [MatReq_no] = " & txtMatReqNo.Text & " "


        If RBFSItem.Checked = True Then


            '   strSQL = "SELECT a.Slno, a.Part_Number AS ItemNumber ,a.Part_Desc as Description ,a.Qty as Reqd_Qty, b.UOM, b.LotNumber, b.LotReceiptDate as LotDate, b.Stockroom as STR, b.Bin, b.qty as Stock_Qty," & _
            '           " case when a.Qty >= b.qty then b.qty else a.Qty end 'Issued_Qty','' as Notes, a.DetailRemarks, b.TotalRolledCost as 'Cost' FROM [FSPrograms].[dbo].[TSS_WH_MatReqAppCompDetail_P]a inner join [dbo].[TSS_WH_StockFSItems_P] b on " & _
            '          " a.Part_Number = b.ItemNumber where a.Issue_Ret = 'I' AND  a.MatReq_no = " & txtMatReqNo.Text & " ORDER BY a.Slno, b.LotReceiptDate Asc"


            strSQL = "SELECT Slno, ItemNumber, Description, Qty as Reqd_Qty, UOM, LotNumber, LotReceiptDate as LotDate,Stockroom as STR,Bin, StockQty as Stock_Qty," & _
                     "NewIssueQty as 'Issued_Qty','' as Notes, DetailRemarks, isnull(TotalRolledCost,0) as 'Cost',IssueBalance FROM [FSPrograms].[dbo].[TSS_WH_Stock_Running_Balance] " & _
                    "where NewIssueQty >0 and  Issue_Ret = 'I' AND  MatReq_no = " & txtMatReqNo.Text & " ORDER BY Slno, LotReceiptDate Asc"


        ElseIf RBNonFSItem.Checked = True Then


            strSQL = "SELECT Slno, ItemNumber, Description, Qty as Reqd_Qty, UOM, LotNumber, LotReceiptDate as LotDate,Stockroom as STR,Bin,StockQty as Stock_Qty," & _
                  "NewIssueQty as 'Issued_Qty','' as Notes, DetailRemarks, isnull(TotalRolledCost,0) as 'Cost',IssueBalance FROM [FSPrograms].[dbo].[TSS_WH_Stock_Running_Balance_NonFSItem] " & _
                 "where NewIssueQty > 0 and Issue_Ret = 'I' AND  MatReq_no = " & txtMatReqNo.Text & " ORDER BY Slno, LotReceiptDate Asc"


        End If


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewMaterialReq.DataSource = stockDC.Tables(0)


        DataGridViewMaterialReq.Columns("Slno").ReadOnly = True
        DataGridViewMaterialReq.Columns("Slno").Width = 45

        DataGridViewMaterialReq.Columns("ItemNumber").ReadOnly = True
        DataGridViewMaterialReq.Columns("ItemNumber").Width = 130
        DataGridViewMaterialReq.Columns("Description").ReadOnly = True
        DataGridViewMaterialReq.Columns("Description").Width = 130

        DataGridViewMaterialReq.Columns("Reqd_Qty").ReadOnly = True
        DataGridViewMaterialReq.Columns("Reqd_Qty").Width = 80
        DataGridViewMaterialReq.Columns("Reqd_Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewMaterialReq.Columns("Reqd_Qty").DefaultCellStyle.Format = "N2"

        DataGridViewMaterialReq.Columns("UOM").ReadOnly = True
        DataGridViewMaterialReq.Columns("UOM").Width = 40

        DataGridViewMaterialReq.Columns("LotNumber").ReadOnly = True
        DataGridViewMaterialReq.Columns("LotNumber").Width = 100

        DataGridViewMaterialReq.Columns("LotDate").ReadOnly = True
        DataGridViewMaterialReq.Columns("LotDate").Width = 80


        DataGridViewMaterialReq.Columns("STR").ReadOnly = True
        DataGridViewMaterialReq.Columns("STR").Width = 50

        DataGridViewMaterialReq.Columns("Bin").ReadOnly = True
        DataGridViewMaterialReq.Columns("Bin").Width = 80

        DataGridViewMaterialReq.Columns("Stock_Qty").ReadOnly = True
        DataGridViewMaterialReq.Columns("Stock_Qty").Width = 80
        DataGridViewMaterialReq.Columns("Stock_Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewMaterialReq.Columns("Stock_Qty").DefaultCellStyle.Format = "N2"

        DataGridViewMaterialReq.Columns("Issued_Qty").ReadOnly = False
        DataGridViewMaterialReq.Columns("Issued_Qty").Width = 80
        DataGridViewMaterialReq.Columns("Issued_Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewMaterialReq.Columns("Issued_Qty").DefaultCellStyle.Format = "N2"
        DataGridViewMaterialReq.Columns("Issued_Qty").HeaderCell.Style.BackColor = Color.Gray

        DataGridViewMaterialReq.Columns("Notes").ReadOnly = False
        DataGridViewMaterialReq.Columns("Notes").Width = 100
        DataGridViewMaterialReq.Columns("Notes").HeaderCell.Style.BackColor = Color.Gray


        DataGridViewMaterialReq.Columns("DetailRemarks").ReadOnly = True
        DataGridViewMaterialReq.Columns("DetailRemarks").Width = 100

        DataGridViewMaterialReq.Columns("Cost").ReadOnly = True
        DataGridViewMaterialReq.Columns("Cost").Width = 0
        DataGridViewMaterialReq.Columns("Cost").Visible = False


        DataGridViewMaterialReq.Columns("IssueBalance").ReadOnly = True
        DataGridViewMaterialReq.Columns("IssueBalance").Width = 0
        DataGridViewMaterialReq.Columns("IssueBalance").Visible = False

        Dim mulstock As Integer
        mulstock = 0
        For x As Integer = 0 To DataGridViewMaterialReq.Rows.Count - 1
            For y As Integer = x + 1 To DataGridViewMaterialReq.Rows.Count - 1
                If DataGridViewMaterialReq.Rows(x).Cells(0).Value.ToString = DataGridViewMaterialReq.Rows(y).Cells(0).Value.ToString Then
                    DataGridViewMaterialReq.Rows(x).Cells(0).Style.BackColor = Color.LightBlue
                    DataGridViewMaterialReq.Rows(x).Cells(1).Style.BackColor = Color.LightBlue
                    DataGridViewMaterialReq.Rows(x).Cells(2).Style.BackColor = Color.LightBlue
                    DataGridViewMaterialReq.Rows(x).Cells(3).Style.BackColor = Color.LightBlue
                    mulstock = mulstock + 1

                    '     DataGridViewMaterialReq.Columns("Reqd_Qty").DefaultCellStyle.BackColor = Color.Gray
                    '      MsgBox("Duplicate Data!")
                    '     Exit Sub
                    'Else
                    '   save_data()
                    '  Me.Close()
                    x = x + 1
                End If
            Next
        Next


        If mulstock > 0 Then

            lblNote.Visible = True
            lblNote.Text = "Required quantities are repeated as stock need to be issued from different lots"


        End If



        cnSQL.Close()
        datagridReqPending.Visible = False
        txtMatReqNo.Enabled = False
        ComboBoxdept.Enabled = False
        ComboBoxSD.Enabled = False
        ComboBoxCell.Enabled = False
        txtMO.Enabled = False
        txtRemarks.Enabled = False
        DTPIssDt.Enabled = False
        '  txtMatIssNo.Enabled = False

    End Sub

    Private Sub btnReqSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        'save the record
        Dim checkdt As Date
        checkdt = Today

        Dim strsql2 As String
        Dim cmSQL As SqlCommand
        Dim msgb As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        msgb = MsgBox("Are you sure of saving ?", vbYesNo)

        If msgb = vbYes Then


            'check stock issued qty should not be greater than stock available for that part number and lot number

            For i As Integer = 0 To DataGridViewMaterialReq.RowCount - 1

                If Len(Me.DataGridViewMaterialReq.Rows(i).Cells("Notes").Value) > 100 Then
                    MsgBox("Notes should not be greater than 100 chrs", vbInformation)
                    Exit Sub
                End If

                If Me.DataGridViewMaterialReq.Rows(i).Cells("Issued_Qty").Value > Me.DataGridViewMaterialReq.Rows(i).Cells("Stock_Qty").Value Then

                    MsgBox("Issue qty Should not be greater than lot wise stock ", vbInformation)
                    Exit Sub

                End If
            Next
            'end of checking stock

            'generating regno

            transtype = "MatIss"
            transmode = "Add"
            nogenerate()

            txtMatIssNo.Text = issno

            If txtMatIssNo.Text > 0 Then

                cnSQL.Open()

                curdate = System.DateTime.Now()

                'save header table


                strsql2 = "insert  TSS_WH_MaterialIssueHeader values (" & txtMatIssNo.Text & ",'I','" & DTPIssDt.Value & "', '" & txtMatReqNo.Text & "','" & txtHeaderNotes.Text & "', '" & username & "', ' " & curdate & "')"
                cmSQL = New SqlCommand(strsql2, cnSQL)

                If cmSQL.ExecuteNonQuery() = 0 Then
                    MsgBox("Material Issue Header Details not saved " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                    '  txtRegNo.Text = 0
                    'Application.Exit()
                    Exit Sub
                End If


                'save detail table

                For i As Integer = 0 To DataGridViewMaterialReq.RowCount - 1
                    strsql2 = "insert TSS_WH_MaterialIssueDetail values (" & txtMatIssNo.Text & ",'" & DTPIssDt.Value & "','" & txtMatReqNo.Text & "','I','" & Me.DataGridViewMaterialReq.Rows(i).Cells("Slno").Value & "'," & _
                          "'" & Me.DataGridViewMaterialReq.Rows(i).Cells("ItemNumber").Value & "','" & Me.DataGridViewMaterialReq.Rows(i).Cells("Description").Value & "','" & Me.DataGridViewMaterialReq.Rows(i).Cells("UOM").Value & "' ," & Me.DataGridViewMaterialReq.Rows(i).Cells("Reqd_Qty").Value & " , " & _
                          "'" & Me.DataGridViewMaterialReq.Rows(i).Cells("LotNumber").Value & "'," & Me.DataGridViewMaterialReq.Rows(i).Cells("Issued_Qty").Value & "," & Me.DataGridViewMaterialReq.Rows(i).Cells("Cost").Value & ",'" & Me.DataGridViewMaterialReq.Rows(i).Cells("Notes").Value & "','" & username & "', ' " & curdate & "')"



                    cmSQL = New SqlCommand(strsql2, cnSQL)

                    If cmSQL.ExecuteNonQuery() = 0 Then
                        MsgBox("Material issue detail section not saved " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                        '  txtRegNo.Text = 0
                        'Application.Exit()
                        Exit Sub
                    End If
                Next

            End If


            transmode = "Update"
            nogenerate()

            MsgBox("Material issue slip generated ", vbInformation)
            btnSave.Enabled = False

        End If

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub MyGroupBox1_Enter(sender As Object, e As EventArgs) Handles MyGroupBox1.Enter

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub txtMatReqNo_DoubleClick(sender As Object, e As EventArgs) Handles txtMatReqNo.DoubleClick
        ApprovedPendingReq()
    End Sub

    Private Sub txtMatReqNo_TextChanged(sender As Object, e As EventArgs) Handles txtMatReqNo.TextChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        clearall()

        MsgBox("Click on Material Request No.", vbInformation)
        txtMatReqNo.Focus()
        Exit Sub

    End Sub

    Private Sub clearall()
        lblNote.Text = ""
        lblNote.Visible = False

        txtMatIssNo.Text = ""
        txtMatIssNo.Enabled = True
        txtMatReqNo.Text = ""
        txtMatReqNo.Enabled = True
        txtRemarks.Text = ""
        txtRemarks.Enabled = True
        txtHeaderNotes.Text = ""
        txtHeaderNotes.Enabled = True
        txtMO.Text = ""
        txtMO.Enabled = True
        ComboBoxCell.Text = ""
        ComboBoxCell.Enabled = True
        ComboBoxSD.Text = ""
        ComboBoxSD.Enabled = True
        ComboBoxdept.Text = ""
        ComboBoxdept.Enabled = True
        DataGridViewMaterialReq.Columns.Clear()
        btnSave.Enabled = True
    End Sub

    Private Sub ComboBoxdept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxdept.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
    Private Sub ShowDCNumbers()

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        datagridDC.Location = New System.Drawing.Point(411, 11)
        datagridDC.Width = 646
        datagridDC.Height = 125
        datagridDC.Visible = True
        datagridDC.Enabled = True

        'not in TSS_WH_MaterialIssueDetail to be included - already issued material
        'ONLY HEADER NEED TO BE CALLED - DISTINCT MAT REQ NO

        strSQL = "SELECT a.MatIssue_no as DC_No, a.MatIssue_Date as DC_Dt, a.MatReq_no,b.MatReq_Date,b.FS_NonFS, b.MONumber, b.Type_Dept, b.Sub_Div,b.Cell, a.Notes FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueHeader]a inner join " & _
                   "[FSPrograms].[dbo].[TSS_WH_MaterialRequestHeader]b ON  a.MatReq_no = b.MatReq_no   where a.Status = 'I' and  a.MatIssue_no NOT IN ( SELECT [MatIssueDC_no] FROM [FSPrograms].[dbo].[TSS_WH_MatRecptMainBldg_Header] WITH (NOLOCK)) order by a.MatIssue_no Asc"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        datagridDC.DataSource = stockDC.Tables(0)

        cnSQL.Close()

    End Sub



    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnCancel.Click

        If (txtMatIssNo.Text) = "" And txtMatIssNo.Text = "" Then
            clearall()
            btnSave.Enabled = False
            txtRemarks.Enabled = False
            txtMO.Enabled = False
            GroupBox4.Enabled = False
            ComboBoxdept.Enabled = False
            ComboBoxSD.Enabled = False
            ComboBoxCell.Enabled = False
            mode = "Cancel"
            MsgBox("Click on DC number ", vbInformation)
            lblNotes.Text = "Reason for Rej"
            Exit Sub
        End If

        If (txtMatIssNo.Text) <> "" And txtMatIssNo.Text <> "" Then

            If Len(txtHeaderNotes.Text) < 5 Then
                MsgBox("Reason for Rejection need to be entered", vbInformation)
                Exit Sub

            End If

            Dim msgb As String
            msgb = MsgBox("Are you sure of Cancelling this DC ?", vbYesNo)
            If msgb = vbYes Then

                'CHECK THIS INWARDED IN MAIN BUILDING 


                'END OF CHECKING INWARDED IN THE MAIN BUILDING


                'update header

                Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
                Dim cmSQL1 As SqlCommand
                Dim strSQL1 As String
                cnSQL1.Open()

                'header section update

                strSQL1 = "UPDATE [FSPrograms].[dbo].[TSS_WH_MaterialIssueHeader] set [Status] = 'X'  where MatIssue_no = " & txtMatIssNo.Text & ""

                cmSQL1 = New SqlCommand(strSQL1, cnSQL1)

                If cmSQL1.ExecuteNonQuery() = 0 Then
                    MsgBox("Cannot Cancel this DC, Contact Administrator " & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                    Exit Sub
                End If
                cnSQL1.Close()

                'update details

                cnSQL1.Open()
                '   For i As Integer = 0 To DataGridViewMaterialReq.RowCount - 1

                strSQL1 = "update [FSPrograms].[dbo].[TSS_WH_MaterialIssueDetail] set [Status] = 'X'  where MatIssue_no = " & txtMatIssNo.Text & ""

                cmSQL1 = New SqlCommand(strSQL1, cnSQL1)

                If cmSQL1.ExecuteNonQuery() = 0 Then
                    MsgBox("Cannot cancel this DC, Contact Administrator " & strSQL1, MsgBoxStyle.Exclamation, "Error!")

                    Exit Sub
                End If
                'Next

                cnSQL1.Close()

                'end of detail update

                'updating to 
                '  strsql2 = "insert  TSS_WH_MaterialIssueHeader values (" & txtMatIssNo.Text & ",'" & DTPIssDt.Value & "', '" & txtMatReqNo.Text & "','" & txtHeaderNotes.Text & "', '" & username & "', ' " & curdate & "')"


                cnSQL1.Open()

                strSQL1 = "insert TSS_WH_DC_Cancel_log values (" & txtMatIssNo.Text & ",'" & DTPIssDt.Value & "','" & curdate & "','" & username & "','" & txtHeaderNotes.Text & "')"

                cmSQL1 = New SqlCommand(strSQL1, cnSQL1)

                If cmSQL1.ExecuteNonQuery() = 0 Then

                    Exit Sub
                End If

                'END OF CANCELLATION LOG
                MsgBox("This DC is cancelled", vbInformation)

            End If

        End If

        mode = ""
        clearall()

    End Sub

    Private Sub datagridDC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridDC.CellContentClick

    End Sub

    Private Sub datagridDC_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridDC.CellContentDoubleClick

    End Sub

    Private Sub datagridDC_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles datagridDC.RowHeaderMouseClick

        '     strSQL = "SELECT a.MatIssue_no as DC_No, a.MatIssue_Date as DC_Dt, a.MatReq_no,b.MatReq_Date, b.FS_NonFS, b.Type_Dept,b.Sub_Div,b.Cell, a.Notes FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueHeader]a inner join " & _
        '                  "[FSPrograms].[dbo].[TSS_WH_MaterialRequestHeader]b ON  a.MatReq_no = b.MatReq_no   where a.MatIssue_no NOT IN ( SELECT [MatIssueDC_no] FROM [FSPrograms].[dbo].[TSS_WH_MatRecptMainBldg_Header] WITH (NOLOCK)) order by a.MatIssue_no Asc"


        txtMatIssNo.Text = datagridDC.CurrentRow.Cells(0).Value
        txtMatIssNo.Enabled = False

        DTPIssDt.Value = datagridDC.CurrentRow.Cells(1).Value
        DTPIssDt.Enabled = False

        txtMatReqNo.Text = datagridDC.CurrentRow.Cells(2).Value
        txtMatReqNo.Enabled = False


        If datagridDC.CurrentRow.Cells(4).Value = "F" Then
            RBFSItem.Checked = True
        ElseIf datagridDC.CurrentRow.Cells(4).Value = "N" Then
            RBNonFSItem.Checked = True
        End If
        txtMO.Text = datagridDC.CurrentRow.Cells(5).Value
        ComboBoxdept.Text = datagridDC.CurrentRow.Cells(6).Value
        ComboBoxSD.Text = datagridDC.CurrentRow.Cells(7).Value
        ComboBoxCell.Text = datagridDC.CurrentRow.Cells(8).Value
        txtRemarks.Text = datagridDC.CurrentRow.Cells(9).Value

       
        datagridDC.Visible = False

        MsgBox("Enter Reason for Rejection and click on Cancel", vbInformation)
        Exit Sub
    End Sub
End Class