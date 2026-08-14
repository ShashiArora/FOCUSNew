Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection

Public Class WHMatReceiptMfgBldg
    ' Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Private Sub WHMatReceiptMfgBldg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpReceiptdt.Format = DateTimePickerFormat.Custom
        dtpReceiptdt.CustomFormat = "dd/MM/yyyy"

        dtpMatReqDt.Format = DateTimePickerFormat.Custom
        dtpMatReqDt.CustomFormat = "dd/MM/yyyy"

        dtpissdt.Format = DateTimePickerFormat.Custom
        dtpissdt.CustomFormat = "dd/MM/yyyy"

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub btnImageClear_Click(sender As Object, e As EventArgs) Handles btnImageClear.Click
        If Len(txtMatIssueNo.Text) > 0 Then

            DataGridViewReceiptsMfg.Columns.Clear()


            DataGridViewReceiptsMfg.Visible = True
            DataGridViewReceiptsMfg.Enabled = True

            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim strSQL As String


            Dim stockDC As DataSet = New DataSet

            strSQL = "SELECT  [Slno],[Part_Number] as ItemNumber,[Part_Desc] as Description,[ReqQty],[Issued_Qty],[Remarks] ,[User_Id] ,[DateTime],[MatIssue_no],[MatIssue_Date],[MatReq_no],[HeaderNotes] " & _
                     "FROM [FSPrograms].[dbo].[TSS_WH_MatReceiptPendingMfg_P] WHERE [MatIssue_no] = '" & txtMatIssueNo.Text & "' order by [Slno]"

            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAC.TableMappings.Add("Table", "Enq")
            'get data
            stockDAC.Fill(stockDC)


            DataGridViewReceiptsMfg.DataSource = stockDC.Tables(0)

            txtMatIssueNo.Text = DataGridViewReceiptsMfg.CurrentRow.Cells(8).Value
            dtpissdt.Value = DataGridViewReceiptsMfg.CurrentRow.Cells(9).Value
            ' dtpMatReqDt.Text = DataGridViewReceiptsMfg.CurrentRow.Cells(9).Value
            txtMatReqNo.Text = DataGridViewReceiptsMfg.CurrentRow.Cells(10).Value
            txtrem.Text = DataGridViewReceiptsMfg.CurrentRow.Cells(11).Value

            '  DataGridViewReceiptsMfg.Columns("Slno").ReadOnly = True
            DataGridViewReceiptsMfg.Columns("Slno").Width = 45
            ' DataGridViewReceiptsMfg.Columns("ItemNumber").ReadOnly = True
            DataGridViewReceiptsMfg.Columns("ItemNumber").Width = 145
            DataGridViewReceiptsMfg.Columns("Description").ReadOnly = True
            DataGridViewReceiptsMfg.Columns("Description").Width = 170

            'DataGridViewReceiptsMfg.Columns("ReqQty").ReadOnly = True
            DataGridViewReceiptsMfg.Columns("ReqQty").Width = 100
            DataGridViewReceiptsMfg.Columns("ReqQty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridViewReceiptsMfg.Columns("ReqQty").DefaultCellStyle.Format = "N2"

            'DataGridViewReceiptsMfg.Columns("Issued_Qty").ReadOnly = True
            DataGridViewReceiptsMfg.Columns("Issued_Qty").Width = 100
            DataGridViewReceiptsMfg.Columns("Issued_Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridViewReceiptsMfg.Columns("Issued_Qty").DefaultCellStyle.Format = "N2"

            'DataGridViewReceiptsMfg.Columns("Remarks").ReadOnly = True
            DataGridViewReceiptsMfg.Columns("Remarks").Width = 200

            'DataGridViewReceiptsMfg.Columns("User_Id").ReadOnly = True
            DataGridViewReceiptsMfg.Columns("User_Id").Width = 70
            DataGridViewReceiptsMfg.Columns("DateTime").Width = 1
            DataGridViewReceiptsMfg.Columns("MatIssue_no").Width = 1
            DataGridViewReceiptsMfg.Columns("MatIssue_Date").Width = 1
            DataGridViewReceiptsMfg.Columns("MatReq_no").Width = 1
            DataGridViewReceiptsMfg.Columns("HeaderNotes").Width = 1
            cnSQL.Close()


        Else
            '   MsgBox("Pl select the Line", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub GroupBoxEdit_Enter(sender As Object, e As EventArgs) Handles GroupBoxEdit.Enter

    End Sub

    Private Sub btnRecAccept_Click(sender As Object, e As EventArgs) Handles btnRecAccept.Click

        'save the record
        Dim checkdt As Date
        checkdt = Today

        Dim strsql2 As String
        Dim cmSQL As SqlCommand
        Dim msgb As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        msgb = MsgBox("Are you sure of accepting the materials ?", vbYesNo)

        If msgb = vbYes Then


            'checking dc number

            Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL3 As SqlCommand
            Dim drSQL3 As SqlDataReader
            Dim strSQL3 As String


            strSQL3 = "SELECT [MatIssueDC_no] FROM [FSPrograms].[dbo].[TSS_WH_MatRecptMainBldg_Header] WHERE [MatIssueDC_no] = '" & txtMatIssueNo.Text & "'"

            cnSQL3.Open()
            cmSQL3 = New SqlCommand(strSQL3, cnSQL3)
            drSQL3 = cmSQL3.ExecuteReader()

            If drSQL3.Read() Then

                ' If IsDBNull(drSQL3.Item(0)) Then
                ' txtdc.Text = 1

                If Len(drSQL3.Item(0)) > 0 Then

                    MsgBox("This DC is already accepted", vbInformation)
                    Exit Sub
                Else
                End If

            End If
            cnSQL3.Close()

            'end of checking dc number









            'check feed back length

            For i As Integer = 0 To DataGridViewReceiptsMfg.RowCount - 1

                If Len(Me.DataGridViewReceiptsMfg.Rows(i).Cells("ItemWise_Feedback").Value) > 100 Then
                    MsgBox("Item wise feed back should not be greater than 100 chrs", vbInformation)
                    Exit Sub
                End If

            Next



            'end of feed back length


            'generating regno

            transtype = "ReceiptsMfg"
            transmode = "Add"
            nogenerate()

            txtRecNo.Text = receiptmfg

            If Val(txtRecNo.Text) > 0 Then

                cnSQL.Open()

                curdate = System.DateTime.Now()






                'save header table


                strsql2 = "insert  [TSS_WH_MatRecptMainBldg_Header] values (" & txtRecNo.Text & ",'" & dtpReceiptdt.Value & "', '" & txtMatIssueNo.Text & "', '" & dtpissdt.Value & "','" & txtMatReqNo.Text & "', '" & dtpMatReqDt.Value & "', '" & txtFeedback.Text & "', '" & username & "', ' " & curdate & "')"
                cmSQL = New SqlCommand(strsql2, cnSQL)


                If cmSQL.ExecuteNonQuery() = 0 Then
                    MsgBox("Material Receipt Header Detail section is not saved " & strsql2, MsgBoxStyle.Exclamation, "Error!")

                    Exit Sub
                End If

                'save detail table

                For i As Integer = 0 To DataGridViewReceiptsMfg.RowCount - 1
                    strsql2 = "insert [TSS_WH_MatRecptMainBldg_Detail] values (" & txtRecNo.Text & ",'" & Me.DataGridViewReceiptsMfg.Rows(i).Cells("Slno").Value & "'," & _
                          "'" & Me.DataGridViewReceiptsMfg.Rows(i).Cells("Part_Number").Value & "','" & Me.DataGridViewReceiptsMfg.Rows(i).Cells("UOM").Value & "'," & Me.DataGridViewReceiptsMfg.Rows(i).Cells("Req_Qty").Value & "," & Me.DataGridViewReceiptsMfg.Rows(i).Cells("Issued_Qty").Value & "," & _
                          "'" & Me.DataGridViewReceiptsMfg.Rows(i).Cells("ItemWise_Feedback").Value & "','" & username & "', ' " & curdate & "')"


                    '   strSQL = " SELECT Slno,Part_Number,Part_Desc,UOM, Qty as Req_Qty, Issued_Qty ,Remarks, '' AS 'ItemWise_Feedback'  FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueDetail] where MatIssue_no = '" & txtMatIssueNo.Text & "'"



                    cmSQL = New SqlCommand(strsql2, cnSQL)

                    If cmSQL.ExecuteNonQuery() = 0 Then
                        MsgBox("Material Receipt detail section is not saved " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                        Exit Sub
                    End If
                Next

            End If


            transmode = "Update"
            nogenerate()

            MsgBox("Receipts updated.", vbInformation)
            btnRecAccept.Enabled = False
        End If

    End Sub

    Private Function GetData(ByVal sql As String) As DataTable
        '  Dim con As SqlConnection = New SqlConnection(ConnectionString)

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmd As SqlCommand = New SqlCommand(sql, cnSQL)
        cmd.CommandType = CommandType.Text
        Dim sda As SqlDataAdapter = New SqlDataAdapter(cmd)
        Dim dt As DataTable = New DataTable
        sda.Fill(dt)
        Return dt
    End Function

    Private Sub txtMatIssueNo_DoubleClick(sender As Object, e As EventArgs) Handles txtMatIssueNo.DoubleClick

        ' GroupBoxDCList.Width = 673
        'GroupBoxDCList.Height = 215
        'GroupBoxDCList.Location = New System.Drawing.Point(390, 0)
        'GroupBoxDCList.Visible = True
        'GroupBoxDCList.Enabled = True


        datagridDC.Location = New System.Drawing.Point(411, 11)
        datagridDC.Width = 646
        datagridDC.Height = 125
        datagridDC.Visible = True
        datagridDC.Enabled = True

        LoadDcNotAck()

        'LoadDCDetails()

    End Sub

    Private Sub txtMatIssueNo_TextChanged(sender As Object, e As EventArgs) Handles txtMatIssueNo.TextChanged

    End Sub

    Private Sub LoadDcNotAck()


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet


        'not in TSS_WH_MaterialIssueDetail to be included - already issued material
        'ONLY HEADER NEED TO BE CALLED - DISTINCT MAT REQ NO

        strSQL = "SELECT a.MatIssue_no as DC_No, a.MatIssue_Date as DC_Dt, a.MatReq_no,b.MatReq_Date, b.Type_Dept,b.Sub_Div,b.Cell, a.Notes FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueHeader]a inner join " & _
                   "[FSPrograms].[dbo].[TSS_WH_MaterialRequestHeader]b ON  a.MatReq_no = b.MatReq_no   where a.MatIssue_no NOT IN ( SELECT [MatIssueDC_no] FROM [FSPrograms].[dbo].[TSS_WH_MatRecptMainBldg_Header] WITH (NOLOCK)) order by a.MatIssue_no Asc"

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

    Private Sub LoadDCDetails()

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        DataGridViewReceiptsMfg.Visible = True
        DataGridViewReceiptsMfg.Enabled = True



        strSQL = "SELECT Slno,Part_Number,Part_Desc,UOM, Qty as Req_Qty, SUM(Issued_Qty) as Issued_Qty ,max(Remarks) as Remarks, '' AS 'ItemWise_Feedback'  FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueDetail] where MatIssue_no = '" & txtMatIssueNo.Text & "' group by Slno,Part_Number,Part_Desc,UOM, Qty "
    

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
            'get data
        stockDAC.Fill(stockDC)

        DataGridViewReceiptsMfg.DataSource = stockDC.Tables(0)

        DataGridViewReceiptsMfg.Columns("Slno").ReadOnly = True
        DataGridViewReceiptsMfg.Columns("Slno").Width = 45

        DataGridViewReceiptsMfg.Columns("Part_Number").ReadOnly = True
        DataGridViewReceiptsMfg.Columns("Part_Number").Width = 145

        DataGridViewReceiptsMfg.Columns("Part_Desc").ReadOnly = True
        DataGridViewReceiptsMfg.Columns("Part_Desc").Width = 170

        DataGridViewReceiptsMfg.Columns("UOM").ReadOnly = True
        DataGridViewReceiptsMfg.Columns("UOM").Width = 45

        DataGridViewReceiptsMfg.Columns("Req_Qty").ReadOnly = True
        DataGridViewReceiptsMfg.Columns("Req_Qty").Width = 80
        DataGridViewReceiptsMfg.Columns("Req_Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewReceiptsMfg.Columns("Req_Qty").DefaultCellStyle.Format = "N2"

        DataGridViewReceiptsMfg.Columns("Issued_Qty").ReadOnly = True
        DataGridViewReceiptsMfg.Columns("Issued_Qty").Width = 80
        DataGridViewReceiptsMfg.Columns("Issued_Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewReceiptsMfg.Columns("Issued_Qty").DefaultCellStyle.Format = "N2"

        DataGridViewReceiptsMfg.Columns("Remarks").ReadOnly = True
        DataGridViewReceiptsMfg.Columns("Remarks").Width = 100

        DataGridViewReceiptsMfg.Columns("ItemWise_Feedback").ReadOnly = False
        DataGridViewReceiptsMfg.Columns("ItemWise_Feedback").Width = 350
        DataGridViewReceiptsMfg.Columns("ItemWise_Feedback").HeaderCell.Style.BackColor = Color.Gray


        cnSQL.Close()


    End Sub

    Private Sub datagridDC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub datagridDC_MouseClick(sender As Object, e As MouseEventArgs)

    End Sub

    Private Sub datagridDC_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles datagridDC.RowHeaderMouseClick

        ' Private Sub DataGridViewMatReqEdit_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewMatReqEdit.RowHeaderMouseClick
        'strSQL = "SELECT a.MatIssue_no as DC_No, a.MatIssue_Date as DC_Dt, a.MatReq_no,b.MatReq_Date, b.Type_Dept,b.Sub_Div,b.Cell, a.Notes FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueHeader]a inner join " & _
        '"[FSPrograms].[dbo].[TSS_WH_MaterialRequestHeader]b ON  a.MatReq_no = b.MatReq_no   where a.MatIssue_no NOT IN ( SELECT [MatIssueDC_no] FROM [FSPrograms].[dbo].[TSS_WH_MatRecptMainBldg_Header] WITH (NOLOCK))"


        txtMatIssueNo.Text = datagridDC.CurrentRow.Cells(0).Value
        txtMatIssueNo.Enabled = False

        dtpissdt.Value = datagridDC.CurrentRow.Cells(1).Value
        dtpissdt.Enabled = False

        txtMatReqNo.Text = datagridDC.CurrentRow.Cells(2).Value
        txtMatReqNo.Enabled = False

        dtpMatReqDt.Value = datagridDC.CurrentRow.Cells(3).Value
        dtpMatReqDt.Enabled = False

        txtrem.Text = datagridDC.CurrentRow.Cells(7).Value
        txtrem.Enabled = False

        datagridDC.Visible = False

        LoadDCDetails()



    End Sub

    Private Sub datagridDC_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles datagridDC.CellContentClick

    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click

        clearall()
        MsgBox("Click on DC Number", vbInformation)
        Exit Sub


    End Sub

    Private Sub clearall()
        txtMatIssueNo.Enabled = True
        txtMatIssueNo.Text = ""
        txtMatReqNo.Text = ""
        txtRecNo.Text = ""
        txtrem.Text = ""
        txtFeedback.Text = ""
        DataGridViewReceiptsMfg.Columns.Clear()
        btnRecAccept.Enabled = True
       
    End Sub
End Class