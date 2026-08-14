Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection

Public Class WHReturnToolReceipt

    ' Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"
    Private Sub WHReturnToolReceipt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DTPToolRetDt.Format = DateTimePickerFormat.Custom
        DTPToolRetDt.CustomFormat = "dd/MM/yyyy"
     
        Me.Width = 1300
        Me.Height = 840

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub txtMatIssNo_DoubleClick(sender As Object, e As EventArgs) Handles txtToolRetNo.DoubleClick
        PendingReturns()
    End Sub

    Private Sub txtMatIssNo_TextChanged(sender As Object, e As EventArgs) Handles txtToolRetNo.TextChanged

    End Sub

    Private Sub PendingReturns()
        datagridToolRetPending.VirtualMode = True
        datagridToolRetPending.BringToFront()
        ' datagridReqPending.Location.X.MaxValue = 574

        datagridToolRetPending.Width = 613
        datagridToolRetPending.Height = 190
        datagridToolRetPending.Enabled = True


        'RBToolNo.Checked = True
        'RadioButtonGroup.Checked = True
        'RadioButtonVendorYes.Checked = True


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet


        'not in TSS_WH_MaterialIssueDetail to be included - already issued material
        'ONLY HEADER NEED TO BE CALLED - DISTINCT MAT REQ NO

        strSQL = "SELECT distinct [MatReq_no],[MatReq_Date],[Type_Dept],[Sub_Div] ,[Cell] ,[MONumber], [Remarks],[User_Id] " & _
                 "FROM [FSPrograms].[dbo].[TSS_WH_ToolReturnPending_P] "

       


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
            'get data
        stockDAC.Fill(stockDC)

        datagridToolRetPending.DataSource = stockDC.Tables(0)


            '  TSS_WH_MatReqPendingAppDet_P()
    End Sub

    Private Sub datagridReqPending_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridToolRetPending.CellContentClick

    End Sub

    Private Sub datagridReqPending_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles datagridToolRetPending.RowHeaderMouseClick

        txtMatReqNo.Text = datagridToolRetPending.CurrentRow.Cells(0).Value
     
        ComboBoxdept.Text = datagridToolRetPending.CurrentRow.Cells(2).Value
        ComboBoxSD.Text = datagridToolRetPending.CurrentRow.Cells(3).Value
        ComboBoxCell.Text = datagridToolRetPending.CurrentRow.Cells(4).Value
        txtMO.Text = datagridToolRetPending.CurrentRow.Cells(5).Value
        txtRemarks.Text = datagridToolRetPending.CurrentRow.Cells(6).Value

      

        LoadMatReqDetails()


    End Sub

    Private Sub LoadMatReqDetails()

        DataGridViewToolRetDetail.Visible = True
        DataGridViewToolRetDetail.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        Dim stockDC As DataSet = New DataSet

        'issued should not come again

        strSQL = "SELECT [Slno] ,[Part_Number] AS ItemNumber ,[Part_Desc] as Description ,[Qty] ,[UsedPartY_N] AS 'Used Part',[Purpose],DetailRemarks " & _
                    "FROM [FSPrograms].[dbo].[TSS_WH_ToolReturnPending_P] where [MatReq_no] = " & txtMatReqNo.Text & " "

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
            'get data
        stockDAC.Fill(stockDC)


        DataGridViewToolRetDetail.DataSource = stockDC.Tables(0)

        DataGridViewToolRetDetail.Columns("Slno").ReadOnly = True
        DataGridViewToolRetDetail.Columns("Slno").Width = 45
        DataGridViewToolRetDetail.Columns("ItemNumber").ReadOnly = True
        DataGridViewToolRetDetail.Columns("ItemNumber").Width = 140
        DataGridViewToolRetDetail.Columns("Description").ReadOnly = True
        DataGridViewToolRetDetail.Columns("Description").Width = 160

        DataGridViewToolRetDetail.Columns("Qty").ReadOnly = True
        DataGridViewToolRetDetail.Columns("Qty").Width = 100
        DataGridViewToolRetDetail.Columns("Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewToolRetDetail.Columns("Qty").DefaultCellStyle.Format = "N2"

        DataGridViewToolRetDetail.Columns("Used Part").ReadOnly = True
        DataGridViewToolRetDetail.Columns("Used Part").Width = 100

        DataGridViewToolRetDetail.Columns("Purpose").ReadOnly = True
        DataGridViewToolRetDetail.Columns("Purpose").Width = 110

        DataGridViewToolRetDetail.Columns("DetailRemarks").ReadOnly = True
        DataGridViewToolRetDetail.Columns("DetailRemarks").Width = 200

        cnSQL.Close()
        datagridToolRetPending.Visible = False
        txtMatReqNo.Enabled = False
        ComboBoxdept.Enabled = False
        ComboBoxSD.Enabled = False
        ComboBoxCell.Enabled = False
        txtMO.Enabled = False
        txtRemarks.Enabled = False
        DTPToolRetDt.Enabled = False
            '  txtMatIssNo.Enabled = False



    End Sub

    Private Sub btnReqSave_Click(sender As Object, e As EventArgs) Handles btnToolRetSave.Click
        Dim checkdt As Date
        checkdt = Today

        Dim strsql2 As String
        Dim cmSQL As SqlCommand
        Dim msgb As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        msgb = MsgBox("Are you sure of saving ?", vbYesNo)

        If msgb = vbYes Then

            'generating regno

            transtype = "ToolRet"
            transmode = "Add"
            nogenerate()

            txtToolRetNo.Text = toolretno

            If txtToolRetNo.Text > 0 Then

                cnSQL.Open()

                curdate = System.DateTime.Now()

                'save header table
                'SELECT     TOP (200) id, ToolRet_no, ToolRet_Date, MatReq_no, Notes, User_Id, Datetime
                'FROM TSS_WH_ToolReturnHeader

                strsql2 = "insert  TSS_WH_ToolReturnHeader values (" & txtToolRetNo.Text & ",'" & DTPToolRetDt.Value & "', " & txtMatReqNo.Text & ",'" & txtHeaderNotes.Text & "', '" & username & "', ' " & curdate & "')"
                cmSQL = New SqlCommand(strsql2, cnSQL)

                If cmSQL.ExecuteNonQuery() = 0 Then
                    MsgBox("Tool Return Header Details are not saved " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                    '  txtRegNo.Text = 0
                    'Application.Exit()
                    Exit Sub
                End If


                'save detail table

                For i As Integer = 0 To DataGridViewToolRetDetail.RowCount - 1

                    'SELECT     TOP (200) id, ToolRet_no, ToolRet_Date, MatReq_no, Slno, Part_Number, Part_Desc, Qty, UsedPartY_N, Purpose, User_Id, DateTime
                    ' FROM TSS_WH_ToolReturnDetail
                    strsql2 = "insert TSS_WH_ToolReturnDetail values (" & txtToolRetNo.Text & ",'" & DTPToolRetDt.Value & "', " & txtMatReqNo.Text & ", " & Me.DataGridViewToolRetDetail.Rows(i).Cells("Slno").Value & "," & _
                              "'" & Me.DataGridViewToolRetDetail.Rows(i).Cells("ItemNumber").Value & "','" & Me.DataGridViewToolRetDetail.Rows(i).Cells("Description").Value & "', " & Me.DataGridViewToolRetDetail.Rows(i).Cells("Qty").Value & " , " & _
                              "'" & Me.DataGridViewToolRetDetail.Rows(i).Cells("Used Part").Value & "','" & Me.DataGridViewToolRetDetail.Rows(i).Cells("Purpose").Value & "','" & username & "', ' " & curdate & "')"

                    cmSQL = New SqlCommand(strsql2, cnSQL)

                    If cmSQL.ExecuteNonQuery() = 0 Then
                        MsgBox("Tool return detail section is not saved " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                        '  txtRegNo.Text = 0
                        'Application.Exit()
                        Exit Sub
                    End If
                Next

                transmode = "Update"
                nogenerate()

                MsgBox("Tool return data saved ", vbInformation)


            End If

        End If

       

    End Sub

    Private Sub DataGridViewToolRetDetail_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewToolRetDetail.CellContentClick

    End Sub
End Class