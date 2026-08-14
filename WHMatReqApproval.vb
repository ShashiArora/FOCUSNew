Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection
Public Class WHMatRequestApproval

    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        curdate = System.DateTime.Now()


    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnMyApproval.Click
        GroupBoxPendApr.Text = "My Approvals"
        btnApprove.Enabled = True
        btnReject.Enabled = True

        Approvals()

        'DataGridViewAprSummary.Columns.Clear()



        'Dim AlarmColumn1 As New DataGridViewCheckBoxColumn(False)

        'AlarmColumn1.Name = "Sel"
        'AlarmColumn1.HeaderText = "Select"
        'AlarmColumn1.ReadOnly = False


        'DataGridViewAprSummary.Columns.Add(AlarmColumn1)
        'DataGridViewAprSummary.ReadOnly = False



        'Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim strSQL As String


        'Dim stockDC As DataSet = New DataSet

        ''  strSQL = "SELECT  Enq_Reg_No, Project_Number, Project_MainStage, Project_SubStage, Project_Stage_Status as Status, Project_Stage_StartDate as StartDate, Project_Stage_EndDate as EndDate, Project_Stage_ActionBy as ActionBy, Project_Stage_Details as Details,Send_Info_To,Int_code FROM ENQ_Project_Masterlist_Update " & _
        '' " where  Enq_Reg_No = " & txtEnqRegNo.Text & " order by Int_code "

        'strSQL = " SELECT [MatReq_no] as 'Req_No',[MatReq_Date] as 'Req Date' ,[Type_Dept] as 'Dept',[Sub_Div],[Cell],[MONumber],[User_Name],[Remarks] FROM [FSPrograms].[dbo].[TSS_WH_MatReqPendingAppHeader_P] where [IstApr_UserId] = '" & username & "' "



        'Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        'Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        'stockDAC.SelectCommand = sqlCmd
        'cnSQL.Open()

        'stockDAC.TableMappings.Add("Table", "Enq")
        '    'get data
        'stockDAC.Fill(stockDC)
        'DataGridViewAprSummary.DataSource = stockDC.Tables(0)

        ''   DataGridViewAprSummary.Columns("Req_No").ReadOnly = True

        'DataGridViewAprSummary.Columns("Req_No").Width = 65
        'DataGridViewAprSummary.Columns("Req Date").Width = 85
        'DataGridViewAprSummary.Columns("Dept").Width = 85
        'DataGridViewAprSummary.Columns("Sub_Div").Width = 85
        'DataGridViewAprSummary.Columns("Cell").Width = 85
        'DataGridViewAprSummary.Columns("MONumber").Width = 85
        'DataGridViewAprSummary.Columns("User_Name").Width = 100
        'DataGridViewAprSummary.Columns("Remarks").Width = 100

        'cnSQL.Close()

    End Sub

    Private Sub BtnApproved_Click(sender As Object, e As EventArgs) Handles BtnApproved.Click
        GroupBoxPendApr.Text = "Completed List"
        Approvals()
    End Sub

    Private Sub DataGridViewAprSummary_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewAprSummary.CellContentClick

        Dim b As Integer
        'Dim custid As String
        '  b = DataGridViewAprSummary.CurrentCell.ColumnNumber()

        b = DataGridViewAprSummary.CurrentCell.ColumnIndex()

        If b = 0 Then





            Dim ColumnName1 As String = DataGridViewAprSummary.Columns(e.ColumnIndex).Name

            Dim strsql As String
            '        Dim cmSQL As SqlCommand
            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

            If CheckBoxItems.Checked = True Then



                Dim CellCheckBox1 As DataGridViewCheckBoxCell = _
                    CType(DataGridViewAprSummary.Rows(e.RowIndex).Cells(ColumnName1), DataGridViewCheckBoxCell)

                Dim CellCheckBoxState1 As String = CellCheckBox1.EditingCellFormattedValue.ToString

                'If CheckBoxItems.Checked = True Then 'And DataGridViewAprSummary.Rows(CellCheckBox1.RowIndex).Cells(0).Value = True Then


                If DataGridViewAprSummary.Rows(CellCheckBox1.RowIndex).Cells(0).Value = True Or DataGridViewAprSummary.Rows(CellCheckBox1.RowIndex).Cells(0).Value = False Then


                    If DataGridViewAprSummary.Rows(CellCheckBox1.RowIndex).Cells(0).Value = True Then

                        DataGridViewAprSummary.Rows(CellCheckBox1.RowIndex).Cells(0).Value = False

                    ElseIf DataGridViewAprSummary.Rows(CellCheckBox1.RowIndex).Cells(0).Value = False Then

                        DataGridViewAprSummary.Rows(CellCheckBox1.RowIndex).Cells(0).Value = True
                    End If

                End If

                '   If CheckBoxItems.Checked = True And DataGridViewAprSummary.Rows(CellCheckBox1.RowIndex).Cells(0).Value = True Then


                'clearing and LOADING DETAIL LIST
                DataGridViewApprovalDetail.Columns.Clear()

                Dim AlarmColumn2 As New DataGridViewCheckBoxColumn(False)

                AlarmColumn2.Name = "Rej"
                AlarmColumn2.HeaderText = "Reject"
                AlarmColumn2.ReadOnly = False


                DataGridViewApprovalDetail.Columns.Add(AlarmColumn2)
                DataGridViewApprovalDetail.ReadOnly = False
                DataGridViewApprovalDetail.Columns("Rej").Width = 65

                Dim stockDC As DataSet = New DataSet

                If GroupBoxPendApr.Text = "Completed List" Then

                    strsql = "SELECT MatReq_no as 'Req_No',Slno, Issue_Ret AS Iss_Ret, Part_Number,Part_Desc,Qty,DetailRemarks as Remarks,[RejLines],case when [RejLines] > 0 THEN 'NO' ELSE [1st_AppStatus] END as 'App_Status' ,[Reason_Rej] as Reason_Rej FROM [FSPrograms].[dbo].[TSS_WH_MatReqAppCompDetail_P]  where MatReq_no  = " & DataGridViewAprSummary.CurrentRow.Cells(1).Value & " order by Slno Asc"


                Else

                    '     If ColumnName1 = "Sel" Then
                    strsql = "SELECT  dbo.TSS_WH_MaterialRequestHeader.MatReq_no as 'Req_No', dbo.TSS_WH_MaterialRequestDetail.Slno, dbo.TSS_WH_MaterialRequestDetail.Issue_Ret AS Iss_Ret,dbo.TSS_WH_MaterialRequestDetail.Part_Number, dbo.TSS_WH_MaterialRequestDetail.Part_Desc, " & _
                                "dbo.TSS_WH_MaterialRequestDetail.Qty, dbo.TSS_WH_MaterialRequestDetail.Remarks,'' AS Reason_Rej, '' as App_Status  FROM  dbo.TSS_WH_MaterialRequestHeader WITH (NOLOCK) INNER JOIN " & _
                                "dbo.TSS_WH_MaterialRequestDetail WITH (NOLOCK) ON dbo.TSS_WH_MaterialRequestHeader.MatReq_no = dbo.TSS_WH_MaterialRequestDetail.MatReq_no  and " & _
                                "dbo.TSS_WH_MaterialRequestHeader.MatReq_no  = " & DataGridViewAprSummary.CurrentRow.Cells(1).Value & " order by dbo.TSS_WH_MaterialRequestDetail.Slno Asc"
                End If

                Dim sqlCmd As SqlCommand = New SqlCommand(strsql, cnSQL)
                Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

                stockDAC.SelectCommand = sqlCmd
                cnSQL.Open()

                stockDAC.TableMappings.Add("Table", "Enq")

                stockDAC.Fill(stockDC)

                DataGridViewApprovalDetail.DataSource = stockDC.Tables(0)
                DataGridViewApprovalDetail.Columns("Req_No").Width = 80
                DataGridViewApprovalDetail.Columns("Req_No").ReadOnly = True

                DataGridViewApprovalDetail.Columns("Slno").Width = 55
                DataGridViewApprovalDetail.Columns("Slno").ReadOnly = True
                DataGridViewApprovalDetail.Columns("Slno").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                DataGridViewApprovalDetail.Columns("Iss_Ret").Width = 55
                DataGridViewApprovalDetail.Columns("Iss_Ret").ReadOnly = True
                DataGridViewApprovalDetail.Columns("Iss_Ret").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                DataGridViewApprovalDetail.Columns("Part_Number").Width = 160
                DataGridViewApprovalDetail.Columns("Part_Number").ReadOnly = True

                DataGridViewApprovalDetail.Columns("Part_Desc").Width = 150
                DataGridViewApprovalDetail.Columns("Part_Desc").ReadOnly = True

                DataGridViewApprovalDetail.Columns("Qty").Width = 60
                DataGridViewApprovalDetail.Columns("Qty").ReadOnly = True
                DataGridViewApprovalDetail.Columns("Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridViewApprovalDetail.Columns("Qty").DefaultCellStyle.Format = "N2"


                DataGridViewApprovalDetail.Columns("Remarks").Width = 150
                DataGridViewApprovalDetail.Columns("Remarks").ReadOnly = True


                DataGridViewApprovalDetail.Columns("App_Status").Width = 100
                DataGridViewApprovalDetail.Columns("App_Status").ReadOnly = True

                If GroupBoxPendApr.Text = "Completed List" Then

                    DataGridViewApprovalDetail.Columns("RejLines").Width = 0
                    DataGridViewApprovalDetail.Columns("RejLines").ReadOnly = True
                    DataGridViewApprovalDetail.Columns("RejLines").Visible = False

                    DataGridViewApprovalDetail.Columns("Reason_Rej").Width = 150
                    DataGridViewApprovalDetail.Columns("Reason_Rej").ReadOnly = True

                Else
                    DataGridViewApprovalDetail.Columns("Reason_Rej").Width = 150
                    DataGridViewApprovalDetail.Columns("Reason_Rej").ReadOnly = False
                    DataGridViewApprovalDetail.Columns("Reason_Rej").HeaderCell.Style.BackColor = Color.Gray

                End If


                cnSQL.Close()


                'If GroupBoxPendApr.Text = "Completed List" Then
                'If Val(DataGridViewApprovalDetail.Rows(CellCheckBox1.RowIndex).Cells(6).Value) > 0 Then
                '    DataGridViewApprovalDetail.Rows(CellCheckBox1.RowIndex).Cells(0).Value = True
                'Else
                '    DataGridViewApprovalDetail.Rows(CellCheckBox1.RowIndex).Cells(0).Value = False
                'End If
                'END OF LOADING DETAIL LIST
                'End If
            End If

        End If




    End Sub

    Private Sub BtnItems_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckBoxItems_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxItems.CheckedChanged
        DataGridViewApprovalDetail.Columns.Clear()
    End Sub

    Private Sub btnApprove_Click(sender As Object, e As EventArgs) Handles btnApprove.Click

        Dim rec As Integer
        rec = DataGridViewAprSummary.Rows.Count

        If rec > 0 Then

            Dim b As Integer

            b = DataGridViewAprSummary.CurrentCell.ColumnIndex

            If b = 0 Then

                Dim msgb As String
                msgb = MsgBox("Are you sure of approving ?", vbYesNo)

                If msgb = vbYes Then

                    For i As Integer = 0 To DataGridViewApprovalDetail.RowCount - 1

                        Dim Checked As Boolean = CType(Me.DataGridViewApprovalDetail.Rows(i).Cells("Rej").Value, Boolean)
                        If Checked Then

                            If Len(Me.DataGridViewApprovalDetail.Rows(i).Cells("Reason_Rej").Value) < 3 Then

                                MsgBox("Reason for Rejection to be entered", vbInformation)
                                Exit Sub

                            End If
                        End If

                    Next


                    Dim errcount As Integer
                    errcount = 0
                    For i As Integer = 0 To DataGridViewApprovalDetail.RowCount - 1

                        If Len(Me.DataGridViewApprovalDetail.Rows(i).Cells("Reason_Rej").Value) > 100 Then
                            errcount = errcount + 1
                        End If

                    Next

                    If errcount > 0 Then

                        MsgBox("Reason for rejection should not exceed 100 chrs", vbInformation)
                        Exit Sub

                    End If


                    'update statement for Material request number wise

                    Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
                    Dim cmSQL1 As SqlCommand
                    ' Dim drSQL1 As SqlDataReader
                    Dim strSQL1 As String
                    'Material Receipt ******
                    cnSQL1.Open()
                    For i As Integer = 0 To DataGridViewAprSummary.RowCount - 1

                        '  Dim Checked As Boolean = CType(DataGridViewAprSummary.CurrentCell.Value, Boolean)

                        Dim Checked As Boolean = CType(Me.DataGridViewAprSummary.Rows(i).Cells("Sel").Value, Boolean)
                        If Checked Then

                            strSQL1 = "Update [FSPrograms].[dbo].[TSS_WH_Approvals] set [1st_AppStatus] = 'YES', [IstApr_UserId] = '" & username & "',[IstApr_DateTime] = '" & curdate & "'  where [Trans_Type] = 'MatReq' and  [Trans_No] = " & Me.DataGridViewAprSummary.Rows(i).Cells("Req_No").Value & ""

                            cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
                            If cmSQL1.ExecuteNonQuery() = 0 Then
                                MsgBox("Approval status not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                                Exit Sub
                            End If
                        End If
                    Next
                    cnSQL1.Close()

                    'end of update statement for Material request number wise

                    'saving rejected items if any


                    cnSQL1.Open()
                    '     Dim rejcount As Integer
                    For i As Integer = 0 To DataGridViewApprovalDetail.RowCount - 1

                        Dim Checked As Boolean = CType(Me.DataGridViewApprovalDetail.Rows(i).Cells("Rej").Value, Boolean)
                        If Checked Then



                            '  Dim Checked As Boolean = CType(Me.DataGridViewApprovalDetail.Rows(i).Cells("Rej").Value, Boolean)
                            ' If Checked Then

                            'If Len(Me.DataGridViewApprovalDetail.Rows(i).Cells("Reason_Rej").Value) < 3 Then

                            'MsgBox("Reson for Rejection to be entered", vbInformation)
                            'Exit Sub

                            'End If



                            '      Dim Checked As Boolean = CType(DataGridViewApprovalDetail.CurrentRow.Cells(0).Value, Boolean)
                            '     If Checked Then
                            'rejcount = rejcount + 1
                            strSQL1 = "insert TSS_WH_MaterialReqDetail_Rej values (" & Me.DataGridViewApprovalDetail.Rows(i).Cells("Req_no").Value & "," & Me.DataGridViewApprovalDetail.Rows(i).Cells("Slno").Value & ", '" & Me.DataGridViewApprovalDetail.Rows(i).Cells("Part_Number").Value & "', '" & Me.DataGridViewApprovalDetail.Rows(i).Cells("Part_Desc").Value & "', " & _
                                " " & Me.DataGridViewApprovalDetail.Rows(i).Cells("Qty").Value & ", '" & Me.DataGridViewApprovalDetail.Rows(i).Cells("Reason_Rej").Value & "', '" & username & "' , '" & curdate & "')"


                            cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
                            If cmSQL1.ExecuteNonQuery() = 0 Then
                                MsgBox("Rejection Status not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                                Exit Sub
                            End If
                        End If
                    Next

                    '    If rejcount >= 1 Then

                    'MsgBox("Material Request Approved and rejection Status upated for Rejected Lines", vbInformation)
                    'Exit Sub

                    'Else
                    MsgBox("Material Request Approved", vbInformation)

                    btnApprove.Enabled = False
                    'End If
                    cnSQL1.Close()

                End If
            Else
                MsgBox("Click on 'Sel' Check box", vbInformation)
                Exit Sub


            End If

        End If

    End Sub

    Private Sub btnReject_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        Dim rec As Integer
        rec = DataGridViewAprSummary.Rows.Count

        If rec > 0 Then


            Dim b As Integer

            b = DataGridViewAprSummary.CurrentCell.ColumnIndex

            If b = 0 Then

                Dim msgb As String
                msgb = MsgBox("Are you sure of Rejecting ?", vbYesNo)

                If msgb = vbYes Then


                    Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
                    Dim cmSQL1 As SqlCommand
                    ' Dim drSQL1 As SqlDataReader
                    Dim strSQL1 As String
                    'Material Receipt ******

                    ' DataGridViewAprSummary.Columns("Reason").ReadOnly = False


                    'reason for rejection length checking

                    For i As Integer = 0 To DataGridViewAprSummary.RowCount - 1

                        Dim Checked As Boolean = CType(DataGridViewAprSummary.CurrentCell.Value, Boolean)
                        If Checked Then


                            If Len(Me.DataGridViewAprSummary.Rows(i).Cells("Reason_Rej").Value) < 3 Then
                                MsgBox("Reason for rejection to be entered", vbInformation)
                                Exit Sub

                            End If
                        End If


                    Next





                    Dim errcount As Integer
                    errcount = 0
                    For i As Integer = 0 To DataGridViewAprSummary.RowCount - 1

                        If Len(Me.DataGridViewAprSummary.Rows(i).Cells("Reason_Rej").Value) > 150 Then
                            errcount = errcount + 1
                        End If

                    Next

                    If errcount > 0 Then

                        MsgBox("Reason for rejection should not exceed 150 chrs", vbInformation)
                        Exit Sub

                    End If



                    For i As Integer = 0 To DataGridViewAprSummary.RowCount - 1

                        '    Dim Checked As Boolean = CType(DataGridViewAprSummary.CurrentCell.Value, Boolean)
                        '   If Checked Then


                        'If Len(Me.DataGridViewAprSummary.Rows(i).Cells("Reason_Rej").Value) < 3 Then
                        'MsgBox("Reason for rejection to be entered", vbInformation)
                        'Exit Sub

                        'End If


                        strSQL1 = "Update [FSPrograms].[dbo].[TSS_WH_Approvals] set [1st_AppStatus] = 'NO', [IstApr_UserId] = '" & username & "',[IstApr_DateTime] = '" & curdate & "',[1st_ Reason] = '" & Me.DataGridViewAprSummary.Rows(i).Cells("Reason_Rej").Value & "'  where [Trans_Type] = 'MatReq' and  [Trans_No] = " & Me.DataGridViewAprSummary.Rows(i).Cells("Req_No").Value & ""
                        cnSQL1.Open()
                        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
                        If cmSQL1.ExecuteNonQuery() = 0 Then
                            MsgBox("Rejection status not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                            Exit Sub
                        End If
                        'End If
                    Next
                    MsgBox("Rejection Status updated", vbInformation)
                    btnReject.Enabled = False
                    cnSQL1.Close()

                End If


            Else

                MsgBox("Click on Selection check box", vbInformation)
                Exit Sub
            End If

        End If


    End Sub

    Private Sub DataGridViewApprovalDetail_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewApprovalDetail.CellContentClick

        Dim b As Integer
        'Dim custid As String
        '  b = DataGridViewAprSummary.CurrentCell.ColumnNumber()

        b = DataGridViewApprovalDetail.CurrentCell.ColumnIndex()

        If b = 0 Then

            Dim ColumnName2 As String = DataGridViewApprovalDetail.Columns(e.ColumnIndex).Name

            Dim CellCheckBox2 As DataGridViewCheckBoxCell = _
                CType(DataGridViewApprovalDetail.Rows(e.RowIndex).Cells(ColumnName2), DataGridViewCheckBoxCell)

            Dim CellCheckBoxState1 As String = CellCheckBox2.EditingCellFormattedValue.ToString


            If DataGridViewApprovalDetail.Rows(CellCheckBox2.RowIndex).Cells(0).Value = True Then

                DataGridViewApprovalDetail.Rows(CellCheckBox2.RowIndex).Cells(0).Value = False

            ElseIf DataGridViewApprovalDetail.Rows(CellCheckBox2.RowIndex).Cells(0).Value = False Then

                DataGridViewApprovalDetail.Rows(CellCheckBox2.RowIndex).Cells(0).Value = True
            End If

        End If


    End Sub

    Private Sub BtnItemReject_Click(sender As Object, e As EventArgs) Handles BtnItemReject.Click

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        ' Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String
        'Material Receipt ******
        cnSQL1.Open()

        For i As Integer = 0 To DataGridViewApprovalDetail.RowCount - 1


            '   Dim Checked As Boolean = CType(DataGridViewApprovalDetail.CurrentCell.Value, Boolean)

            'NEED TO CHANGED - IT IS NOT SAVING PROPERLY - 9TH MAY 2019

            Dim Checked As Boolean = CType(DataGridViewApprovalDetail.CurrentRow.Cells(0).Value, Boolean)
            'Table rejected lines - Insert to be done

            If Checked Then
                strSQL1 = "insert TSS_WH_MaterialReqDetail_Rej values (" & Me.DataGridViewApprovalDetail.Rows(i).Cells("MatReq_no").Value & "," & Me.DataGridViewApprovalDetail.Rows(i).Cells("Slno").Value & ", '" & Me.DataGridViewApprovalDetail.Rows(i).Cells("Part_Number").Value & "', '" & Me.DataGridViewApprovalDetail.Rows(i).Cells("Part_Desc").Value & "', " & _
                    " " & Me.DataGridViewApprovalDetail.Rows(i).Cells("Qty").Value & ", '" & Me.DataGridViewApprovalDetail.Rows(i).Cells("Reason_Rej").Value & "', '" & username & "' , '" & curdate & "')"


                'SELECT     TOP (200) id, MatReq_no, MatReq_Date, Slno, Part_Number, Part_Desc, Qty, Reason_Rej, User_Id, Datetime
                'FROM TSS_WH_MaterialReqDetail_Rej


                cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
                If cmSQL1.ExecuteNonQuery() = 0 Then
                    MsgBox("Rejection Status not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                    Exit Sub
                End If
            End If
        Next
        MsgBox("Rejection Status upated", vbInformation)
        cnSQL1.Close()



    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        GroupBoxPendApr.Text = "Others Approvals"
        Approvals()
    End Sub


    Private Sub Approvals()


        DataGridViewAprSummary.Columns.Clear()
        DataGridViewApprovalDetail.Columns.Clear()

        Dim AlarmColumn1 As New DataGridViewCheckBoxColumn(False)

        AlarmColumn1.Name = "Sel"
        AlarmColumn1.HeaderText = "Select"
        AlarmColumn1.ReadOnly = False


        DataGridViewAprSummary.Columns.Add(AlarmColumn1)
        DataGridViewAprSummary.ReadOnly = False
        DataGridViewAprSummary.Columns("Sel").Width = 65


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        If GroupBoxPendApr.Text = "My Approvals" Then

            strSQL = " SELECT [MatReq_no] as 'Req_No',[MatReq_Date] as 'Req Date' ,[Type_Dept] as 'Dept',[Sub_Div],[Cell],[MONumber],[User_Name],[Remarks],''as Reason_Rej FROM [FSPrograms].[dbo].[TSS_WH_MatReqPendingAppHeader_P] where [IstApr_UserId] = '" & username & "' order by [MatReq_no] Asc "

        ElseIf GroupBoxPendApr.Text = "Others Approvals" Then

            'strSQL = " SELECT [MatReq_no] as 'Req_No',[MatReq_Date] as 'Req Date' ,[Type_Dept] as 'Dept',[Sub_Div],[Cell],[MONumber],[User_Name],[Remarks], IstApr_EmpName as 'Ist Approver','' as Reason_Rej FROM [FSPrograms].[dbo].[TSS_WH_MatReqPendingAppHeader_P] where [IstApr_UserId] <> '" & username & "' "
            strSQL = "SELECT [MatReq_no] as 'Req_No',[MatReq_Date] as 'Req Date' ,[Type_Dept] as 'Dept',[Sub_Div],[Cell],[MONumber],[User_Name],[Remarks], IstApr_EmpName as 'Ist Approver','' as Reason_Rej FROM [FSPrograms].[dbo].[TSS_WH_MatReqPendingAppHeader_P] where ([2ndApr_UserId] = '" & username & "' or [3rdApr_UserId] = '" & username & "') order by [MatReq_no] Asc"

        ElseIf GroupBoxPendApr.Text = "Completed List" Then
            strSQL = " SELECT  [MatReq_no] as  'Req_No' ,[MatReq_Date] as 'Req Date',[Type_Dept] as 'Dept' ,Sub_Div,Cell, User_Id as 'Req By',MONumber ,Remarks ,[IstApr_UserId] as Approver ,[IstApr_DateTime] 'Approved on' ,[1st_AppStatus] 'App_Status' ,[ApprovedBy]  FROM [FSPrograms].[dbo].[TSS_WH_MatReqAppComp_P] where [IstApr_UserId] = '" & username & "' order by [MatReq_no] Asc"

        End If

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)
        DataGridViewAprSummary.DataSource = stockDC.Tables(0)

        '   DataGridViewAprSummary.Columns("Req_No").ReadOnly = True
        If GroupBoxPendApr.Text = "My Approvals" Or GroupBoxPendApr.Text = "Others Approvals" Then
            DataGridViewAprSummary.Columns("Req_No").Width = 65
            DataGridViewAprSummary.Columns("Req_No").ReadOnly = True
            DataGridViewAprSummary.Columns("Req Date").Width = 85
            DataGridViewAprSummary.Columns("Req Date").ReadOnly = True
            DataGridViewAprSummary.Columns("Dept").Width = 85
            DataGridViewAprSummary.Columns("Dept").ReadOnly = True
            DataGridViewAprSummary.Columns("Sub_Div").Width = 85
            DataGridViewAprSummary.Columns("Sub_Div").ReadOnly = True
            DataGridViewAprSummary.Columns("Cell").Width = 85
            DataGridViewAprSummary.Columns("Cell").ReadOnly = True
            DataGridViewAprSummary.Columns("MONumber").Width = 85
            DataGridViewAprSummary.Columns("MONumber").ReadOnly = True
            DataGridViewAprSummary.Columns("User_Name").Width = 100
            DataGridViewAprSummary.Columns("User_Name").ReadOnly = True
            DataGridViewAprSummary.Columns("Remarks").Width = 150
            DataGridViewAprSummary.Columns("Remarks").ReadOnly = True
            If GroupBoxPendApr.Text = "Others Approvals" Then
                DataGridViewAprSummary.Columns("Ist Approver").Width = 100
                DataGridViewAprSummary.Columns("Ist Approver").ReadOnly = True
            End If

            DataGridViewAprSummary.Columns("Reason_Rej").Width = 190
            DataGridViewAprSummary.Columns("Reason_Rej").HeaderCell.Style.BackColor = Color.Gray
            'DataGridViewAprSummary.Columns("Reason_Rej").ReadOnly = False
        ElseIf GroupBoxPendApr.Text = "Completed List" Then

            DataGridViewAprSummary.Columns("Req_No").Width = 65
            DataGridViewAprSummary.Columns("Req_No").ReadOnly = True
            DataGridViewAprSummary.Columns("Req Date").Width = 85
            DataGridViewAprSummary.Columns("Req Date").ReadOnly = True
            DataGridViewAprSummary.Columns("Dept").Width = 85
            DataGridViewAprSummary.Columns("Dept").ReadOnly = True
            DataGridViewAprSummary.Columns("Sub_Div").Width = 85
            DataGridViewAprSummary.Columns("Sub_Div").ReadOnly = True
            DataGridViewAprSummary.Columns("Cell").Width = 85
            DataGridViewAprSummary.Columns("Cell").ReadOnly = True
            DataGridViewAprSummary.Columns("Req By").Width = 85
            DataGridViewAprSummary.Columns("Req By").ReadOnly = True

            DataGridViewAprSummary.Columns("MONumber").Width = 85
            DataGridViewAprSummary.Columns("MONumber").ReadOnly = True
            DataGridViewAprSummary.Columns("Remarks").Width = 100
            DataGridViewAprSummary.Columns("Remarks").ReadOnly = True

            DataGridViewAprSummary.Columns("Approver").Width = 85
            DataGridViewAprSummary.Columns("Approver").ReadOnly = True

            DataGridViewAprSummary.Columns("Approved on").Width = 100
            DataGridViewAprSummary.Columns("Approved on").ReadOnly = True

            DataGridViewAprSummary.Columns("App_Status").Width = 100
            DataGridViewAprSummary.Columns("APP_Status").ReadOnly = True

            DataGridViewAprSummary.Columns("ApprovedBy").Width = 100
            DataGridViewAprSummary.Columns("ApprovedBy").ReadOnly = True

            btnApprove.Enabled = False
            btnReject.Enabled = False

        End If

        cnSQL.Close()

    End Sub


    Private Sub GroupBoxPendApr_Enter(sender As Object, e As EventArgs) Handles GroupBoxPendApr.Enter

    End Sub
End Class