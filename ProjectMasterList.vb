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


Public Class ProjectMasterList
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Dim mode As String
    Dim quote As String
    Dim quotedate As Date
    Dim checkstatus As String
    Dim checklost As String

    Dim seq As String







    Private Sub ProjectMasterList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'GroupBoxEnqDetails1.Location = New Point(6, 7)
        'GroupBoxenqDetails.Visible = True
        'GroupBoxenqDetails.Width = 1129
        'GroupBoxenqDetails.Height = 189

        listloadMainStages()
        LOADCOMBOBOXAE()

        ' Me.ReportViewer1.RefreshReport()
    End Sub
    Private Sub LOADCOMBOBOXAE()
        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand

        strSql = "SELECT Int_code,AE_NAME FROM ENQ_Project_Engineers " & _
                 "WHERE Status like 'A%' ORDER BY AE_NAME"
        cmSQL = New SqlCommand(strSql, sqlCon)
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource As SqlDataAdapter = New SqlDataAdapter
        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eSource")
        With ComboBoxActionBy
            .DataSource = source.Tables("eSource")
            .DisplayMember = "AE_NAME"
            .ValueMember = "Int_code"
            .SelectedIndex = 0
        End With

    End Sub
    Private Sub listloadMainStages()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        CheckedListBoxMainStatus.Items.Clear()

        cnSQL1.Open()

        strSQL1 = "SELECT Project_MainStatus,Project_MainCode FROM ENQ_Project_Progress_MainStatus " & _
                 "WHERE  Status = 'A' order by Int_code "
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        Dim ColumnValue As String = Nothing
        Do While drSQL1.Read()

            ColumnValue = (drSQL1.GetValue(0)).ToString
            CheckedListBoxMainStatus.Items.Add(ColumnValue)
            CheckedListBoxSubStatus.ValueMember = "Project_MainCode"

        Loop


    End Sub

    Private Sub listloadSubStages()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String


        CheckedListBoxSubStatus.Items.Clear()

        cnSQL1.Open()

        ' SELECT     Project_MainStatus, Project_SubCode, Project_MainCode, Project_SubStatus, '' AS a
        'FROM(TSS_Enq_Project_Main_SubStages)


        strSQL1 = "SELECT Project_SubStatus,Project_SubCode FROM TSS_Enq_Project_Main_SubStages " & _
                 "WHERE   Project_MainStatus = '" & txtCurMainStage.Text & "' order by Project_SubCode "

        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        Dim ColumnValue As String = Nothing


        Do While drSQL1.Read()

            ColumnValue = (drSQL1.GetValue(0)).ToString
            CheckedListBoxSubStatus.Items.Add(ColumnValue)
            CheckedListBoxSubStatus.ValueMember = "Project_SubCode"

        Loop


    End Sub

    Private Sub DataGridQty_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ButtonRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRefresh.Click
        DataGridViewProjectMasterList.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String
        'Dim cmSQL As SqlCommand
        'Dim drSQL As SqlDataReader

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL1 As String
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader


        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL2 As String
        Dim cmSQL2 As SqlCommand
        Dim drSQL2 As SqlDataReader

        Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL3 As String
        Dim cmSQL3 As SqlCommand
        Dim drSQL3 As SqlDataReader


        mode = "ADD"

        Dim stockDC As DataSet = New DataSet
        strSQL = "SELECT  a.Enq_Reg_No, a.Project_Number, a.Project_RegDate, a.Project_Alotted,  a.CustomerName, a.CSR, a.Project_Type, a.EAR_Number, a.Project_Name, a.Application, a.Project_Status,a.QtnNumber,a.Class3, a.Class1,a.Enq_Int_code,a.Cust_IntCode,a.CustomerID FROM TSS_Project_MasterList a "
        'temp commented nested query on 9-apr-2017
        ' & _
        '    "wherea.Enq_Reg_No not in (select b.Enq_Reg_No from TSS_Enq_Project_Not_Rej_Lost b where a.Enq_Reg_No = b.Enq_Reg_No) ORDER BY a.Enq_Reg_No"


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter
        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewProjectMasterList.DataSource = stockDC.Tables(0)
        'cnSQL.Close()

        txttotal.Text = DataGridViewProjectMasterList.Rows.Count

        'colouring

        For i As Integer = 0 To DataGridViewProjectMasterList.Rows.Count - 1
            If DataGridViewProjectMasterList.Rows(i).Cells("Project_Status").Value = "PriceRecd" Then
                'DataGridViewProjectMasterList.Rows(i).Cells("Project_Status").Style.ForeColor = Color.Red
                DataGridViewProjectMasterList.Rows(i).Cells("Enq_Reg_No").Style.BackColor = Color.Purple
                'DataGridViewProjectMasterList.Rows(i).DefaultCellStyle.BackColor = Color.Blue
            End If
        Next

        'assigning project totals

        strSQL1 = "SELECT count(b.Enq_Reg_No) as tot FROM TSS_ENQ_Project_ProgressPerc b where b.Project_MainStage = 'Proto Order' and b.Project_Stage_Status ='Completed' and b.Enq_Reg_No not in " & _
        "(select a.Enq_Reg_No from TSS_ENQ_Project_ProgressPerc a where a.Project_MainStage = 'Production Order' and a.Enq_Reg_No = b.Enq_Reg_No)"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then


            If IsDBNull(drSQL1.Item(0)) Then
                txtproto.Text = 0
            Else

                txtproto.Text = drSQL1.Item(0)

            End If
        End If


        strSQL2 = "SELECT count(Enq_Reg_No) as tot FROM TSS_ENQ_Project_ProgressPerc where Project_MainStage = 'Production Order' and Project_Stage_Status ='Completed'"
        cnSQL2.Open()
        cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
        drSQL2 = cmSQL2.ExecuteReader()

        If drSQL2.Read() Then


            If IsDBNull(drSQL2.Item(0)) Then
                txtprod.Text = 0
            Else

                txtprod.Text = drSQL2.Item(0)

            End If
        Else
            txtprod.Text = 0
        End If

        strSQL3 = "SELECT count(Enq_Reg_No) as tot FROM TSS_ENQ_Project_ProgressPerc where Project_MainStage = 'Lost' and Project_Stage_Status ='Completed'"
        cnSQL3.Open()
        cmSQL3 = New SqlCommand(strSQL3, cnSQL3)
        drSQL3 = cmSQL3.ExecuteReader()

        If drSQL3.Read() Then


            If IsDBNull(drSQL3.Item(0)) Then
                txtlost.Text = 0
            Else

                txtlost.Text = drSQL3.Item(0)

            End If
        Else
            txtlost.Text = 0
        End If

        txtpend.Text = Val(txttotal.Text) - (Val(txtproto.Text) + Val(txtprod.Text) + Val(txtlost.Text))

        cnSQL.Close()
        cnSQL1.Close()
        cnSQL2.Close()
        cnSQL3.Close()

    End Sub

    Private Sub CheckedListBoxMainStatus_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles CheckedListBoxMainStatus.ItemCheck

        ' txtremarks.Text = CheckedListBoxSubStatus.SelectedValue.ToString
        ' Dim itemChecked As Object
        'Dim a As String
        'txtremarks.Text = CheckedListBoxMainStatus.SelectedItems.ToString
        ''txtremarks.Text = CheckedListBoxMainStatus.GetItemCheckState(CheckedListBoxMainStatus.Items.IndexOf(itemChecked)).ToString()
        'txtremarks.Text = CheckedListBoxMainStatus.GetItemCheckState(CheckedListBoxMainStatus.Items.ToString())

        'txtremarks.Text = CheckedListBoxMainStatus.GetItemCheckState(CheckedListBoxMainStatus.Items(0).ToString())

        'a = CheckedListBoxMainStatus.GetItemCheckState(CheckedListBoxMainStatus.Items(0).ToString())
        'a = CheckedListBoxMainStatus.GetItemCheckState(CheckedListBoxMainStatus.Items(0))

        'Dim indexChecked As Integer

        'For Each indexChecked In CheckedListBoxMainStatus.CheckedIndices

        'txtremarks.Text = (indexChecked.ToString() + ",")
        'txtremarks.Text = CheckedListBoxMainStatus.CheckedIndices.Item(1)


        ' MsgBox(txtremarks.Text)

        'Next


        'Dim i As Integer
        'i = CheckedListBoxMainStatus.CheckedItems.Count
        ''Dim a As Integer
        ''Dim cert As String



        ''Do While a < i
        'If i > 0 Then
        '    'Do While a = i
        '    'cert = ""
        '    txtremarks.Text = CheckedListBoxMainStatus.CheckedItems.Item(i)
        '    ' a = a + 1

        '    'Loop
        'End If


        'listloadSubStages()

    End Sub

    Private Sub CheckedListBoxMainStatus_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CheckedListBoxMainStatus.MouseClick
        ' MsgBox("mouseclick event")

    End Sub

    Private Sub CheckedListBoxMainStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBoxMainStatus.SelectedIndexChanged
        'txtremarks.Text = CheckedListBoxMainStatus.SelectedValue.ToString()
    End Sub

    Private Sub CheckedListBoxMainStatus_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBoxMainStatus.SelectedValueChanged
        ' txtremarks.Text = CheckedListBoxMainStatus.SelectedValue.ToString()

    End Sub


    Private Sub CheckedListBoxMainStatus_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBoxMainStatus.Validated

    End Sub

    Private Sub CheckedListBoxSubStatus_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles CheckedListBoxSubStatus.ItemCheck


    End Sub

    Private Sub CheckedListBoxSubStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBoxSubStatus.SelectedIndexChanged

    End Sub

    Private Sub btnsub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsub.Click


        CheckedListBoxSubStatus.Items.Clear()

        Dim i As Integer
        Dim m As Integer

        i = CheckedListBoxMainStatus.CheckedItems.Count

        If i > 0 Then

            CheckedListBoxSubStatus.Visible = True

        End If


        If i > 0 Then
            'Do While a = i
            txtCurMainStage.Text = CheckedListBoxMainStatus.CheckedItems.Item(i - 1)
            'txtCurMainStage.Text = CheckedListBoxMainStatus.CheckedIndices.Item(i - 1)

        Else
            MsgBox("Select main stage  first")
        End If

        listloadSubStages()

        m = CheckedListBoxSubStatus.Items.Count
        If m > 0 Then
        Else

            txtcursubstage.Text = ""

        End If




    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub GroupBoxSelect_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxSelect.Enter

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Static TikTok As Integer

        TikTok = TikTok + 1

        ProgressBarProjectProgress.Value = TikTok

        If ProgressBarProjectProgress.Value = ProgressBarProjectProgress.Maximum Then
            Timer1.Enabled = False 'deactivate Timer1

            TikTok = 0 'set static value back to 0
        End If


    End Sub

    Private Sub progress()


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String
        Dim a As Integer


        strSQL1 = "Select SUM(Final)FROM TSS_ENQ_Project_ProgressPerc where Enq_Reg_No = " & txtEnqRegNo.Text & ""

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then


            If IsDBNull(drSQL1.Item(0)) Then
                a = 0
            Else

                a = drSQL1.Item(0)
            End If



            If a > 0 Then
                ProgressBarProjectProgress.Value = (a * 100) / 100
                ProgressBarProjectProgress.Update()

                lblcompletion.Text = Convert.ToString(a) + "%"
            Else
                ProgressBarProjectProgress.Value = 1
                ProgressBarProjectProgress.Update()
                lblcompletion.Text = Convert.ToString(0) + "%"

            End If

        End If

        'Select SUM(Final)
        'FROM([FSPrograms].[dbo].[TSS_ENQ_Project_ProgressPerc])

        'Timer1.Enabled = True 'activate Timer1

        ' Timer1.Interval = 100 'set intervalo to 100

        'ProgressBarProjectProgress.Value = 1 'set ProgressBar value  

        cnSQL1.Close()


    End Sub
    Private Sub quoteupdate()


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cnSQL4 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim cmSQL3 As SqlCommand
        Dim cmSQL4 As SqlCommand

        Dim drSQL1 As SqlDataReader
        Dim drsql3 As SqlDataReader
        Dim strsql As String
        Dim strsql1 As String
        Dim strsql2 As String

        curdate = System.DateTime.Now()



        strsql = "SELECT CONumber, QtnDate,EnqNum  FROM [TSSAdditionaldata].[dbo].[TSS_Quote_EnqNo_2013] where EnqNum = " & txtEnqRegNo.Text & ""

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strsql, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then


            If IsDBNull(drSQL1.Item(0)) Then
                quote = "NO"
            Else

                quote = drSQL1.Item(0)
                quotedate = drSQL1.Item(1)
            End If
        Else
            quote = "NO"

        End If

        If quote <> "NO" Then

            strsql1 = "SELECT  Enq_Reg_No FROM ENQ_Project_Masterlist_Update WHERE  Project_MainStage = 'Quote Submitted to Customer' and Enq_Reg_No = '" & txtEnqRegNo.Text & "'"
            cnSQL3.Open()
            cmSQL3 = New SqlCommand(strsql1, cnSQL3)
            drsql3 = cmSQL3.ExecuteReader()

            projectintcodegen()

            If drsql3.Read() Then

                If IsDBNull(drsql3.Item(0)) Then


                    strsql2 = "insert ENQ_Project_Masterlist_Update values(" & txtprojectintcode.Text & ", " & txtEnqRegNo.Text & ", '" & DataGridViewProjectMasterList.CurrentRow.Cells(1).Value & "'," & _
                                   "'Quote Submitted to Customer', '-', 'Completed', '" & Format(quotedate, "dd-MMM-yyyy") & "'," & _
                                   " '" & Format(quotedate, "dd-MMM-yyyy") & "', '" & username & "', 'This entry automatically updated', '-'," & _
                                   " '" & username & "', '" & curdate & "','" & curdate & "','-'"

                    cmSQL4 = New SqlCommand(strsql2, cnSQL4)


                    If cmSQL4.ExecuteNonQuery() = 0 Then
                        MsgBox("Quote details not feteched automatically, for your information" & strsql, MsgBoxStyle.Exclamation, "Error!")
                        txtprojectintcode.Text = ""
                    Else
                        MsgBox("Quote details feteched automatically, for your information.", vbInformation)
                        txtprojectintcode.Text = ""
                    End If
                Else

                End If




            Else

                strsql2 = "insert ENQ_Project_Masterlist_Update values(" & txtprojectintcode.Text & ", " & txtEnqRegNo.Text & ", '" & DataGridViewProjectMasterList.CurrentRow.Cells(1).Value & "'," & _
                               "'Quote Submitted to Customer', '-', 'Completed', '" & Format(quotedate, "dd-MMM-yyyy") & "'," & _
                               " '" & Format(quotedate, "dd-MMM-yyyy") & "', '" & username & "', 'This entry automatically updated', '-'," & _
                               " '" & username & "', '" & curdate & "','" & curdate & "','-')"

                cnSQL4.Open()
                cmSQL4 = New SqlCommand(strsql2, cnSQL4)


                If cmSQL4.ExecuteNonQuery() = 0 Then
                    MsgBox("Quote details not feteched automatically, for your information" & strsql, MsgBoxStyle.Exclamation, "Error!")
                    txtprojectintcode.Text = ""
                Else
                    MsgBox("Quote details feteched automatically, for your information.", vbInformation)
                    txtprojectintcode.Text = ""
                End If



            End If

        End If


        cnSQL1.Close()


    End Sub

    Private Sub DataGridViewProjectMasterList_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridViewProjectMasterList_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs)



        'For Each row As DataGridViewRow In DataGridViewProjectMasterList.Rows



        '    If Trim(row.Cells("Project_Alotted").Value) = "Bharath" Then

        '        e.CellStyle.BackColor = Color.LightSkyBlue
        '    ElseIf Trim(row.Cells("Project_Alotted").Value) = "Deepak" Then
        '        e.CellStyle.BackColor = Color.LightYellow

        '    End If

        'Next

        'Dim drv As DataRowView
        'If e.RowIndex >= 0 Then
        '    If e.RowIndex <= ds.Tables("Employee").Rows.Count - 1 Then
        '        drv = ds.Tables("Employee").DefaultView.Item(e.RowIndex)
        '        Dim c As Color
        '        If drv.Item("Gender").ToString = "M" Then
        '            c = Color.LightBlue
        '        Else
        '            c = Color.Pink
        '        End If
        '        e.CellStyle.BackColor = c
        '    End If
        'End If


    End Sub

    Private Sub DataGridViewProjectMasterList_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        txtEnqRegNo.Text = DataGridViewProjectMasterList.CurrentRow.Cells(0).Value.ToString

        If GroupBoxProject.Visible = True Then
            GroupBoxProject.Visible = False
        End If

        '   If GroupBoxenqDetails.Visible = True Then
        'GroupBoxenqDetails.Visible = False
        'End If
        '
        If GroupBoxCustDetails.Visible = True Then
            GroupBoxCustDetails.Visible = False
        End If

        If GroupBoxItemDetails.Visible = True Then
            GroupBoxCustDetails.Visible = True
        End If

    End Sub

    Private Sub DataGridViewProjectMasterList_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)

        'If e.RowIndex < Me.dgv_EmployeeTraining.RowCount - 1 Then
        '    Dim dgvRow As DataGridViewRow = Me.dgv_EmployeeTraining.Rows(e.RowIndex)

        '    '<== This is the header Name
        '    'If CInt(dgvRow.Cells("EmployeeStatus_Training_e26").Value) <> 2 Then  


        '    '<== But this is the name assigned to it in the properties of the control
        '    If CInt(dgvRow.Cells("DataGridViewTextBoxColumn15").Value.ToString) <> 2 Then

        '        dgvRow.DefaultC ellStyle.BackColor = Color.FromArgb(236, 236, 255)

        '    Else
        '        dgvRow.DefaultCellStyle.BackColor = Color.LightPink

        '    End If

        'End If
    End Sub

    Private Sub GroupBoxProjectUpdate_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxProjectUpdate.Enter

    End Sub

    Private Sub GroupBox1_Enter_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxDetailsUpdate.Enter

    End Sub

    Private Sub ButtonDetailsUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDetailsUpdate.Click

        'check substages avble or not for the selected main stage
        Dim s As Integer
        Dim k As Integer
        s = 0
        k = 0

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        strSQL = "SELECT  count(Project_SubStatus) FROM TSS_ENQ_Project_Main_SubStages where Project_MainStatus = '" & txtCurMainStage.Text & "' "
        cnSQL.Open()
        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()

        If drSQL.Read() Then


            If IsDBNull(drSQL.Item(0)) Then
                s = 0
            Else

                s = drSQL.Item(0)

            End If
        End If

        If s > 0 Then



            k = CheckedListBoxSubStatus.CheckedItems.Count

            GroupBoxDetailsUpdate.Text = ""

            If k > 0 Then
                'if sub stage existing then only it should happen

                GroupBoxDetailsUpdate.Enabled = True
                txtcursubstage.Text = CheckedListBoxSubStatus.CheckedItems.Item(k - 1)



            End If

        End If


        If k > 0 Then
            GroupBoxDetailsUpdate.Text = "Call/visit details update of Stage:" & txtCurMainStage.Text & "-" & CheckedListBoxSubStatus.CheckedItems.Item(k - 1)

        Else
            GroupBoxDetailsUpdate.Enabled = True
            GroupBoxDetailsUpdate.Text = ""

            GroupBoxDetailsUpdate.Text = "Call/visit details update of Stage:" & txtCurMainStage.Text
        End If
        ' End If

    End Sub

    Private Sub ButtonEnquiryDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEnquiryDetails.Click
        If txtEnqRegNo.Text = "" Then
            MsgBox("Select the project before viewing enquiry details", vbInformation)
            Exit Sub

        End If
        GroupBoxEnqDetails1.Location = New Point(6, 435)
        GroupBoxEnqDetails1.Visible = True
        GroupBoxEnqDetails1.Width = 1188
        GroupBoxEnqDetails1.Height = 300
        EnqDetails()


    End Sub
    Private Sub EnqDetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "SELECT RegNo,  [Reg.Date], CustomerID, CustomerName, City, Class,Class1, Cust_Exist_New as Exis_Cust, CSR, TSSISeg, TSSSeg,MarketType, " & _
        " Enq_Ref_no, Enq_Ref_date, " & _
              "Enq_Source, Enq_Recd_date,Doc_upload,Doc_Details,Special_instructions from TSS_Enq_Pending_Project_Aproval where RegNo = " & txtEnqRegNo.Text & ""

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            txtRegNo1.Text = drSQL1.Item(0)
            DtpRegDate1.Value = drSQL1.Item(1)
            txtcustomerid1.Text = drSQL1.Item(2)
            txtcustomer1.Text = drSQL1.Item(3)
            txtcustcity1.Text = drSQL1.Item(4)
            txtEnqRef1.Text = drSQL1.Item(12)

            dtpActionStartDate1.Value = drSQL1.Item(13)

            dtpActionStartDate1.Format = DateTimePickerFormat.Custom
            dtpActionStartDate1.CustomFormat = "MMM yyyy"



            txtEnqSource1.Text = drSQL1.Item(14)
            DTPEnqRecd1.Value = drSQL1.Item(15)


            If Trim(drSQL1.Item(16)) = "YES" Then
                rbdocyes1.Checked = True
                rbdocno1.Checked = False
            Else
                rbdocno1.Checked = True
                rbdocyes1.Checked = True
            End If

            txtdocdetails1.Text = drSQL1.Item(17)
            txtspecial1.Text = drSQL1.Item(18)

        End If


    End Sub

    Private Sub dtpEnqDt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpActionStartDate.ValueChanged

    End Sub

    Private Sub BtnCustClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBoxEnqDetails1.Visible = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCustDetails.Click
        If txtEnqRegNo.Text = "" Then
            MsgBox("Select the project before viewing customer details", vbInformation)
            Exit Sub

        End If
        GroupBoxCustDetails.Location = New Point(6, 435)
        GroupBoxCustDetails.Visible = True
        GroupBoxCustDetails.Height = 300
        GroupBoxCustDetails.Width = 1188
        custdetails()

    End Sub
    Private Sub custdetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        'If Len(txtcustomerid.Text) = 6 Then

        If DataGridViewProjectMasterList.CurrentRow.Cells(15).Value = 0 Then

            strSQL1 = "select '', CustomerName, CustomerAddress1, CustomerAddress2,'', CustomerCity, CustomerZip,CustomerState,  CustomerCountry, CustomerContact, " & _
            "'','',CustomerContactPhone,CustomerContactFax,CustomerContactEmail,'','','','', CustomerClass3,CSR,'',CustomerClass7, FOBPoint,CustomerID,'','','','','',CustomerClass1 from FSDBBR.dbo.FS_Customer where  CustomerID  = '" & DataGridViewProjectMasterList.CurrentRow.Cells(16).Value & "'"

            'ist line 9 'custclssss3 -19
            'nt_code FROM ENQ_Details where Enq_Int_code = " & DataGridViewProjectMasterList.CurrentRow.Cells(12).Value & "

            'SELECT     CustomerID, CustomerName, CustomerAddress1, CustomerAddress2, CustomerCity, CustomerState, CustomerZip, CustomerCountry, CustomerContact, 
            'CustomerContactPhone, CSR, CustomerClass1, CustomerClass3, CustomerClass7, FOBPoint
            'FROM(FS_Customer)


        Else
            strSQL1 = "select * from ENQ_New_Customers where Cust_IntCode = " & DataGridViewProjectMasterList.CurrentRow.Cells(15).Value & ""


        End If
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then



            txtCId.Text = drSQL1.Item(24)
            txtCName.Text = drSQL1.Item(1)

            txtCustad1.Text = drSQL1.Item(2)
            txtcustadr2.Text = drSQL1.Item(3)
            txtcustadr3.Text = drSQL1.Item(4)
            txtCcity.Text = drSQL1.Item(5)
            txtcustzip.Text = drSQL1.Item(6)
            txtCustState.Text = drSQL1.Item(7)
            txtCustCountry.Text = drSQL1.Item(8)
            txtContact.Text = drSQL1.Item(9)
            txtDesignation.Text = drSQL1.Item(10)
            txtDept.Text = drSQL1.Item(11)
            txtMobile.Text = drSQL1.Item(12)
            txtPhone.Text = drSQL1.Item(13)
            txtFax.Text = drSQL1.Item(14)
            txtemail.Text = drSQL1.Item(15)

            txtcustclass.Text = Trim(drSQL1.Item(19))
            txtcustcsr.Text = Trim(drSQL1.Item(20))
            txtcustisr.Text = Trim(drSQL1.Item(21))
            txtcustTSSISeg.Text = Trim(drSQL1.Item(23))
            txtcustTSSSeg.Text = Trim(drSQL1.Item(22))

            txtcustclass1.Text = Trim(drSQL1.Item(30))



        End If


    End Sub



    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBoxCustDetails.Visible = False
    End Sub

    ' Private Sub GroupBoxItemDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxItemDetails.Enter

    ' End Sub

    Private Sub ButtonPartList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPartList.Click
        If txtEnqRegNo.Text = "" Then
            MsgBox("Select the project before viewing part details", vbInformation)
            Exit Sub

        End If
        GroupBoxItemDetails.Location = New Point(6, 435)
        GroupBoxItemDetails.Visible = True
        GroupBoxItemDetails.Height = 300
        GroupBoxItemDetails.Width = 1188
        'cleardetails()
        fillPartList()
        listloadCertificateDetails()

    End Sub
    Private Sub listloadCertificateDetails()

        For m As Integer = 0 To CheckedListBoxCertificate.Items.Count - 1
            CheckedListBoxCertificate.SetItemChecked(m, False)
        Next


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        'Dim a As ListView


        cnSQL1.Open()
        strSQL1 = "SELECT Certificates,Int_code FROM ENQ_Certificates " & _
                 "WHERE  Status = 'A'"
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        Dim ColumnValue As String = Nothing
        Do While drSQL1.Read()

            ColumnValue = (drSQL1.GetValue(0)).ToString
            CheckedListBoxCertificate.Items.Add(ColumnValue)
            CheckedListBoxCertificate.ValueMember = "Int_code"

        Loop


    End Sub
    Sub fillPartList()

        DataGridViewItemDetail.Show()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        'strSql = "SELECT Sl_no,PartNumber,'' as Tss_Drawin_No, PartDescription,CustPartNumber,CustPartDescription,uom,Dimension,Material,FS_Yes_NO,Part_Source,Special,Req,Enq_Detail_code as DetailKey,Enq_Int_code FROM ENQ_Details where Enq_Int_code = " & DataGridViewProjectMasterList.CurrentRow.Cells(13).Value & " order by Sl_no"

        strSql = "SELECT  Sl_no, PartNumber, Tss_Drawin_No, Part_Accepted, PartDescription, CustPartNumber, CustPartDescription, uom, Dimension, Material, FS_Yes_NO, Special, Req, DetailKey, Enq_Int_code FROM TSS_Enq_Details_Drawg_No where Enq_Int_code = " & DataGridViewProjectMasterList.CurrentRow.Cells(14).Value & " order by Sl_no"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd

        sqlCon.Open()
        stockDAC.TableMappings.Add("Table", "Part")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewItemDetail.DataSource = stockDC.Tables(0)
        sqlCon.Close()


        'DataGridViewItemDetail.Update()


    End Sub

    Private Sub ButtonItemDetailsClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBoxItemDetails.Visible = False
    End Sub

    Private Sub ButtonApprovalDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonApprovalDetails.Click
        If txtEnqRegNo.Text = "" Then
            MsgBox("Select the project before viewing approval details", vbInformation)
            Exit Sub

        End If

        GroupBoxProject.Location = New Point(6, 435)
        GroupBoxProject.Visible = True
        GroupBoxProject.Height = 300
        GroupBoxProject.Width = 1188

        filldocdetails()
        ApprovalDetails()
    End Sub
    Private Sub filldocdetails()

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        '    Dim i As Integer

        CheckedListBoxDoc.Items.Clear()

        Dim ColumnValue As String = Nothing


        strsql = "Select Documents from ENQ_Projects_Documents_Status where Enq_Reg_No = " & txtEnqRegNo.Text & " "

        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        ' i = CheckedListBoxDoc.Items.Count

        Do While drSQL1.Read()

            ColumnValue = (drSQL1.GetValue(0)).ToString
            CheckedListBoxDoc.Items.Add(ColumnValue)
            CheckedListBoxDoc.ValueMember = "Project_MainCode"

        Loop

        For idx As Integer = 0 To Me.CheckedListBoxDoc.Items.Count - 1
            Me.CheckedListBoxDoc.SetItemCheckState(idx, CheckState.Checked)
        Next


    End Sub


    Private Sub ButtonApprovalClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'GroupBoxApprovalDetails.Visible = False

    End Sub

    Private Sub BtnCustClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCustClose.Click
        GroupBoxEnqDetails1.Visible = False

    End Sub

    Private Sub GroupBoxenqDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxenqDetails.Enter

    End Sub

    Private Sub DtpRegDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DataGridViewItemDetail_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridViewItemDetail_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        fillqtydetails()
        fillcertificates()
    End Sub
    Private Sub fillcertificates()


        For m As Integer = 0 To CheckedListBoxCertificate.Items.Count - 1
            CheckedListBoxCertificate.SetItemChecked(m, False)
        Next


        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert As String
        'Dim b As Integer

        strsql = "Select Certificates from ENQ_EnqWise_Certificates where Enq_Int_code =  " & DataGridViewItemDetail.CurrentRow.Cells(14).Value & " and Enq_Detail_code =  " & DataGridViewItemDetail.CurrentRow.Cells(13).Value & " "

        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        i = CheckedListBoxCertificate.Items.Count

        Do While drSQL1.Read()
            cert = drSQL1.Item(0)
            a = 0
            Do While a < i

                If cert = CheckedListBoxCertificate.Items(a) Then

                    CheckedListBoxCertificate.SetItemChecked(a, True)

                    a = i
                Else

                    a = a + 1

                End If


            Loop

        Loop

    End Sub

    Public Sub fillqtydetails()

        DataGridQty.Show()
        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim stockDCQ As DataSet = New DataSet

        'Dim cmSQL As SqlCommand
        ' Dim drSQL As SqlDataReader

        strSql = "SELECT Qty,Qty_Type, Enq_Qty_IntCode FROM ENQ_Qty_Details " & _
                "WHERE  Enq_Int_code =  " & DataGridViewItemDetail.CurrentRow.Cells(14).Value & "  and Enq_Detail_code =  " & DataGridViewItemDetail.CurrentRow.Cells(13).Value & " " & _
                  "ORDER BY Qty"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDQC As SqlDataAdapter = New SqlDataAdapter

        stockDQC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDQC.TableMappings.Add("Table", "Part")
        'get data
        stockDQC.Fill(stockDCQ)

        DataGridQty.DataSource = stockDCQ.Tables(0)
        sqlCon.Close()
        DataGridQty.Expand(-1)


    End Sub


    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Buttoncustomerclose.Click
        GroupBoxCustDetails.Visible = False

    End Sub

    Private Sub ButtonItemDetailsClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItemDetailsClose.Click
        GroupBoxItemDetails.Visible = False

    End Sub

    Private Sub DataGridViewItemDetail_RowHeaderMouseClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridViewItemDetail.RowHeaderMouseClick
        'MsgBox("WHAT ")

        fillqtydetails()
        fillcertificates()


    End Sub

    Private Sub DataGridViewItemDetail_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewItemDetail.CellContentClick

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ButtonProjectAprClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBoxProject.Visible = False

    End Sub
    Private Sub ApprovalDetails()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "SELECT * FROM  ENQ_Project_Approval_Status where Enq_Reg_No = " & txtEnqRegNo.Text & ""



        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            txtProjectRemarks.Text = Trim(drSQL1.Item(2))
            txtProjectNumber.Text = Trim(drSQL1.Item(3))
            txtProjectAlloted.Text = Trim(drSQL1.Item(5))
            txtProjectType.Text = Trim(drSQL1.Item(9))
            txtEarNumber.Text = Trim(drSQL1.Item(10))
            txtProject.Text = Trim(drSQL1.Item(11))
            txtApplication.Text = Trim(drSQL1.Item(12))




        End If


    End Sub
    Public Sub projectintcodegen()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(Int_code)from ENQ_Project_Masterlist_Update"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtprojectintcode.Text = 1
            Else
                txtprojectintcode.Text = drSQL1.Item(0) + 1
            End If


        End If

    End Sub


    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrnMasterListUpdate.Click


        If GroupBoxDetailsUpdate.Text = "" Or GroupBoxDetailsUpdate.Text = "-" Then
            MsgBox("Pl click on Detail Update.", vbInformation)
            Exit Sub
        End If

        If ComboboxSendInfo.Text = "Price Request to Purchase Dept" Then
            If LblRecVendor.Visible = False Then
                LblRecVendor.Visible = True
                txtRecVendor.Visible = True
            End If

            If Len(txtRecVendor.Text) < 3 Then
                MsgBox("Pl enter recommended vendor Name,vbinformation")
                Exit Sub
            End If
        End If

        Dim strsql As String
        Dim cmSQL As SqlCommand
        'Dim st As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        curdate = System.DateTime.Now()

        If mode = "ADD" Then

            If ComboBoxActionStatus.Text = "Completed" Then
                checkstatus = "YES"
                checksavestatus()
            End If


            If txtCurMainStage.Text = "Lost" Then
                checklost = "YES"
                checkloststatus()

            End If




            If ComboBoxActionStatus.Text = "Completed" And txtCurMainStage.Text <> "Lost" And checkstatus = "YES" Then
                seq = "YES"
                'checkprevcompstages() 'commented this procedure till Jan-14. this procedure to check user has to save all stage sequentially.
            End If



            If checkstatus = "NO" Or checklost = "NO" Then
                '        MsgBox("This status is already completed.", vbInformation)

                Exit Sub
            ElseIf seq = "NO" Then
                MsgBox("Project stages needs to be saved sequentially.  Earlier stages are not completed", vbInformation)
                Exit Sub
            End If



            projectintcodegen()

            strsql = "insert ENQ_Project_Masterlist_Update values(" & txtprojectintcode.Text & ", " & txtEnqRegNo.Text & ", '" & DataGridViewProjectMasterList.CurrentRow.Cells(1).Value & "'," & _
            "'" & txtCurMainStage.Text & "', '" & txtcursubstage.Text & "', '" & ComboBoxActionStatus.Text & "', '" & Format(dtpActionStartDate.Value, "dd-MMM-yyyy") & "'," & _
            " '" & Format(dtpActionEndDate.Value, "dd-MMM-yyyy") & "', '" & ComboBoxActionBy.Text & "', '" & txtCurStageremarks.Text & "', '" & ComboboxSendInfo.Text & "'," & _
            " '" & username & "', '" & curdate & "','" & curdate & "','" & txtRecVendor.Text & "')"


        End If

        If mode = "EDIT" Then

            strsql = "update ENQ_Project_Masterlist_Update  set " & _
            "Project_MainStage = '" & txtCurMainStage.Text & "', Project_SubStage = '" & txtcursubstage.Text & "',Project_Stage_Status = '" & ComboBoxActionStatus.Text & "'," & _
            "Project_Stage_StartDate =  '" & Format(dtpActionStartDate.Value, "dd-MMM-yyyy") & "',Project_Stage_EndDate = '" & Format(dtpActionEndDate.Value, "dd-MMM-yyyy") & "'," & _
            "Project_Stage_ActionBy = '" & ComboBoxActionBy.Text & "', Project_Stage_Details = '" & txtCurStageremarks.Text & "',Send_Info_To= '" & ComboboxSendInfo.Text & "' " & _
            "Rec_Vendor = '" & txtRecVendor.Text & "'" & _
            " where Enq_Reg_No= " & txtEnqRegNo.Text & " and Int_code  = " & DataGridViewProjectMasterEdit.CurrentRow.Cells(11).Value & " "


        End If

        cnSQL.Open()

        cmSQL = New SqlCommand(strsql, cnSQL)

        'MsgBox(Err.Number)

        If cmSQL.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot save details" & strsql, MsgBoxStyle.Exclamation, "Error!")
            Exit Sub

        Else
            MsgBox("Details saved.", vbInformation)
            GroupBoxDetailsUpdate.Text = ""
            GroupBoxDetailsUpdate.Enabled = False
            txtCurStageremarks.Text = ""

            ' Exit Sub
        End If


        'update enquiry details as itemstatus = 'p' if price requested selected.'strsql = "update ENQ_Project_Approval_Status  set " & _
        '"Status = '" & st & "', Remarks = '" & txCurStageremarks.Text & "', Project_Number = '" & txtProjectNumber.Text & "', " & _


        If ComboboxSendInfo.Text = "Price Request to Purchase Dept" Then

            strsql = "update ENQ_Details  set ItemStatus = 'P' where Enq_Int_code = '" & DataGridViewProjectMasterList.CurrentRow.Cells(13).Value & "'  and (Req = 'Price' or Req = 'Both')"
            'cnSQL.Open()

            cmSQL = New SqlCommand(strsql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot save Enquiry details" & strsql, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub

            Else
                '   MsgBox("Details saved.", vbInformation)
                '  Exit Sub
            End If




        End If








    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcursubstage.TextChanged

    End Sub

    Private Sub txtCurStageremarks_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCurStageremarks.TextChanged

    End Sub

    Private Sub DataGridViewProjectMasterList_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewProjectMasterList.CellContentClick

    End Sub

    Private Sub DataGridViewProjectMasterList_RowHeaderMouseClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridViewProjectMasterList.RowHeaderMouseClick

        txtEnqRegNo.Text = DataGridViewProjectMasterList.CurrentRow.Cells(0).Value.ToString

        progress()

        quoteupdate()


        If GroupBoxProject.Visible = True Then
            GroupBoxProject.Visible = False
        End If
        If GroupBoxEnqDetails1.Visible = True Then
            GroupBoxEnqDetails1.Visible = False
        End If

        If GroupBoxCustDetails.Visible = True Then
            GroupBoxCustDetails.Visible = False
        End If

        If GroupBoxItemDetails.Visible = True Then
            GroupBoxCustDetails.Visible = True
        End If
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEnquiryDetails.Click
        GroupBoxEnqDetails1.Visible = False

    End Sub

    Private Sub GroupBoxCustDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxCustDetails.Enter

    End Sub

    Private Sub GroupBoxProjectDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxProjectDetails.Enter

    End Sub

    Private Sub BtnTssDrwgSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnTssDrwgSave.Click

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim strSql As String
        Dim cmSQL As SqlCommand

        cnSQL.Open()

        checktssdrwing()

        For i As Integer = 0 To DataGridViewItemDetail.RowCount - 1

            strSql = "insert ENQ_Project_Items_TSSDrwgNo  values(" & DataGridViewItemDetail.Rows(i).Cells("Enq_Int_Code").Value & "," & DataGridViewItemDetail.Rows(i).Cells("DetailKey").Value & ",'" & DataGridViewItemDetail.Rows(i).Cells("PartNumber").Value & "','" & DataGridViewItemDetail.Rows(i).Cells("Tss_Drawin_No").Value & "','" & DataGridViewItemDetail.Rows(i).Cells("Part_Accepted").Value & "')"

            cmSQL = New SqlCommand(strSql, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot save details" & strSql, MsgBoxStyle.Exclamation, "Error!")
                Application.Exit()

            End If

        Next
        MsgBox("Data saved", vbInformation)
        Exit Sub

    End Sub

    Private Sub checktssdrwing()
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim cmsql2 As SqlCommand

        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String
        Dim strsql2 As String


        strSQL1 = "select * from ENQ_Project_Items_TSSDrwgNo WHERE  Enq_Int_Code = " & DataGridViewItemDetail.CurrentRow.Cells(14).Value & ""
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then

            Else
                strsql2 = "delete   from ENQ_Project_Items_TSSDrwgNo WHERE  Enq_Int_Code = " & DataGridViewItemDetail.CurrentRow.Cells(14).Value & ""
                cnSQL2.Open()
                cmsql2 = New SqlCommand(strsql2, cnSQL2)


                If cmsql2.ExecuteNonQuery() = 0 Then


                End If

            End If

        End If

    End Sub


    Private Sub ButtonProjectAprClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonProjectAprClose.Click
        GroupBoxProject.Visible = False

    End Sub

    Private Sub ButtonProjDetailEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonProjDetailEdit.Click

        If Val(txtEnqRegNo.Text) > 0 Then


            mode = "EDIT"
            DataGridViewProjectMasterEdit.Columns.Clear()

            'adding check box
            'Dim checkCol As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            'checkCol.HeaderText = "Del"
            'DataGridViewProjectMasterEdit.Columns.Add(checkCol)
            'end of adding check box


            Dim AlarmColumn1 As New DataGridViewCheckBoxColumn(False)

            AlarmColumn1.Name = "Del"
            AlarmColumn1.HeaderText = "Delete"
            AlarmColumn1.ReadOnly = False


            DataGridViewProjectMasterEdit.Columns.Add(AlarmColumn1)
            DataGridViewProjectMasterEdit.ReadOnly = False


            ButtonProjDetailsDelete.Enabled = True

            GroupBoxEdit.Visible = True
            DataGridViewProjectMasterEdit.Visible = True
            DataGridViewProjectMasterEdit.Enabled = True

            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim strSQL As String


            Dim stockDC As DataSet = New DataSet

            strSQL = "SELECT  Enq_Reg_No, Project_Number, Project_MainStage, Project_SubStage, Project_Stage_Status as Status, Project_Stage_StartDate as StartDate, Project_Stage_EndDate as EndDate, Project_Stage_ActionBy as ActionBy, Project_Stage_Details as Details,Send_Info_To,Int_code FROM ENQ_Project_Masterlist_Update " & _
            " where  Enq_Reg_No = " & txtEnqRegNo.Text & " order by Int_code "



            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAC.TableMappings.Add("Table", "Enq")
            'get data
            stockDAC.Fill(stockDC)


            DataGridViewProjectMasterEdit.DataSource = stockDC.Tables(0)



            'DataGridViewProjectMasterEdit.Columns.Add(checkCol)



            cnSQL.Close()

        Else
            MsgBox("Pl select the project", vbInformation)
            Exit Sub
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub ButtonGridClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGridClose.Click
        GroupBoxEdit.Visible = False
        mode = "ADD"
    End Sub

    Private Sub DataGridViewProjectMasterEdit_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewProjectMasterEdit.CellContentClick


        Dim ColumnName1 As String = DataGridViewProjectMasterEdit.Columns(e.ColumnIndex).Name

        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)


        If ColumnName1 = "Del" Then
            Dim CellCheckBox1 As DataGridViewCheckBoxCell = _
                CType(DataGridViewProjectMasterEdit.Rows(e.RowIndex).Cells(ColumnName1), DataGridViewCheckBoxCell)

            Dim CellCheckBoxState1 As String = CellCheckBox1.EditingCellFormattedValue.ToString


            DataGridViewProjectMasterEdit.Rows(CellCheckBox1.RowIndex).Cells(0).Value = True

            Dim msgb As String
            msgb = MsgBox("Are you sure of deleting this line ?", vbYesNo)

            If msgb = vbNo Then
                DataGridViewProjectMasterEdit.Rows(CellCheckBox1.RowIndex).Cells(0).Value = False
                Exit Sub
            Else
                'deletion procedure


                strsql = "delete from  ENQ_Project_Masterlist_Update where Enq_Reg_No= " & txtEnqRegNo.Text & " and Int_code  = " & DataGridViewProjectMasterEdit.CurrentRow.Cells(11).Value & " "

                cnSQL.Open()

                cmSQL = New SqlCommand(strsql, cnSQL)


                If cmSQL.ExecuteNonQuery() = 0 Then
                    MsgBox("Cannot delete" & strsql, MsgBoxStyle.Exclamation, "Error!")
                    Exit Sub

                Else
                    MsgBox("deleted", vbInformation)
                    'Exit Sub

                End If

            End If
        End If

        'CALLING REFRESH AGAIN

        DataGridViewProjectMasterEdit.Columns.Clear()


        Dim AlarmColumn1 As New DataGridViewCheckBoxColumn(False)

        AlarmColumn1.Name = "Del"
        AlarmColumn1.HeaderText = "Delete"
        AlarmColumn1.ReadOnly = False


        DataGridViewProjectMasterEdit.Columns.Add(AlarmColumn1)
        DataGridViewProjectMasterEdit.ReadOnly = False


        ButtonProjDetailsDelete.Enabled = True

        GroupBoxEdit.Visible = True
        DataGridViewProjectMasterEdit.Visible = True
        DataGridViewProjectMasterEdit.Enabled = True

        ' Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        strsql = "SELECT  Enq_Reg_No, Project_Number, Project_MainStage, Project_SubStage, Project_Stage_Status as Status, Project_Stage_StartDate as StartDate, Project_Stage_EndDate as EndDate, Project_Stage_ActionBy as ActionBy, Project_Stage_Details as Details,Send_Info_To,Int_code FROM ENQ_Project_Masterlist_Update " & _
        " where  Enq_Reg_No = " & txtEnqRegNo.Text & " order by Int_code "



        Dim sqlCmd As SqlCommand = New SqlCommand(strsql, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        ' cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewProjectMasterEdit.DataSource = stockDC.Tables(0)






        cnSQL.Close()

    End Sub


    Private Sub DataGridViewProjectMasterEdit_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridViewProjectMasterEdit.RowHeaderMouseClick

        'strSQL = "SELECT  Enq_Reg_No, Project_Number, Project_MainStage, Project_SubStage, 4
        'Project_Stage_Status as Status, Project_Stage_StartDate as StartDate, Project_Stage_EndDate as EndDate, 7
        'Project_Stage_ActionBy as ActionBy, Project_Stage_Details as Details,  Int_code FROM ENQ_Project_Masterlist_Update " & _
        '" where  Enq_Reg_No = " & txtEnqRegNo.Text & ""

        mode = "EDIT"

        txtCurMainStage.Text = ""
        txtcursubstage.Text = ""


        listloadMainStages()
        MainStages()

        txtCurMainStage.Text = DataGridViewProjectMasterEdit.CurrentRow.Cells(3).Value.ToString
        CheckedListBoxSubStatus.Visible = True

        listloadSubStages()
        SubStages()

        ComboBoxActionBy.Text = DataGridViewProjectMasterEdit.CurrentRow.Cells(8).Value.ToString
        ComboBoxActionStatus.Text = DataGridViewProjectMasterEdit.CurrentRow.Cells(5).Value.ToString
        ComboboxSendInfo.Text = Trim(DataGridViewProjectMasterEdit.CurrentRow.Cells(10).Value.ToString)
        dtpActionStartDate.Value = DataGridViewProjectMasterEdit.CurrentRow.Cells(6).Value
        dtpActionEndDate.Value = DataGridViewProjectMasterEdit.CurrentRow.Cells(7).Value.ToString
        txtCurStageremarks.Text = DataGridViewProjectMasterEdit.CurrentRow.Cells(9).Value.ToString



        'txtEnqRegNo.Text = DataGridViewProjectMasterList.CurrentRow.Cells(0).Value.ToString





    End Sub

    Private Sub ButtonMainStage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMainStage.Click

    End Sub
    Private Sub MainStages()


        For m As Integer = 0 To CheckedListBoxMainStatus.Items.Count - 1
            CheckedListBoxMainStatus.SetItemChecked(m, False)
        Next


        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert As String
        'Dim b As Integer

        strsql = "Select Project_MainStage from ENQ_Project_Masterlist_Update where Enq_Reg_No =  " & txtEnqRegNo.Text & " and Int_code = " & DataGridViewProjectMasterEdit.CurrentRow.Cells(11).Value & ""

        'strsql = "SELECT  Enq_Reg_No, Project_Number, Project_MainStage, Project_SubStage, Project_Stage_Status as Status, Project_Stage_StartDate as StartDate, Project_Stage_EndDate as EndDate, Project_Stage_ActionBy as ActionBy, Project_Stage_Details as Details,Send_Info_To,Int_code FROM ENQ_Project_Masterlist_Update " & _
        '" where  Enq_Reg_No = " & txtEnqRegNo.Text & " order by Int_code "


        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        i = CheckedListBoxMainStatus.Items.Count

        Do While drSQL1.Read()
            cert = drSQL1.Item(0)
            a = 0
            Do While a < i

                If cert = CheckedListBoxMainStatus.Items(a) Then

                    CheckedListBoxMainStatus.SetItemChecked(a, True)

                    a = i
                Else

                    a = a + 1

                End If


            Loop

        Loop

    End Sub

    Private Sub SubStages()


        For m As Integer = 0 To CheckedListBoxSubStatus.Items.Count - 1
            CheckedListBoxSubStatus.SetItemChecked(m, False)
        Next


        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader

        cnSQL.Open()

        Dim i As Integer
        Dim a As Integer
        Dim cert1 As String
        'Dim b As Integer

        strsql = "Select Project_SubStage from ENQ_Project_Masterlist_Update where Enq_Reg_No =  " & txtEnqRegNo.Text & " and Int_code = " & DataGridViewProjectMasterEdit.CurrentRow.Cells(11).Value & " "

        'strsql = "SELECT  Enq_Reg_No, Project_Number, Project_MainStage, Project_SubStage, Project_Stage_Status as Status, Project_Stage_StartDate as StartDate, Project_Stage_EndDate as EndDate, Project_Stage_ActionBy as ActionBy, Project_Stage_Details as Details,Send_Info_To,Int_code FROM ENQ_Project_Masterlist_Update " & _
        '" where  Enq_Reg_No = " & txtEnqRegNo.Text & " order by Int_code "


        cmSQL = New SqlCommand(strsql, cnSQL)
        drSQL1 = cmSQL.ExecuteReader()

        i = CheckedListBoxSubStatus.Items.Count

        Do While drSQL1.Read()
            cert1 = drSQL1.Item(0)
            a = 0
            Do While a < i

                If cert1 = CheckedListBoxSubStatus.Items(a) Then

                    CheckedListBoxSubStatus.SetItemChecked(a, True)

                    a = i
                Else

                    a = a + 1

                End If


            Loop

        Loop

    End Sub


    Private Sub ButtonStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ButtonReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonReport.Click



    End Sub

    Private Sub ComboboxSendInfo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboboxSendInfo.KeyDown


    End Sub

    Private Sub ComboboxSendInfo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboboxSendInfo.LostFocus
        If ComboboxSendInfo.Text = "Price Request to Purchase Dept" Then
            If LblRecVendor.Visible = False Then
                LblRecVendor.Visible = True
                txtRecVendor.Visible = True
            End If
        End If


    End Sub

    Private Sub ComboboxSendInfo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboboxSendInfo.SelectedIndexChanged

    End Sub

    Private Sub ComboboxSendInfo_StyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboboxSendInfo.StyleChanged

    End Sub

    Private Sub CheckedListBoxCertificate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBoxCertificate.SelectedIndexChanged

    End Sub
    Private Sub checksavestatus()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand

        Dim drSQL1 As SqlDataReader
        Dim strsql1 As String

        curdate = System.DateTime.Now()



        strsql1 = "SELECT  Enq_Reg_No FROM ENQ_Project_Masterlist_Update WHERE  Project_MainStage = '" & txtCurMainStage.Text & " ' and Project_SubStage = '" & txtcursubstage.Text & " ' and Project_Stage_Status = 'Completed' and Enq_Reg_No = '" & txtEnqRegNo.Text & "'"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strsql1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                checkstatus = "YES"

            Else
                MsgBox("This status is already saved. Duplicate entry not allowed", vbInformation)
                checkstatus = "NO"
                Exit Sub
            End If

        Else
            'MsgBox("This status is already saved. Duplicate entry not allowed", vbInformation)
            checkstatus = "YES"
            'Exit Sub
        End If


        cnSQL1.Close()


    End Sub

    Private Sub checkprevcompstages()



        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim cmSQL2 As SqlCommand

        Dim drSQL1 As SqlDataReader
        Dim drSQL2 As SqlDataReader

        Dim strsql As String
        Dim strsql1 As String

        Dim count1 As Integer
        Dim count2 As Integer

        'curdate = System.DateTime.Now()

        If Len(txtcursubstage.Text) > 4 Then
            strsql = "SELECT count(*) FROM TSS_Enq_Project_Status_Master where  Project_SubCode < (select (Project_SubCode) from TSS_Enq_Project_Status_Master a where Project_SubStatus = '" & txtcursubstage.Text & "')"
        Else
            strsql = "SELECT count(*) FROM TSS_Enq_Project_Status_Master where  Project_MainCode < (select max(Project_MainCode) from TSS_Enq_Project_Status_Master a where Project_MainStatus = '" & txtCurMainStage.Text & "')"
        End If

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strsql, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then


            If IsDBNull(drSQL1.Item(0)) Then

            Else

                count1 = drSQL1.Item(0)
            End If


        End If
        cnSQL1.Close()

        If count1 > 0 Then

            strsql1 = "SELECT COUNT(*) FROM ENQ_Project_Masterlist_Update WHERE Project_Stage_Status = 'Completed' AND Enq_Reg_No = '" & txtEnqRegNo.Text & "'"

            cnSQL2.Open()
            cmSQL2 = New SqlCommand(strsql1, cnSQL2)
            drSQL2 = cmSQL2.ExecuteReader()


            If drSQL2.Read() Then

                If IsDBNull(drSQL2.Item(0)) Then

                Else

                    count2 = drSQL2.Item(0)
                End If


            End If

            If count1 <> count2 Then
                seq = "NO"

            End If
            cnSQL2.Close()

        End If
    End Sub

    Private Sub checkloststatus()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand

        Dim drSQL1 As SqlDataReader
        Dim strsql1 As String

        'curdate = System.DateTime.Now()

        strsql1 = "SELECT  Enq_Reg_No FROM ENQ_Project_Masterlist_Update WHERE  Project_MainStage = 'Production Order' and Project_Stage_Status = 'Completed' "
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strsql1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                checklost = "YES"

            Else
                MsgBox("Production Order is already recd. This  entry is invalid !!", vbInformation)
                checklost = "NO"
                Exit Sub
            End If


        Else
            'MsgBox("This status is already saved. Duplicate entry not allowed", vbInformation)
            checklost = "YES"
            'Exit Sub
        End If


        cnSQL1.Close()


    End Sub


    Private Sub ComboBoxActionBy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxActionBy.SelectedIndexChanged

    End Sub
End Class
