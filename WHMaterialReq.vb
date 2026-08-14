Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection

Public Class WHMaterialReq
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"
    Private Sub MyGroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub MaterialReq_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        
        RBNonFSItem.Checked = True

        transtype = "MatReq"

        comboload()

        DTPReqDt.Format = DateTimePickerFormat.Custom
        DTPReqDt.CustomFormat = "dd/MM/yyyy"

        txtReqNo.Enabled = False
        DTPReqDt.Enabled = False
        txtMO.Focus()



    End Sub

    Private Sub _DoubleClick(sender As Object, e As EventArgs) Handles txtPartNumber.DoubleClick

        '     If RBFSItem.Checked = True Then
        DataGridViewWHPartNumbers.Visible = True

        Whfillparts()

        'Exit Sub

        'End If



    End Sub

    Private Sub Whfillparts()

        DataGridViewWHPartNumbers.Visible = True
        DataGridViewWHPartNumbers.Show()



        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim strSql As String
        Dim stockDC As DataSet = New DataSet

        txtPartNumber.Text = txtPartNumber.Text & "%"


        If RBFSItem.Checked = True And ComboBoxIssueType.Text = "Issue" Then

            strSql = "SELECT distinct ItemNumber,ItemDescription ,ItemUM   FROM [FSDBBR].[dbo].[TSS_WH_ToolsCons_Stock_P] " & _
                     "WHERE  ItemNumber like '" & txtPartNumber.Text & "' " & _
                        "ORDER BY ItemNumber"

        ElseIf RBFSItem.Checked = True And ComboBoxIssueType.Text = "Return" Then

            strSql = "SELECT  ItemNumber,ItemDescription ,ItemUM    FROM [FSDBBR].[dbo].[_NoLock_FS_Item] where ItemNumber like 'U%'  and InventoryAccount  like '%-61000' " & _
                     "AND  ItemNumber like '" & txtPartNumber.Text & "' " & _
                        "ORDER BY ItemNumber"

        ElseIf RBNonFSItem.Checked = True Then

            strSql = "SELECT  Part_Number AS 'ItemNumber',Part_Description  AS 'ItemDescription', UOM AS 'ItemUM' FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] with (nolock) " & _
                     "WHERE  Part_Number like '" & txtPartNumber.Text & "' " & _
                     "ORDER BY Part_Number"



        End If



        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDAC.TableMappings.Add("Table", "Part")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewWHPartNumbers.Location = New System.Drawing.Point(247, 49)
        DataGridViewWHPartNumbers.Width = 600 '1000
        DataGridViewWHPartNumbers.Height = 243 ' 412



        DataGridViewWHPartNumbers.DataSource = stockDC.Tables(0)
        sqlCon.Close()
        '  DataGridViewWHPartNumbers.Expand(-1)

        DataGridViewWHPartNumbers.Columns("ItemNumber").Width = 200

        DataGridViewWHPartNumbers.Columns("ItemDescription").Width = 250


        DataGridViewWHPartNumbers.Columns("ItemUM").Width = 50


    End Sub

    Private Sub txtPartNumber_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPartNumber.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If


        'If e.KeyCode = Keys.Enter Then
        '    txtDescription.Focus()
        'End If
    End Sub

    Private Sub txtPartNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPartNumber.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtPartNumber_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtPartNumber.MouseDoubleClick







    End Sub

    Private Sub txtPartNumber_TextChanged(sender As Object, e As EventArgs) Handles txtPartNumber.TextChanged
        If txtSlNo.Text = "" And txtReqNo.Text = "" Then
            MsgBox("Click on add", vbInformation)
            Exit Sub
        End If

        If ComboBoxIssueType.Text = "Issue" Then
            LabelDetQty.Text = "Req. Qty"
            ComboBoxUYN.Enabled = False
            ComboBoxPur.Text = ""
            ComboBoxPur.Enabled = False

        ElseIf ComboBoxIssueType.Text = "Return" Then
            LabelDetQty.Text = "Return Qty"
            ComboBoxUYN.Enabled = True
            ComboBoxPur.Enabled = True
        End If

    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click



        clearall()
        txtMO.Focus()

        '    DataGridViewMatReqEdit.EditMode = DataGridViewEditMode.EditProgrammatically

    End Sub

    Private Sub GroupBoxMenu_Enter(sender As Object, e As EventArgs) Handles GroupBoxMenu.Enter

    End Sub
    Private Sub comboload()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand


        'dept_type load

        strSql = "SELECT * FROM [TSS_WH_Dept] " & _
                 "WHERE Status like 'A%' ORDER BY [Department]"
        cmSQL = New SqlCommand(strSql, sqlCon)
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource As SqlDataAdapter = New SqlDataAdapter
        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eitem")
        With ComboBoxdept
            .DataSource = source.Tables("eitem")
            .DisplayMember = "Department"
            .ValueMember = "DeptKey"
            .SelectedIndex = 0
        End With

        'SUB DIVISION
        strSql = "SELECT * FROM [TSS_WH_SUB_Division] " & _
                 "WHERE Status like 'A%' ORDER BY [Sub_Division]"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet

        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "esd")

        With ComboBoxSD
            .DataSource = source.Tables("esd")
            .DisplayMember = "Sub_Division"
            .ValueMember = "id"
            .SelectedIndex = 0
        End With

        'CELL
        strSql = "SELECT * FROM  TSS_WH_Cell " & _
                        "WHERE Status like 'A%' ORDER BY Cell_Name"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet

        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "ecell")
        With ComboBoxCell
            .DataSource = source.Tables("ecell")
            .DisplayMember = "Cell_Name"
            .ValueMember = "id"
            .SelectedIndex = 0
        End With

        'PURPOSE

        strSql = "SELECT *  FROM TSS_WH_Return_Purpose WHERE Status like 'A%' order by Return_Purpose"
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet

        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eret")
        With ComboBoxPur
            .DataSource = source.Tables("eret")
            .DisplayMember = "Return_Purpose"
            .ValueMember = "ID"
            .SelectedIndex = 0
        End With


        'strSql = "SELECT * FROM ENQ_Clarity " & _
        '                   "WHERE Status like 'A%' ORDER BY Clarity"
        'cmSQL = New SqlCommand(strSql, sqlCon)

        'sqlCmd = New SqlCommand(strSql, sqlCon)
        'ESource = New SqlDataAdapter
        'source = New DataSet


        'ESource.SelectCommand = sqlCmd
        'ESource.Fill(source, "eclarity")
        'With ComboBoxClarity
        '    .DataSource = source.Tables("eclarity")
        '    .DisplayMember = "Clarity"
        '    .ValueMember = "Int_code"
        '    .SelectedIndex = 0
        'End With




    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub slnogenerate()
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "Select MAX(Slno) from [TSS_WH_MaterialRequestDetail] where [MatReq_no] = " & txtReqNo.Text & "  and User_Id = '" & username & "' "
        cnSQL1.Open()

        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then
            txtSlNo.Text = drSQL1.Item(0) + 1
        End If
        cnSQL1.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnReqSave.Click
        '      transmode = "Add"
        '     If txtReqNo.Text = "" And transmode = "Add" Then
        ''generate req no.
        'nogenerate()

        'End If

        'header save - insert query

        If RBNonFSItem.Checked = True Then

            If ComboBoxIssueType.Text = "Return" Then
                MsgBox("Return is possible  only for FS items", vbInformation)
                Exit Sub
            End If
        End If


        If ComboBoxIssueType.Text = "Return" Then


            If Len(ComboBoxUYN.Text) <= 2 Or Len(ComboBoxPur.Text) <= 2 Then

                MsgBox("Used part Yes or No and Purpose of return is mandatory, whenever Issue type is Return", vbInformation)
                Exit Sub

            End If



        End If




        If mode = "Edit" And txtSlNo.Text = "" Then
            MsgBox("Click on slno.", vbInformation)
            Exit Sub
        End If

        If txtPartNumber.Text = "" Or Len(txtPartNumber.Text) < 3 Then

            MsgBox("Enter Part Number", vbInformation)
            Exit Sub
        End If

        If ComboBoxIssueType.Text = "Issue" Then
            If Val(txtQty.Text) > Val(LblStockavble.Text) Then

                MsgBox("Quanity required should not be more than stock available ", vbInformation)
                Exit Sub
            End If
        End If


        Dim IR As String
        Dim FN As String


        Dim checkdt As Date
        checkdt = Today

        Dim strsql2 As String
        Dim cmSQL As SqlCommand
        Dim msgb As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)


        msgb = MsgBox("Are you sure of saving ?", vbYesNo)

        If msgb = vbYes Then


            If txtReqNo.Text = "" Then

                If RBFSItem.Checked = True Then
                    FN = "F"
                ElseIf RBNonFSItem.Checked = True Then
                    FN = "N"
                End If


                'generating regno

                transmode = "Add"
                nogenerate()

                txtReqNo.Enabled = True
                txtReqNo.Text = reqno
                txtReqNo.Enabled = False

                cnSQL.Open()

                curdate = System.DateTime.Now()

                strsql2 = "insert TSS_WH_MaterialRequestHeader values (" & txtReqNo.Text & ",'" & DTPReqDt.Value & "', '" & FN & "'," & _
                          "'" & txtMO.Text & "','" & txtNotes.Text & "','" & ComboBoxdept.Text & "','" & ComboBoxSD.Text & "','" & ComboBoxCell.Text & "','" & username & "', '" & curdate & "')"

                cmSQL = New SqlCommand(strsql2, cnSQL)

                If cmSQL.ExecuteNonQuery() = 0 Then
                    MsgBox("Cannot Save the Material Request Header Details. " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                    '  txtRegNo.Text = 0
                    Exit Sub
                End If
                cnSQL.Close()

                '   End If

                'UPDATE REQ NO to control table
                transmode = "Update"
                If transmode = "Update" Then
                    nogenerate()
                End If

            End If

            'save detail section.

            If ComboBoxIssueType.Text = "Issue" Then
                IR = "I"
            ElseIf ComboBoxIssueType.Text = "Return" Then
                IR = "R"
            End If



            cnSQL.Open()

            curdate = System.DateTime.Now()

            strsql2 = "insert TSS_WH_MaterialRequestDetail values (" & txtReqNo.Text & "," & txtSlNo.Text & ", '" & txtPartNumber.Text & "','" & txtDescription.Text & "'," & _
                      "" & txtQty.Text & ",'" & ComboBoxuom.Text & "' ,'" & IR & "','" & ComboBoxUYN.Text & "','" & ComboBoxPur.Text & "','" & txtDetailRemark.Text & "','" & username & "', '" & curdate & "')"

            cmSQL = New SqlCommand(strsql2, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot Save the Line Details. " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                '  txtRegNo.Text = 0
                Exit Sub
            End If
            cnSQL.Close()

            'load the grid

            Dim stockDC As DataSet = New DataSet


            strsql2 = "SELECT Slno, Issue_Ret as 'Issue Type',Part_Number as 'Part Number', Part_Desc as 'Description', Qty, UOM,  UsedPartY_N as 'Used Part', Purpose, Remarks FROM [FSPrograms].[dbo].[TSS_WH_MaterialRequestDetail] WHERE  [MatReq_no] = '" & txtReqNo.Text & "' order by Slno "

            Dim sqlCmd As SqlCommand = New SqlCommand(strsql2, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAC.TableMappings.Add("Table", "Enq")

            stockDAC.Fill(stockDC)

            DataGridViewMaterialReq.DataSource = stockDC.Tables(0)

            DataGridViewMaterialReq.Columns("Slno").ReadOnly = True
            DataGridViewMaterialReq.Columns("Slno").Width = 45

            DataGridViewMaterialReq.Columns("Issue Type").ReadOnly = True
            DataGridViewMaterialReq.Columns("Issue Type").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridViewMaterialReq.Columns("Issue Type").Width = 85

            DataGridViewMaterialReq.Columns("Part Number").ReadOnly = True
            DataGridViewMaterialReq.Columns("Part Number").Width = 150

            DataGridViewMaterialReq.Columns("Description").ReadOnly = True
            DataGridViewMaterialReq.Columns("Description").Width = 200

            DataGridViewMaterialReq.Columns("Qty").ReadOnly = True
            DataGridViewMaterialReq.Columns("Qty").Width = 80
            DataGridViewMaterialReq.Columns("Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridViewMaterialReq.Columns("Qty").DefaultCellStyle.Format = "N2"

            DataGridViewMaterialReq.Columns("UOM").ReadOnly = True
            DataGridViewMaterialReq.Columns("UOM").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridViewMaterialReq.Columns("UOM").Width = 85

            DataGridViewMaterialReq.Columns("Used Part").ReadOnly = True
            DataGridViewMaterialReq.Columns("Used Part").Width = 85

            DataGridViewMaterialReq.Columns("Purpose").ReadOnly = True
            DataGridViewMaterialReq.Columns("Purpose").Width = 125

            DataGridViewMaterialReq.Columns("Remarks").ReadOnly = True
            DataGridViewMaterialReq.Columns("Remarks").Width = 200

            'clear the text boxs
            txtSlNo.Text = ""
            txtPartNumber.Text = ""
            txtDescription.Text = ""
            txtQty.Text = ""
            ComboBoxuom.Text = "      "
            ComboBoxuom.Text = ""
            ComboBoxUYN.Text = ""
            ComboBoxPur.Text = ""
            txtDetailRemark.Text = ""
            lblStk.Text = ""
            LblStockavble.Text = ""
            lblStk.Visible = False
            LblStockavble.Visible = False


            'generate slno
            slnogenerate()

            If txtSlNo.Text > 1 Then
                GroupBoxFS.Enabled = False
            End If
        End If


    End Sub

    Private Sub txtMO_KeyDown(sender As Object, e As KeyEventArgs) Handles txtMO.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtMO_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMO.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If


    End Sub

    Private Sub txtMO_TextChanged(sender As Object, e As EventArgs) Handles txtMO.TextChanged
        If txtSlNo.Text = "" And txtReqNo.Text = "" Then
            MsgBox("Click on New", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub txtNotes_KeyDown(sender As Object, e As KeyEventArgs) Handles txtNotes.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If


        'If e.KeyCode = Keys.Enter Then
        'RBFSItem.Focus()
        'End If
    End Sub

    Private Sub txtNotes_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNotes.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If



    End Sub

    Private Sub txtRemarks_TextChanged(sender As Object, e As EventArgs) Handles txtNotes.TextChanged
        If txtSlNo.Text = "" And txtReqNo.Text = "" Then
            MsgBox("Click on New", vbInformation)
            Exit Sub
        End If
    End Sub

    Private Sub clearall()

        txtReqNo.Enabled = True
        txtReqNo.Text = ""
        txtReqNo.Enabled = False

        txtMO.Text = ""
        txtNotes.Text = ""
        ComboBoxdept.Text = ""
        ComboBoxSD.Text = ""
        ComboBoxCell.Text = ""

        DataGridViewMaterialReq.Columns.Clear()

        GroupBoxFS.Enabled = True


        'clear the text boxs
        txtSlNo.Text = ""
        txtPartNumber.Text = ""
        txtDescription.Text = ""
        txtQty.Text = ""
        ComboBoxUYN.Text = ""
        ComboBoxPur.Text = ""
        txtDetailRemark.Text = ""

        txtSlNo.Text = 1
        btnApproval.Enabled = True

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnApproval.Click


        If Val(txtReqNo.Text) > 0 Then

            'confirmation msg

            '    Dim checkdt As Date
            '   checkdt = Today

            Dim strsql2 As String
            Dim cmSQL As SqlCommand
            Dim msgb As String
            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

            msgb = MsgBox("Are you sure of sending for Approval ?", vbYesNo)

            If msgb = vbYes Then


                '  CheckPendingforApp()

                'update the table             
                curdate = System.DateTime.Now()
                cnSQL.Open()
                strsql2 = "insert TSS_WH_Approvals values ('MatReq', " & txtReqNo.Text & ",'" & username & "', '" & curdate & "',''," & _
                              "'','','01/01/1900','')"

                cmSQL = New SqlCommand(strsql2, cnSQL)

                If cmSQL.ExecuteNonQuery() = 0 Then
                    MsgBox("Not able to send the approval " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                    '  txtRegNo.Text = 0
                    Exit Sub
                End If
                cnSQL.Close()

                'End If




                MailMatReqApproval()
            End If


        End If



    End Sub

    Private Sub MailMatReqApproval()
        'Dim cnn As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim sql As String

        ' Create an Outlook application.
        Dim oApp As Outlook._Application
        oApp = New Outlook.Application()

        ' Create a new MailItem.
        Dim oMsg As Outlook._MailItem

        oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
        '  oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
        oMsg.Subject = "Material Request : Approval Required,  Request No.'" & txtReqNo.Text & "' "

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        'Dim strSQL1 As String
        Dim t As String
        Dim cc As String
        ' Dim name As String


        sql = "Select AprReq_UserName, [1stApr_EmailId] FROM [FSPrograms].[dbo].[TSS_WH_Approval_Master] where Trans_Type = 'MatReq'  and AprReq_UserId = '" & username & "' "

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(sql, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then

            Else

                t = drSQL1.Item(1)
                cc = drSQL1(0) ' NAME OF THE SENDER
                '    cc = cc & ";" & drSQL1(2) & ";" & "indira.shetty@trelleborg.com;Rajeesh.Ambadi@trelleborg.com"


                Dim name As String = t
                name = name.Substring(0, name.Length - 15)

                oMsg.Body = "Dear " & name & "," & vbCrLf & vbCrLf & "Material Request No.  '" & txtReqNo.Text & "' is sent for approval. " & vbCr & " Please approve the same " & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & vbCrLf & " '" & cc & "' "


                '                oMsg.Body = "Dear " & name & "," & vbCrLf & vbCrLf & "Enquiry of customer '" & txtCustomer.Text & "' is registered in Focus Software. " & vbCr & " Registration Number as above. " & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Customer Support Team "


                oMsg.To = t
                '  oMsg.CC = cc


            End If

        End If

        Dim sBodyLen As String = oMsg.Body.Length

        oMsg.Send()

        MsgBox("Approval mail sent to  '" & drSQL1(0) & "'", vbInformation)
        btnApproval.Enabled = False


        ' Clean up
        oApp = Nothing
        oMsg = Nothing
        cnSQL1.Close()

    End Sub

    Private Sub txtReqNo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtReqNo.KeyDown
        'Private Sub txtUser_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUser.KeyDown


        If e.KeyCode = Keys.Enter Then
            DTPReqDt.Focus()
        End If

    End Sub


    Private Sub txtReqNo_TextChanged(sender As Object, e As EventArgs) Handles txtReqNo.TextChanged

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBoxIssueType_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxIssueType.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If


        'If e.KeyCode = Keys.Enter Then
        '    txtPartNumber.Focus()
        'End If
    End Sub

    Private Sub ComboBoxIssueType_LostFocus(sender As Object, e As EventArgs) Handles ComboBoxIssueType.LostFocus

        If ComboBoxIssueType.Text = "Issue" Then
            LabelDetQty.Text = "Req. Qty"
            ComboBoxUYN.Enabled = False
            ComboBoxPur.Text = ""
            ComboBoxPur.Enabled = False

        ElseIf ComboBoxIssueType.Text = "Return" Then
            LabelDetQty.Text = "Return Qty"
            ComboBoxUYN.Enabled = True
            ComboBoxPur.Enabled = True
        End If


    End Sub

    Private Sub ComboBoxIssueType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxIssueType.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxUYN_GotFocus(sender As Object, e As EventArgs) Handles ComboBoxUYN.GotFocus


        'If ComboBoxIssueType.Text = "Issue" Then
        '    LabelDetQty.Text = "Req. Qty"
        '    ComboBoxUYN.Enabled = False
        '    ComboBoxPur.Text = ""
        '    ComboBoxPur.Enabled = False

        'ElseIf ComboBoxIssueType.Text = "Return" Then
        '    LabelDetQty.Text = "Return Qty"
        '    ComboBoxUYN.Enabled = True
        '    ComboBoxPur.Enabled = True
        'End If





    End Sub

    Private Sub ComboBoxUYN_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxUYN.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If



        'If e.KeyCode = Keys.Enter Then
        '    ComboBoxPur.Focus()
        'End If
    End Sub

    Private Sub ComboBoxUYN_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxUYN.SelectedIndexChanged

    End Sub

    Private Sub DataGridViewWHPartNumbers_CurrentCellChanged(sender As Object, e As EventArgs)



    End Sub

    Private Sub DataGridPartNumbers_Navigate(sender As Object, ne As NavigateEventArgs)

    End Sub

    Private Sub DataGridViewWHPartNumbers_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewWHPartNumbers.CellContentClick

        txtPartNumber.Text = DataGridViewWHPartNumbers.CurrentRow.Cells(0).Value.ToString
        txtDescription.Text = DataGridViewWHPartNumbers.CurrentRow.Cells(1).Value.ToString
        ComboBoxuom.Text = DataGridViewWHPartNumbers.CurrentRow.Cells(2).Value.ToString
        txtQty.Focus()

        If RBNonFSItem.Checked = True Then
            ComboBoxIssueType.Text = "Issue"

        End If


        If ComboBoxIssueType.Text = "Issue" Then
            LabelDetQty.Text = "Req. Qty"
            ComboBoxUYN.Enabled = False
            ComboBoxPur.Text = ""
            ComboBoxPur.Enabled = False

        ElseIf ComboBoxIssueType.Text = "Return" Then
            LabelDetQty.Text = "Return Qty"
            ComboBoxUYN.Enabled = True
            ComboBoxPur.Enabled = True
        End If



    End Sub

    Private Sub DataGridViewWHPartNumbers_MouseClick(sender As Object, e As MouseEventArgs) Handles DataGridViewWHPartNumbers.MouseClick

    End Sub

    Private Sub DataGridViewWHPartNumbers_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridViewWHPartNumbers.RowHeaderMouseClick
        txtPartNumber.Text = DataGridViewWHPartNumbers.CurrentRow.Cells(0).Value.ToString
        txtDescription.Text = DataGridViewWHPartNumbers.CurrentRow.Cells(1).Value.ToString
        ComboBoxuom.Text = Trim(DataGridViewWHPartNumbers.CurrentRow.Cells(2).Value.ToString)

        txtQty.Focus()

        DataGridViewWHPartNumbers.Visible = False



    End Sub

    Private Sub DTPReqDt_KeyDown(sender As Object, e As KeyEventArgs) Handles DTPReqDt.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtMO.Focus()
        End If
    End Sub




    Private Sub DTPReqDt_ValueChanged(sender As Object, e As EventArgs) Handles DTPReqDt.ValueChanged

    End Sub

    Private Sub RBFSItem_CheckedChanged(sender As Object, e As EventArgs) Handles RBFSItem.CheckedChanged

    End Sub

    Private Sub RBFSItem_KeyDown(sender As Object, e As KeyEventArgs) Handles RBFSItem.KeyDown
        If e.KeyCode = Keys.Enter Then
            RBNonFSItem.Focus()
        End If
    End Sub

    Private Sub ComboBoxuom_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxuom.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxuom_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxuom.SelectedIndexChanged

    End Sub

    Private Sub txtQty_DoubleClick(sender As Object, e As EventArgs) Handles txtQty.DoubleClick


        lblStk.Visible = True
        LblStockavble.Visible = True

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim drSQL1 As SqlDataReader
        Dim PART As String
        PART = txtPartNumber.Text
        cnSQL1.Open()


        If RBFSItem.Checked = True Then


            Dim cmd As SqlCommand = New SqlCommand("TSS_WH_PartNumberAgainstStock_FSItems", cnSQL1)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@PartNumber", PART)

            drSQL1 = cmd.ExecuteReader()

            If drSQL1.Read() Then

                If IsDBNull(drSQL1.Item(0)) Then

                    LblStockavble.Text = 0
                Else
                    LblStockavble.Text = drSQL1.GetDecimal(0)
                End If

            End If

        ElseIf RBNonFSItem.Checked = True Then


            Dim cmd As SqlCommand = New SqlCommand("TSS_WH_PartNumberAgainstStock", cnSQL1)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@PartNumber", PART)

            drSQL1 = cmd.ExecuteReader()

            If drSQL1.Read() Then

                ' If IsDBNull(drSQL1) Then
                If IsDBNull(drSQL1.Item(0)) Then
                    LblStockavble.Text = 0
                Else
                    LblStockavble.Text = drSQL1.GetDecimal(0)
                End If

            End If
        End If

    End Sub

    Private Sub txtQty_GotFocus(sender As Object, e As EventArgs) Handles txtQty.GotFocus



    End Sub

    Private Sub txtQty_KeyDown(sender As Object, e As KeyEventArgs) Handles txtQty.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If



        If ComboBoxIssueType.Text = "Issue" Then

            If e.KeyCode = Keys.Enter Then
                txtDetailRemark.Focus()
            End If


        ElseIf ComboBoxIssueType.Text = "Return" Then

            If e.KeyCode = Keys.Enter Then
                ComboBoxUYN.Focus()
            End If
        End If


    End Sub

    Private Sub txtQty_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQty.KeyPress

        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)
        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtQty_TextChanged(sender As Object, e As EventArgs) Handles txtQty.TextChanged

    End Sub

    Private Sub RBNonFSItem_CheckedChanged(sender As Object, e As EventArgs) Handles RBNonFSItem.CheckedChanged

    End Sub

    Private Sub txtDescription_KeyDown(sender As Object, e As KeyEventArgs) Handles txtDescription.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If


        'If e.KeyCode = Keys.Enter Then
        '    txtQty.Focus()
        'End If
    End Sub

    Private Sub txtDescription_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDescription.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtDescription_LostFocus(sender As Object, e As EventArgs) Handles txtDescription.LostFocus
        If Len(txtPartNumber.Text) > 3 And Len(Trim(txtDescription.Text)) < 2 Then

            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL As SqlCommand
            Dim drSQL As SqlDataReader
            Dim strSQL As String

            If RBFSItem.Checked = True Then

                strSQL = "SELECT ItemDescription, ItemUM   FROM [FSDBBR].[dbo].[TSS_WH_ToolsCons_Stock_P] " & _
                         "WHERE  ItemNumber like '" & txtPartNumber.Text & "' "

            ElseIf RBNonFSItem.Checked = True Then

                strSQL = "SELECT Part_Description,UOM FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] with (nolock) " & _
                         "WHERE  ItemNumber like '" & txtPartNumber.Text & "' "

            End If

            cnSQL.Open()

            cmSQL = New SqlCommand(strSQL, cnSQL)
            drSQL = cmSQL.ExecuteReader()

            If drSQL.Read() Then
                txtDescription.Text = drSQL.Item(0)
                ComboBoxuom.Text = drSQL.Item(1)

            Else
                MsgBox("This part number not existing", vbInformation)
                Exit Sub
            End If
            cnSQL.Close()

        End If

    End Sub

    Private Sub txtDescription_TextChanged(sender As Object, e As EventArgs) Handles txtDescription.TextChanged

    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click

        mode = "Edit"

        clearall()

        btnSaveChanges.Enabled = True

        DataGridViewMatReqEdit.Visible = True
        DataGridViewMatReqEdit.Enabled = True
        DataGridViewMatReqEdit.EditMode = DataGridViewEditMode.EditOnEnter
        DataGridViewMatReqEdit.BringToFront()

        MatReqList()





        '  MsgBox("Function not available", vbInformation)
        ' Exit Sub
    End Sub

    Private Sub ComboBoxdept_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxdept.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If

        'If e.KeyCode = Keys.Enter Then
        '    ComboBoxSD.Focus()

        'End If
    End Sub

    Private Sub ComboBoxdept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxdept.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxSD_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxSD.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
        'If e.KeyCode = Keys.Enter Then
        '    ComboBoxCell.Focus()
        'End If
    End Sub

    Private Sub ComboBoxSD_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxSD.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxCell_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxCell.KeyDown

        'If e.KeyCode = Keys.Enter Then
        'Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        'End If

        If e.KeyCode = Keys.Enter Then
            ComboBoxIssueType.Focus()
        End If

        'End Sub


    End Sub

    Private Sub txtSlNo_DoubleClick(sender As Object, e As EventArgs) Handles txtSlNo.DoubleClick

        If mode = "Edit" Then
            ' take next sl.no.

            slnogenerate()


        End If

    End Sub

    Private Sub txtSlNo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSlNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtSlNo_TextChanged(sender As Object, e As EventArgs) Handles txtSlNo.TextChanged

    End Sub

    Private Sub ComboBoxPur_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxPur.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If



        'If e.KeyCode = Keys.Enter Then
        '    txtDetailRemark.Focus()
        'End If
    End Sub

    Private Sub txtDetailRemark_KeyDown(sender As Object, e As KeyEventArgs) Handles txtDetailRemark.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If



        'If e.KeyCode = Keys.Enter Then
        '    btnReqSave.Focus()
        'End If
    End Sub

    Private Sub txtDetailRemark_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDetailRemark.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtDetailRemark_TextChanged(sender As Object, e As EventArgs) Handles txtDetailRemark.TextChanged

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        Dim msgb As String
        Dim strsql As String
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim strsql2 As String
        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL2 As SqlCommand

        If Len(txtMO.Text) > 0 Then


            msgb = MsgBox("Are you sure of deleting ?", vbYesNo)

            If msgb = vbYes Then

                'check in approval table
                strsql = "SELECT Trans_No  FROM [FSPrograms].[dbo].[TSS_WH_Approvals] WHERE  Trans_Type = 'MatReq' and  Trans_No = " & txtReqNo.Text & "" '  and len([1st_AppStatus]) >= 1 "


                cnSQL.Open()

                cmSQL = New SqlCommand(strsql, cnSQL)
                drSQL = cmSQL.ExecuteReader()

                If drSQL.Read() Then
                    MsgBox("This Request is already referred in approvals, you can't delete it", vbInformation)
                    btnSaveChanges.Enabled = False
                    btnReqSave.Enabled = False
                    btnApproval.Enabled = False
                    Exit Sub


                Else
                    'delete sql for header
                    strsql2 = "delete  FROM [FSPrograms].[dbo].[TSS_WH_MaterialRequestHeader] where [MatReq_no] = " & txtReqNo.Text & ""
                    cnSQL2.Open()
                    cmSQL2 = New SqlCommand(strsql2, cnSQL2)

                    If cmSQL2.ExecuteNonQuery() = 0 Then

                        MsgBox("Error while deleting the Material request header section", vbInformation)
                        Exit Sub
                    End If
                    cnSQL2.Close()

                    'delete for detail

                    strsql2 = "delete  FROM [FSPrograms].[dbo].[TSS_WH_MaterialRequestDetail] where [MatReq_no] = " & txtReqNo.Text & ""
                    cnSQL2.Open()
                    cmSQL2 = New SqlCommand(strsql2, cnSQL2)

                    If cmSQL2.ExecuteNonQuery() = 0 Then

                        MsgBox("Error while deleting Material request detail section", vbInformation)
                        Exit Sub
                    End If
                    cnSQL2.Close()
                    MsgBox("Material Request deleted", vbInformation)
                    Exit Sub


                End If
                cnSQL.Close()
            End If
        Else

            MsgBox("Please select Material Request Number through Edit mode", vbInformation)
            Exit Sub
        End If



    End Sub

    Private Sub MatReqList()
        '  DataGridViewMatReqEdit.Visible = True
        ' DataGridViewMatReqEdit.Show()



        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim strSql As String
        Dim stockDC As DataSet = New DataSet


        strSql = "SELECT MatReq_no, MatReq_Date ,Type_Dept as Dept, Sub_Div,Cell, FS_NonFS, MONumber, User_Id, Remarks FROM FSPrograms.dbo.TSS_WH_MaterialRequestHeader WHERE  User_Id = '" & username & "' order by MatReq_no"


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        sqlCon.Open()

        stockDAC.TableMappings.Add("Table", "Part")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewMatReqEdit.Location = New System.Drawing.Point(450, 17)
        DataGridViewMatReqEdit.Width = 605 '1000
        DataGridViewMatReqEdit.Height = 121 ' 412

        DataGridViewMatReqEdit.BringToFront()



        DataGridViewMatReqEdit.DataSource = stockDC.Tables(0)
        sqlCon.Close()
        '  DataGridViewWHPartNumbers.Expand(-1)

        DataGridViewMatReqEdit.Columns("MatReq_no").Width = 100
        DataGridViewMatReqEdit.Columns("MatReq_Date").Width = 150
        DataGridViewMatReqEdit.Columns("Dept").Width = 150
        DataGridViewMatReqEdit.Columns("Sub_Div").Width = 150
        DataGridViewMatReqEdit.Columns("Cell").Width = 150
        DataGridViewMatReqEdit.Columns("User_Id").Width = 150
        DataGridViewMatReqEdit.Focus()

    End Sub

    Private Sub DataGridViewMatReqEdit_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewMatReqEdit.CellContentClick



    End Sub

    Private Sub DataGridViewMatReqEdit_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewMatReqEdit.RowHeaderMouseClick
        Dim strsql2 As String
        '  Dim cmSQL As SqlCommand
        ' Dim msgb As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)



        txtReqNo.Text = DataGridViewMatReqEdit.CurrentRow.Cells(0).Value.ToString
        DTPReqDt.Value = DataGridViewMatReqEdit.CurrentRow.Cells(1).Value.ToString
        ComboBoxdept.Text = DataGridViewMatReqEdit.CurrentRow.Cells(2).Value.ToString
        ComboBoxSD.Text = DataGridViewMatReqEdit.CurrentRow.Cells(3).Value.ToString
        ComboBoxCell.Text = DataGridViewMatReqEdit.CurrentRow.Cells(4).Value.ToString

        If DataGridViewMatReqEdit.CurrentRow.Cells(5).Value.ToString = "F" Then
            RBFSItem.Checked = True
        ElseIf DataGridViewMatReqEdit.CurrentRow.Cells(5).Value.ToString = "N" Then
            RBNonFSItem.Checked = True
        End If
        txtMO.Text = DataGridViewMatReqEdit.CurrentRow.Cells(6).Value.ToString
        txtNotes.Text = DataGridViewMatReqEdit.CurrentRow.Cells(8).Value.ToString


        'load all items

        DataGridViewMatReqEdit.Visible = False


        Dim stockDC As DataSet = New DataSet


        strsql2 = "SELECT Slno,Issue_Ret AS 'Issue Type',Part_Number AS  'Part Number',Part_Desc AS 'Description' ,Qty ,UOM , UsedPartY_N as 'Used Part', Purpose ,DetailRemarks as 'Remarks' ,1st_AppStatus ,Reason_Rej " & _
             " FROM [FSPrograms].[dbo].[TSS_WH_MatRequest_List] WHERE   MatReq_no= '" & txtReqNo.Text & "' order by Slno"


        Dim sqlCmd As SqlCommand = New SqlCommand(strsql2, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")

        stockDAC.Fill(stockDC)

        DataGridViewMaterialReq.DataSource = stockDC.Tables(0)

        DataGridViewMaterialReq.Columns("Slno").ReadOnly = True
        DataGridViewMaterialReq.Columns("Slno").Width = 45

        DataGridViewMaterialReq.Columns("Issue Type").ReadOnly = True
        DataGridViewMaterialReq.Columns("Issue Type").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridViewMaterialReq.Columns("Issue Type").Width = 55


        DataGridViewMaterialReq.Columns("Part Number").ReadOnly = True
        DataGridViewMaterialReq.Columns("Part Number").Width = 150

        DataGridViewMaterialReq.Columns("Description").ReadOnly = True
        DataGridViewMaterialReq.Columns("Description").Width = 200

        DataGridViewMaterialReq.Columns("Qty").ReadOnly = False
        DataGridViewMaterialReq.Columns("Qty").Width = 80
        DataGridViewMaterialReq.Columns("Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewMaterialReq.Columns("Qty").DefaultCellStyle.Format = "N2"
        DataGridViewMaterialReq.Columns("Qty").HeaderCell.Style.BackColor = Color.Gray

        DataGridViewMaterialReq.Columns("UOM").ReadOnly = True
        DataGridViewMaterialReq.Columns("UOM").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridViewMaterialReq.Columns("UOM").Width = 55


        DataGridViewMaterialReq.Columns("Used Part").ReadOnly = True
        DataGridViewMaterialReq.Columns("Used Part").Width = 55

        DataGridViewMaterialReq.Columns("Purpose").ReadOnly = True
        DataGridViewMaterialReq.Columns("Purpose").Width = 125

        DataGridViewMaterialReq.Columns("Remarks").ReadOnly = False
        DataGridViewMaterialReq.Columns("Remarks").Width = 250
        DataGridViewMaterialReq.Columns("Remarks").HeaderCell.Style.BackColor = Color.Gray

        DataGridViewMaterialReq.Columns("st_AppStatus").ReadOnly = True
        DataGridViewMaterialReq.Columns("st_AppStatus").Width = 0
        DataGridViewMaterialReq.Columns("st_AppStatus").Visible = False

        DataGridViewMaterialReq.Columns("Reason_Rej").ReadOnly = True
        DataGridViewMaterialReq.Columns("Reason_Rej").Width = 0
        DataGridViewMaterialReq.Columns("Reason_Rej").Visible = False


        cnSQL.Close()

        'check is it approved.

        'Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String



        strSQL = "SELECT Trans_No  FROM [FSPrograms].[dbo].[TSS_WH_Approvals] WHERE  Trans_Type = 'MatReq' and  Trans_No = " & txtReqNo.Text & "  and len([1st_AppStatus]) >= 1 "


        cnSQL.Open()

        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()

        If drSQL.Read() Then
            MsgBox("This Request is already approved, you can't make any changes", vbInformation)
            btnSaveChanges.Enabled = False
            btnReqSave.Enabled = False
            btnApproval.Enabled = False
            Exit Sub

        End If
        cnSQL.Close()
        'end of checking



    End Sub

    Private Sub DataGridViewMaterialReq_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewMaterialReq.CellContentClick

    End Sub

    Private Sub btnSaveChanges_Click(sender As Object, e As EventArgs) Handles btnSaveChanges.Click

        Dim cnSQL0 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL0 As SqlCommand
        '   Dim drSQL1 As SqlDataReader
        Dim strSQL0 As String
        cnSQL0.Open()

        'header section update

        strSQL0 = "UPDATE [FSPrograms].[dbo].[TSS_WH_MaterialRequestHeader] set MONumber = '" & txtMO.Text & "', Remarks = '" & txtNotes.Text & "', Type_Dept = '" & ComboBoxdept.Text & "', Sub_Div = '" & ComboBoxSD.Text & "',Cell = '" & ComboBoxCell.Text & "',User_Id = '" & username & "', Datetime = '" & curdate & "' where MatReq_no = " & txtReqNo.Text & ""

        cmSQL0 = New SqlCommand(strSQL0, cnSQL0)

        If cmSQL0.ExecuteNonQuery() = 0 Then
            MsgBox("Cannot Save the Header Details. " & strSQL0, MsgBoxStyle.Exclamation, "Error!")
            Exit Sub
        End If
        cnSQL0.Close()

        'end of header update

        'stock checking before saving
        Dim drSQL01 As SqlDataReader
        Dim stock As Integer
        Dim cnSQL01 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim PART As String

        'cnSQL2.Open()

        For i As Integer = 0 To DataGridViewMaterialReq.RowCount - 1
            cnSQL01.Open()

            If RBFSItem.Checked = True And DataGridViewMaterialReq.CurrentRow.Cells(1).Value.ToString = "I" Then

                PART = DataGridViewMaterialReq.CurrentRow.Cells(2).Value.ToString

                Dim cmd As SqlCommand = New SqlCommand("TSS_WH_PartNumberAgainstStock_FSItems", cnSQL01)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@PartNumber", PART)

                drSQL01 = cmd.ExecuteReader()

            ElseIf RBNonFSItem.Checked = True And DataGridViewMaterialReq.CurrentRow.Cells(1).Value.ToString = "R" Then

                PART = DataGridViewMaterialReq.CurrentRow.Cells(2).Value.ToString

                Dim cmd As SqlCommand = New SqlCommand("TSS_WH_PartNumberAgainstStock", cnSQL01)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@PartNumber", PART)

                drSQL01 = cmd.ExecuteReader()

            End If

            If drSQL01.Read() Then

                If IsDBNull(drSQL01.Item(0)) Then

                    stock = 0
                Else
                    stock = drSQL01.GetDecimal(0)
                End If

                If Val(stock) < Val(DataGridViewMaterialReq.CurrentRow.Cells(4).Value) Then
                    MsgBox("Quantity entered should not be greater than stock available : " & PART, vbInformation)
                    lblStk.Visible = True
                    LblStockavble.Visible = True
                    LblStockavble.Text = stock
                    cnSQL01.Close()
                    Exit Sub
                End If

            End If

            cnSQL01.Close()

        Next

        'saving detail section

        cnSQL0.Open()
        For i As Integer = 0 To DataGridViewMaterialReq.RowCount - 1

            strSQL0 = "update [FSPrograms].[dbo].[TSS_WH_MaterialRequestDetail] set Qty = " & Me.DataGridViewMaterialReq.Rows(i).Cells("Qty").Value & " , Remarks  = '" & Me.DataGridViewMaterialReq.Rows(i).Cells("Remarks").Value & "', User_Id = '" & username & "', [DateTime] = '" & curdate & "' where MatReq_no = " & txtReqNo.Text & " and Slno = " & Me.DataGridViewMaterialReq.Rows(i).Cells("Slno").Value & " "

            cmSQL0 = New SqlCommand(strSQL0, cnSQL0)

            If cmSQL0.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot Save the Details. " & strSQL0, MsgBoxStyle.Exclamation, "Error!")
                'txtRegNo.Text = 0
                'Application.Exit()
                Exit Sub
            End If
        Next

        cnSQL0.Close()

    End Sub

    Private Sub ComboBoxPur_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPur.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxCell_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxCell.SelectedIndexChanged

    End Sub
End Class