Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
'Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection


Public Class WHItemMaster
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Private Sub MyGroupBox2_Enter(sender As Object, e As EventArgs) Handles MyGroupBox2.Enter

    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim strsql2 As String
        Dim cmSQL As SqlCommand
        Dim msgb As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)


        msgb = MsgBox("Are you sure of saving ?", vbYesNo)

        If msgb = vbYes Then

            'check part number duplicacy

            Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL3 As SqlCommand
            Dim drSQL3 As SqlDataReader
            Dim strSQL3 As String
            Dim loc As String

            If transmode = "Add" Then
                strSQL3 = "SELECT Part_Number FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] WHERE [Part_Number] = '" & txtPartNumber.Text & "'"

                cnSQL3.Open()
                cmSQL3 = New SqlCommand(strSQL3, cnSQL3)
                drSQL3 = cmSQL3.ExecuteReader()

                If drSQL3.Read() Then

                    If Len(drSQL3.Item(0)) > 0 Then

                        MsgBox("This Part Number is already existing", vbInformation)
                        Exit Sub
                    Else
                    End If

                End If
                cnSQL3.Close()
            End If

            'end of checking part number duplicacy

            'stock room checking



            loc = LTrim(RTrim(txtSTR.Text)) + LTrim(RTrim(txtBin.Text))

            If Len(loc) > 3 Then

                ' strSQL3 = "SELECT ltrim(rtrim([Stockroom]))+ ltrim(rtrim([Bin])) FROM [FSDBBR].[dbo].[_NoLock_FS_InventoryLocation] union SELECT ltrim(rtrim([StockRoom]))+ ltrim(rtrim([BIN])) FROM [FSPrograms].[dbo].[TSS_WH_StockRooms] where  ltrim(rtrim([Stockroom]))+ ltrim(rtrim([Bin])) = '" & loc & "' "

                strSQL3 = "select * from (((SELECT ltrim(rtrim([Stockroom]))+ ltrim(rtrim([Bin])) as loc FROM [FSDBBR].[dbo].[_NoLock_FS_InventoryLocation] union SELECT ltrim(rtrim([StockRoom]))+ ltrim(rtrim([BIN])) as loc FROM [FSPrograms].[dbo].[TSS_WH_StockRooms]))) a   where loc = '" & loc & "' "

                cnSQL3.Open()
                cmSQL3 = New SqlCommand(strSQL3, cnSQL3)
                drSQL3 = cmSQL3.ExecuteReader()


                If drSQL3.Read() Then

                    If IsDBNull(drSQL3.Item(0)) Then

                        MsgBox("Stock room / location is not existing ", vbInformation)
                        Exit Sub
                    Else
                    End If

                Else
                    MsgBox("Stock room / location is not existing ", vbInformation)
                    Exit Sub
                End If

                cnSQL3.Close()
            End If
        End If

        'end of stock room and bin checking

        cnSQL.Open()

        curdate = System.DateTime.Now()



        ComboBoxStatus.Text = LTrim(RTrim(ComboBoxStatus.Text))
        txtSTR.Text = LTrim(RTrim(txtSTR.Text))
        txtBin.Text = LTrim(RTrim(txtBin.Text))
        ComboBoxUOM.Text = LTrim(RTrim(ComboBoxUOM.Text))

        If Val(txtSafetystk.Text) > 0 Then
        Else
            txtSafetystk.Text = 0
        End If

        If Val(txtLeadTime.Text) > 0 Then
        Else
            txtLeadTime.Text = 0
        End If


        If transmode = "Add" Then


            strsql2 = "insert TSS_WH_ItemMaster values ('" & ComboBoxItemType.Text & "','" & txtPartNumber.Text & "', '" & txtPartDesc.Text & "'," & _
                                  "'" & ComboBoxUOM.Text & "','" & ComboBoxStatus.Text & "'," & txtSafetystk.Text & ",'" & ComboBoxItemClass.Text & "'," & txtLeadTime.Text & "," & _
                                  "'" & txtSTR.Text & "', '" & txtBin.Text & "', '" & txtUDF1.Text & "', '" & txtUDF2.Text & "','" & txtUDF3.Text & "','" & txtUDF4.Text & "','" & txtUDF5.Text & "','" & txtUDF6.Text & "','" & txtUDF7.Text & "'," & _
                                  "'" & txtUDF8.Text & "', '" & txtUDF9.Text & "','" & txtUDF10.Text & "','" & username & "', '" & curdate & "')"

            cmSQL = New SqlCommand(strsql2, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot Save the Item Master Details. " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                '  txtRegNo.Text = 0
                Exit Sub
            End If

        ElseIf transmode = "Edit" Then



            strsql2 = "update [FSPrograms].[dbo].[TSS_WH_ItemMaster] set " & _
                      "Part_Description = '" & txtPartDesc.Text & "', UOM = '" & ComboBoxUOM.Text & "', Status = '" & ComboBoxStatus.Text & "', Safety_Stock = " & txtSafetystk.Text & "," & _
                      "ItemClass = '" & ComboBoxItemClass.Text & "', Lead_Time = " & txtLeadTime.Text & ", Pre_StockRoom = '" & txtSTR.Text & "',Pre_BIN = '" & txtBin.Text & "'," & _
                      "UDF1 = '" & txtUDF1.Text & "',UDF2 = '" & txtUDF2.Text & "',UDF3 = '" & txtUDF3.Text & "',UDF4 = '" & txtUDF4.Text & "',UDF5 = '" & txtUDF5.Text & "'," & _
                      "UDF6 = '" & txtUDF6.Text & "',UDF7 = '" & txtUDF7.Text & "',UDF8 = '" & txtUDF8.Text & "',UDF9 = '" & txtUDF9.Text & "',UDF10 = '" & txtUDF10.Text & "',User_Id = '" & username & "',DateTime = '" & curdate & "' WHERE Part_Number = '" & txtPartNumber.Text & "'"


            cmSQL = New SqlCommand(strsql2, cnSQL)

            If cmSQL.ExecuteNonQuery() = 0 Then
                MsgBox("Cannot Save the Item Master Details. " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                '  txtRegNo.Text = 0
                Exit Sub
            End If


        End If

        cnSQL.Close()

        MsgBox("Entered Part Number Saved Successfully", vbInformation)
        btnSave.Enabled = False
        clearall()
        Exit Sub


    End Sub
    Private Sub clearall()
        ComboBoxItemType.Text = ""
        txtPartNumber.Text = ""
        txtPartNumber.Enabled = True
        txtPartDesc.Text = ""
        ComboBoxStatus.Text = ""
        ComboBoxUOM.Text = ""
        ComboBoxItemClass.Text = ""
        txtSafetystk.Text = ""
        txtLeadTime.Text = ""
        txtSTR.Text = ""
        txtBin.Text = ""

        txtUDF1.Text = ""
        txtUDF2.Text = ""
        txtUDF3.Text = ""
        txtUDF4.Text = ""
        txtUDF5.Text = ""
        txtUDF6.Text = ""
        txtUDF7.Text = ""
        txtUDF8.Text = ""
        txtUDF9.Text = ""
        txtUDF10.Text = ""
        transmode = ""

    End Sub

    Private Sub txtPartNumber_DoubleClick(sender As Object, e As EventArgs) Handles txtPartNumber.DoubleClick

        partlist()



    End Sub

    Private Sub partlist()
        DataGridItem.Visible = True
        DataGridItem.BringToFront()
        ' datagridReqPending.Location.X.MaxValue = 574

        DataGridItem.Location = New System.Drawing.Point(506, 20)  '516 25

        DataGridItem.Width = 525  '535
        DataGridItem.Height = 366 '276

        DataGridItem.Enabled = True


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String

        Dim stockDC As DataSet = New DataSet


        strSQL = "Select Part_Number,Part_Description,UOM  FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] with (nolock) where [Item_Type] = '" & ComboBoxItemType.Text & "'"

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridItem.DataSource = stockDC.Tables(0)


        DataGridItem.Columns("Part_Number").ReadOnly = True
        DataGridItem.Columns("Part_Number").Width = 175

        DataGridItem.Columns("Part_Description").ReadOnly = True
        DataGridItem.Columns("Part_Description").Width = 240

        DataGridItem.Columns("UOM").ReadOnly = True
        DataGridItem.Columns("UOM").Width = 85

        cnSQL.Close()
    End Sub

    Private Sub txtPartNumber_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPartNumber.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtPartNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPartNumber.KeyPress
        '   Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-_<>(){}[]" & Chr(Keys.Back) ' & Chr(Keys.Space)"

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtPartNumber_TextChanged(sender As Object, e As EventArgs) Handles txtPartNumber.TextChanged

    End Sub

    Private Sub txtPartDesc_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPartDesc.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtPartDesc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPartDesc.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtPartDesc_TextChanged(sender As Object, e As EventArgs) Handles txtPartDesc.TextChanged

    End Sub

    Private Sub txtSTR_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSTR.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtSTR_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSTR.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtSTR_TextChanged(sender As Object, e As EventArgs) Handles txtSTR.TextChanged

    End Sub

    Private Sub txtBin_KeyDown(sender As Object, e As KeyEventArgs) Handles txtBin.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtBin_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtBin.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtBin_TextChanged(sender As Object, e As EventArgs) Handles txtBin.TextChanged

    End Sub

    Private Sub txtUDF1_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF1.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF1_TextChanged(sender As Object, e As EventArgs) Handles txtUDF1.TextChanged

    End Sub

    Private Sub txtUDF2_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF2.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF3_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF3.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF4_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF4.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF5_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF5.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF5.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF6_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF6.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF6.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF7_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF7.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF7.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF7_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles txtUDF7.PreviewKeyDown

    End Sub

    Private Sub txtUDF8_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF8.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF8.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF9_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF9.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF9.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub txtUDF10_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUDF10.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtUDF10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUDF10.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If

    End Sub

    Private Sub ComboBoxItemType_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxItemType.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxItemType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxItemType.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxStatus_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxStatus.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxStatus.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxUOM_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxUOM.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxUOM_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxUOM.SelectedIndexChanged

    End Sub

    Private Sub txtSafetystk_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSafetystk.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtSafetystk_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSafetystk.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)
        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtSafetystk_TextChanged(sender As Object, e As EventArgs) Handles txtSafetystk.TextChanged

    End Sub

    Private Sub txtLeadTime_KeyDown(sender As Object, e As KeyEventArgs) Handles txtLeadTime.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtLeadTime_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtLeadTime.KeyPress
        Dim allowedChars As String = "0123456789" & Chr(Keys.Back)
        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtLeadTime_TextChanged(sender As Object, e As EventArgs) Handles txtLeadTime.TextChanged

    End Sub

    Private Sub ComboBoxItemClass_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxItemClass.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub ComboBoxItemClass_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxItemClass.SelectedIndexChanged

    End Sub

    Private Sub WHItemMaster_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtPartNumber.Enabled = False

        comboloaditemMaster()




    End Sub

    Private Sub comboloaditemMaster()

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand
        'item type load

        'dept_type load

        strSql = "SELECT TypeKey,Item_Type FROM [TSS_WH_ItemTypes] "
        cmSQL = New SqlCommand(strSql, sqlCon)
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource As SqlDataAdapter = New SqlDataAdapter
        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eitem")
        With ComboBoxItemType
            .DataSource = source.Tables("eitem")
            .DisplayMember = "Item_Type"
            .ValueMember = "TypeKey"
            .SelectedIndex = 0
        End With

        'uom load

        'UOM
        strSql = "SELECT ID, UOM FROM [TSS_WH_UOM] "
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet

        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eitem1")
        With ComboBoxUOM
            .DataSource = source.Tables("eitem1")
            .DisplayMember = "UOM"
            .ValueMember = "ID"
            .SelectedIndex = 0
        End With

        'Itemclass


        strSql = "SELECT ID, Item_Class FROM TSS_WH_ItemClass "
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet

        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "esd")

        With ComboBoxItemClass
            .DataSource = source.Tables("esd")
            .DisplayMember = "Item_Class"
            .ValueMember = "ID"
            .SelectedIndex = 0
        End With

        'status

        strSql = "SELECT ID, Status FROM TSS_WH_ItemStatus "
        cmSQL = New SqlCommand(strSql, sqlCon)

        sqlCmd = New SqlCommand(strSql, sqlCon)
        ESource = New SqlDataAdapter
        source = New DataSet

        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "esd1")

        With ComboBoxStatus
            .DataSource = source.Tables("esd1")
            .DisplayMember = "Status"
            .ValueMember = "ID"
            .SelectedIndex = 0
        End With

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        clearall()

        ComboBoxItemType.Enabled = True
        ComboBoxItemType.Focus()

        transmode = "Add"
    End Sub

    Private Sub DataGridItem_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridItem_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)

        'DataGridItem.Visible = False


        'Dim msgb As String
        'Dim cnSQL8 As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cmSQL8 As SqlCommand
        'Dim drSQL8 As SqlDataReader
        'Dim strSQL8 As String

        'txtPartNumber.Text = DataGridItem.CurrentRow.Cells(0).Value
        'txtPartDesc.Text = DataGridItem.CurrentRow.Cells(1).Value
        'ComboBoxUOM.Text = DataGridItem.CurrentRow.Cells(2).Value
        'txtPartNumber.Enabled = False


        'If transmode = "Delete" Then
        '    msgb = MsgBox("Are you sure of deleting ?", vbYesNo)


        '    If msgb = vbYes Then

        '        'check transactions and delete
        '        'check receipts table

        '        strSQL8 = "SELECT PartNumber  FROM [FSPrograms].[dbo].[TSS_WH_Material_Receipt] with (NOLOCK) WHERE PartNumber = '" & txtPartNumber.Text & "'"

        '        cnSQL8.Open()
        '        cmSQL8 = New SqlCommand(strSQL8, cnSQL8)
        '        drSQL8 = cmSQL8.ExecuteReader()

        '        If drSQL8.Read() Then

        '            MsgBox("This Part Number is already referred in Receipts,Deletion is not allowed.", vbInformation)
        '            Exit Sub
        '        End If
        '        cnSQL8.Close()
        '        'check issues

        '        strSQL8 = "SELECT Part_Number  FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueDetail] with (NOLOCK) WHERE Part_Number = '" & txtPartNumber.Text & "'"


        '        cnSQL8.Open()
        '        cmSQL8 = New SqlCommand(strSQL8, cnSQL8)
        '        drSQL8 = cmSQL8.ExecuteReader()

        '        If drSQL8.Read() Then

        '            MsgBox("This Part Number is already referred in Issues,Deletion is not allowed.", vbInformation)
        '            Exit Sub
        '        End If
        '        cnSQL8.Close()

        '        'end of checking and deleting
        '        strSQL8 = "delete  FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] where Part_Number = '" & txtPartNumber.Text & "'"
        '        cnSQL8.Open()
        '        cmSQL8 = New SqlCommand(strSQL8, cnSQL8)

        '        If cmSQL8.ExecuteNonQuery() = 0 Then

        '            MsgBox("Error while deleting the PartNumber", vbInformation)
        '            Exit Sub

        '        Else
        '            MsgBox("Selected part number is deleted", vbInformation)
        '            Exit Sub

        '        End If
        '        cnSQL8.Close()
        '        'end of delete sql

        '    Else
        '        Exit Sub
        '    End If

        'End If

        ''load item details




        'strSQL8 = "SELECT Status,Safety_Stock,ItemClass,Lead_Time,Pre_StockRoom,Pre_BIN,UDF1,UDF2,UDF3,UDF4,UDF5,UDF6,UDF7,UDF8,UDF9,UDF10 FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] with (nolock) where Part_Number = '" & txtPartNumber.Text & "'"

        'cnSQL8.Open()
        'cmSQL8 = New SqlCommand(strSQL8, cnSQL8)
        'drSQL8 = cmSQL8.ExecuteReader()

        'If drSQL8.Read() Then

        '    ComboBoxStatus.Text = drSQL8.Item(0)
        '    txtSafetystk.Text = drSQL8.Item(1)
        '    ComboBoxItemClass.Text = drSQL8.Item(2)
        '    txtLeadTime.Text = drSQL8.Item(3)
        '    txtSTR.Text = drSQL8.Item(4)
        '    txtBin.Text = drSQL8.Item(5)
        '    txtUDF1.Text = drSQL8.Item(6)
        '    txtUDF2.Text = drSQL8.Item(7)
        '    txtUDF3.Text = drSQL8.Item(8)
        '    txtUDF4.Text = drSQL8.Item(9)
        '    txtUDF5.Text = drSQL8.Item(10)
        '    txtUDF6.Text = drSQL8.Item(11)
        '    txtUDF7.Text = drSQL8.Item(12)
        '    txtUDF8.Text = drSQL8.Item(13)
        '    txtUDF9.Text = drSQL8.Item(14)
        '    txtUDF10.Text = drSQL8.Item(15)

        'End If

        'cnSQL8.Close()

        ''end of loading item details

    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        clearall()
        transmode = "Edit"
        btnSave.Enabled = True
        MsgBox("Select the Item Type and click on Part Number ", vbInformation)
        ComboBoxItemType.Focus()
        Exit Sub
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        clearall()
        transmode = "Delete"
        MsgBox("Select the Item Type and click on Part Number ", vbInformation)
        ComboBoxItemType.Focus()
        Exit Sub
    End Sub

    Private Sub DataGridItem_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridItem.CellContentClick

    End Sub

    Private Sub DataGridItem_RowHeaderMouseClick1(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridItem.RowHeaderMouseClick

        DataGridItem.Visible = False


        Dim msgb As String
        Dim cnSQL8 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL8 As SqlCommand
        Dim drSQL8 As SqlDataReader
        Dim strSQL8 As String

        txtPartNumber.Text = DataGridItem.CurrentRow.Cells(0).Value
        txtPartDesc.Text = DataGridItem.CurrentRow.Cells(1).Value
        ComboBoxUOM.Text = DataGridItem.CurrentRow.Cells(2).Value
        txtPartNumber.Enabled = False


        If transmode = "Delete" Then
            msgb = MsgBox("Are you sure of deleting ?", vbYesNo)


            If msgb = vbYes Then

                'check transactions and delete
                'check receipts table

                strSQL8 = "SELECT PartNumber  FROM [FSPrograms].[dbo].[TSS_WH_Material_Receipt] with (NOLOCK) WHERE PartNumber = '" & txtPartNumber.Text & "'"

                cnSQL8.Open()
                cmSQL8 = New SqlCommand(strSQL8, cnSQL8)
                drSQL8 = cmSQL8.ExecuteReader()

                If drSQL8.Read() Then

                    MsgBox("This Part Number is already referred in Receipts,Deletion is not allowed.", vbInformation)
                    Exit Sub
                End If
                cnSQL8.Close()
                'check issues

                strSQL8 = "SELECT Part_Number  FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueDetail] with (NOLOCK) WHERE Part_Number = '" & txtPartNumber.Text & "'"


                cnSQL8.Open()
                cmSQL8 = New SqlCommand(strSQL8, cnSQL8)
                drSQL8 = cmSQL8.ExecuteReader()

                If drSQL8.Read() Then

                    MsgBox("This Part Number is already referred in Issues,Deletion is not allowed.", vbInformation)
                    Exit Sub
                End If
                cnSQL8.Close()

                'end of checking and deleting
                strSQL8 = "delete  FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] where Part_Number = '" & txtPartNumber.Text & "'"
                cnSQL8.Open()
                cmSQL8 = New SqlCommand(strSQL8, cnSQL8)

                If cmSQL8.ExecuteNonQuery() = 0 Then

                    MsgBox("Error while deleting the PartNumber", vbInformation)
                    Exit Sub

                Else
                    MsgBox("Selected part number is deleted", vbInformation)
                    Exit Sub

                End If
                cnSQL8.Close()
                'end of delete sql

            Else
                Exit Sub
            End If

        End If

        'load item details




        strSQL8 = "SELECT Status,Safety_Stock,ItemClass,Lead_Time,Pre_StockRoom,Pre_BIN,UDF1,UDF2,UDF3,UDF4,UDF5,UDF6,UDF7,UDF8,UDF9,UDF10 FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] with (nolock) where Part_Number = '" & txtPartNumber.Text & "'"

        cnSQL8.Open()
        cmSQL8 = New SqlCommand(strSQL8, cnSQL8)
        drSQL8 = cmSQL8.ExecuteReader()

        If drSQL8.Read() Then

            ComboBoxStatus.Text = drSQL8.Item(0)
            txtSafetystk.Text = drSQL8.Item(1)
            ComboBoxItemClass.Text = drSQL8.Item(2)
            txtLeadTime.Text = drSQL8.Item(3)
            txtSTR.Text = drSQL8.Item(4)
            txtBin.Text = drSQL8.Item(5)
            txtUDF1.Text = drSQL8.Item(6)
            txtUDF2.Text = drSQL8.Item(7)
            txtUDF3.Text = drSQL8.Item(8)
            txtUDF4.Text = drSQL8.Item(9)
            txtUDF5.Text = drSQL8.Item(10)
            txtUDF6.Text = drSQL8.Item(11)
            txtUDF7.Text = drSQL8.Item(12)
            txtUDF8.Text = drSQL8.Item(13)
            txtUDF9.Text = drSQL8.Item(14)
            txtUDF10.Text = drSQL8.Item(15)

        End If

        cnSQL8.Close()

        'end of loading item details


    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub MyGroupBox1_Enter(sender As Object, e As EventArgs) Handles MyGroupBox1.Enter

    End Sub
End Class