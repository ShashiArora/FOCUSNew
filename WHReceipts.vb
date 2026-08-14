
Option Explicit On


Imports System.IO
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection

Public Class WHReceipts
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"
    Private Sub txtGRNNO_KeyDown(sender As Object, e As KeyEventArgs) Handles txtGRNNO.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If


        btnImageClear.Enabled = True
    End Sub

    Private Sub txtGRNNO_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtGRNNO.KeyPress

        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If



        btnImageClear.Enabled = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles txtGRNNO.TextChanged

    End Sub

    Private Sub btnImageClear_Click(sender As Object, e As EventArgs) Handles btnImageClear.Click

        'check GRN no. already saved or not.*****************************************************************************************

        Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL3 As SqlCommand
        Dim drSQL3 As SqlDataReader
        Dim strSQL3 As String


        strSQL3 = "SELECT [GRN_NO] FROM [FSPrograms].[dbo].[TSS_WH_Material_Receipt] WHERE [GRN_NO] = '" & txtGRNNO.Text & "'"

        cnSQL3.Open()
        cmSQL3 = New SqlCommand(strSQL3, cnSQL3)
        drSQL3 = cmSQL3.ExecuteReader()

        If drSQL3.Read() Then

            ' If IsDBNull(drSQL3.Item(0)) Then
            ' txtdc.Text = 1

            If Len(drSQL3.Item(0)) > 0 Then

                MsgBox("This GRN is already added to the stock", vbInformation)
                Exit Sub
            Else
            End If

        End If
        cnSQL3.Close()


        'end of checking GRN*******************************************************************************************************


        If Len(txtGRNNO.Text) > 0 Then



            ''  mode = "EDIT"
            DataGridViewReceipts.Columns.Clear()

            Dim AlarmColumn1 As New DataGridViewCheckBoxColumn(False)

            AlarmColumn1.Name = "Sel"
            AlarmColumn1.HeaderText = "Select"
            AlarmColumn1.ReadOnly = False
            ' DataGridViewReceipts.Columns("Sel").Width = 45


            DataGridViewReceipts.Columns.Add(AlarmColumn1)
            DataGridViewReceipts.ReadOnly = False

            DataGridViewReceipts.Visible = True
            DataGridViewReceipts.Enabled = True

            DataGridViewReceipts.Columns("Sel").Width = 50


            Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim strSQL As String


            Dim stockDC As DataSet = New DataSet


            strSQL = "SELECT [POLineNumber] as PO_LN, NonInventoryItemNumber as ItemNumber,	NonInventoryItemDescription as Description,	ActualQty as GRN_Qty,POLineUM as UOM, 'O'as Type, '' as STR, '' as BIN , (PONumber + '-' + convert(varchar(3),POLineNumber) + '-' + convert(varchar,getdate(),12)) as LotNumber, [GRNDate], [VendorID] ,[PONumber],[HistoryPOReceiptKey] as HKey, ItemStandardLocalUnitPrice as UnitPrice " & _
                     "FROM [FSPrograms].[dbo].[TSS_WH_GRNDetails_GTYPE] WHERE  [GRNNo] = '" & txtGRNNO.Text & "'"
            'order BY [POLineNumber]"


            '[PONumber] & [POLineNumber] + " / " + [GRNDate] as LotNumber,

            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()

            stockDAC.TableMappings.Add("Table", "Enq")
            'get data
            stockDAC.Fill(stockDC)


            DataGridViewReceipts.DataSource = stockDC.Tables(0)

            'txtGRNDate.Text = DataGridViewReceipts.CurrentRow.Cells(4).Value
            'txtVendorID.Text = DataGridViewReceipts.CurrentRow.Cells(5).Value
            'txtPONumber.Text = DataGridViewReceipts.CurrentRow.Cells(6).Value

            ' DataGridViewReceipts.Columns.Add(checkCol)

            DataGridViewReceipts.Columns("PO_LN").ReadOnly = True
            DataGridViewReceipts.Columns("PO_LN").Width = 60
            DataGridViewReceipts.Columns("PO_LN").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



            DataGridViewReceipts.Columns("ItemNumber").ReadOnly = True
            DataGridViewReceipts.Columns("ItemNumber").Width = 125
            DataGridViewReceipts.Columns("Description").ReadOnly = True
            DataGridViewReceipts.Columns("Description").Width = 175

            DataGridViewReceipts.Columns("GRN_Qty").ReadOnly = True
            DataGridViewReceipts.Columns("GRN_Qty").Width = 75
            DataGridViewReceipts.Columns("GRN_Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridViewReceipts.Columns("GRN_Qty").DefaultCellStyle.Format = "N2"

            DataGridViewReceipts.Columns("UOM").ReadOnly = True
            DataGridViewReceipts.Columns("UOM").Width = 40

            DataGridViewReceipts.Columns("Type").ReadOnly = True
            DataGridViewReceipts.Columns("Type").Width = 40

            DataGridViewReceipts.Columns("STR").Width = 45
            DataGridViewReceipts.Columns("STR").HeaderCell.Style.BackColor = Color.Gray


            DataGridViewReceipts.Columns("BIN").Width = 85
            DataGridViewReceipts.Columns("BIN").HeaderCell.Style.BackColor = Color.Gray


            DataGridViewReceipts.Columns("LotNumber").ReadOnly = True
            DataGridViewReceipts.Columns("LotNumber").Width = 0
            DataGridViewReceipts.Columns("LotNumber").Visible = False

            DataGridViewReceipts.Columns("GRNDate").ReadOnly = True
            DataGridViewReceipts.Columns("GRNDate").Width = 0
            DataGridViewReceipts.Columns("GRNDate").Visible = False

            DataGridViewReceipts.Columns("VendorID").ReadOnly = True
            DataGridViewReceipts.Columns("VendorID").Width = 0
            DataGridViewReceipts.Columns("VendorID").Visible = False

            DataGridViewReceipts.Columns("PONumber").ReadOnly = True
            DataGridViewReceipts.Columns("PONumber").Width = 75


            DataGridViewReceipts.Columns("HKey").ReadOnly = True
            DataGridViewReceipts.Columns("HKey").Width = 0
            DataGridViewReceipts.Columns("HKey").Visible = False

            DataGridViewReceipts.Columns("UnitPrice").ReadOnly = True
            DataGridViewReceipts.Columns("UnitPrice").Width = 0
            DataGridViewReceipts.Columns("UnitPrice").Visible = False

            cnSQL.Close()

            'calling items from item master - Combo box loading
            Dim comboBoxColumn As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
            comboBoxColumn.HeaderText = "Items"

            comboBoxColumn.HeaderCell.Style.BackColor = Color.Gray


            comboBoxColumn.Width = 200
            comboBoxColumn.Name = "comboBoxColumn"
            DataGridViewReceipts.Columns.Add(comboBoxColumn)

            'Loop through the DataGridView Rows.
            For Each row As DataGridViewRow In DataGridViewReceipts.Rows

                'Reference the ComboBoxCell.
                Dim comboBoxCell As DataGridViewComboBoxCell = CType(row.Cells(15), DataGridViewComboBoxCell)


                'Fetch the Countries from Database.
                Dim dt As DataTable = Me.GetData("SELECT 0 as Item_Key, ' ' as [Part_Number] union SELECT Item_Key ,[Part_Number] FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster]  where Item_Type = '" & ComboBoxMatType.Text & "' ORDER BY Item_Key Desc ")
                '   dt.Columns("Items").DefaultValue = "Selet Item"
                'Loop through the DataTable Rows.
                For Each drow As DataRow In dt.Rows

                    'Fetch the CustomerId (Primary Key) value.Asc

                    Dim Part As String = drow(1).ToString
                    'drow(0)
                    'Add the Country value to the ComboBoxCell.
                    comboBoxCell.Items.Add(drow(1))

                    'Except for CustomerId #3.
                    '  If (Part <> "3") Then

                    'Compare the value of PART
                    '    If (row.Cells(0).Value.ToString = Part) Then

                    'Once CustomerId is matched, select the Country in ComboBoxCell.
                    comboBoxCell.Value = drow(1)
                    '   dt.Columns("Items").DefaultValue = "Selet Item"
                    'End If
                    'End If
                Next
                '  dt.Columns("Items").DefaultValue = "Selet Item"
            Next


        Else
            '   MsgBox("Pl select the Line", vbInformation)
            Exit Sub
        End If

    End Sub

    Private Sub Receipts_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        transtype = "Receipts"

        dtpReceiptdt.Format = DateTimePickerFormat.Custom
        dtpReceiptdt.CustomFormat = "dd/MM/yyyy"

        comboloadtype()

        'for combo box
    End Sub

    Private Sub comboloadtype()
        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand


        'dept_type load

        strSql = "Select TypeKey, Item_Type  FROM [FSPrograms].[dbo].[TSS_WH_ItemTypes] where Status like 'A%' "

        cmSQL = New SqlCommand(strSql, sqlCon)
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource As SqlDataAdapter = New SqlDataAdapter
        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eitem")
        With ComboBoxMatType
            .DataSource = source.Tables("eitem")
            .DisplayMember = "Item_Type"
            .ValueMember = "TypeKey"
            .SelectedIndex = 0
        End With
    End Sub

    Private Sub COMBOlOAD()  ' not used anywhere just example

        Dim AutoList As AutoCompleteStringCollection = New AutoCompleteStringCollection
        '  Dim cbo As ComboBox
        Dim ComboColumnName As String = "C1"


        Dim Source As String() = New String() {"O", "H", "I"}

        Dim Column1 As New DataGridViewComboBoxColumn With
            {
                .DataSource = Source,
                .DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                .Name = ComboColumnName,
                .HeaderText = "Type",
                .SortMode = DataGridViewColumnSortMode.NotSortable
            }

        Dim Column2 As New DataGridViewTextBoxColumn With
            {
                .Name = "C2",
                .HeaderText = "Type"
            }
        AutoList.AddRange(Source)

        DataGridViewReceipts.Columns.AddRange(New DataGridViewColumn() {Column1, Column2})

        DataGridViewReceipts.Rows.Add(New Object() {"Onhand", 1})
        DataGridViewReceipts.Rows.Add(New Object() {"Hold", 2})
        DataGridViewReceipts.Rows.Add(New Object() {"Inspection", 3})

    End Sub

    Private Sub btnRecAccept_Click(sender As Object, e As EventArgs) Handles btnRecAccept.Click
        Dim checkdt As Date
        checkdt = Today
        Dim item As String

        Dim msgb2 As String

        Dim cnSQL6 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL6 As SqlCommand
        Dim drSQL6 As SqlDataReader
        Dim strSQL6 As String
        Dim partno As Integer
        partno = 0

        msgb2 = MsgBox("Are you sure of saving ?", vbYesNo)

        If msgb2 = vbYes Then

            Dim Iblankc As Integer
            Dim stkbinc As Integer
            Iblankc = 0
            stkbinc = 0


            'Dim sr As New IO.StreamReader
            Dim sr As IO.StreamReader
            '  A string to hold each line as it is read
            Dim line As String = String.Empty



            For i As Integer = 0 To DataGridViewReceipts.RowCount - 1

                ' Dim Checked As Boolean = CType(DataGridViewReceipts.CurrentCell.Value, Boolean)
                Dim Checked As Boolean = CType(Me.DataGridViewReceipts.Rows(i).Cells("Sel").Value, Boolean)
                If Checked Then
                    '   For Each row As DataGridViewRow In DataGridViewReceipts.Rows
                    'Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("Sel").Value)
                    'If isSelected Then
                    '   If IsDBNull(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) And IsDBNull(Me.DataGridViewReceipts.Rows(i).Cells("comboBoxColumn").Value) Then
                    'Iblankc = Iblankc + 1
                    '     If Len(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) < 2 And Len(Me.DataGridViewReceipts.Rows(i).Cells("comboBoxColumn").Value) < 2 Then

                    ' If IsDBNull(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) And Me.DataGridViewReceipts.Rows(i).Cells("comboBoxColumn").Value = "Select Item" Then



                    'CHECK ITEM Existing in Item Master or not 

                    '   If Len((Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value)) > 3 Then

                    '   If Not IsDBNull((Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value)) Then

                    If Val(Len(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value)) > 3 Then

                        strSQL6 = "SELECT Part_Number FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster] with (nolock) WHERE Part_Number = '" & (Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) & "'"

                        cnSQL6.Open()
                        cmSQL6 = New SqlCommand(strSQL6, cnSQL6)
                        drSQL6 = cmSQL6.ExecuteReader()

                        If drSQL6.Read() Then

                            ' If IsDBNull(drSQL3.Item(0)) Then
                            ' txtdc.Text = 1

                            If Len(drSQL6.Item(0)) < 3 Then

                                partno = partno + 1

                            End If

                        Else
                            partno = partno + 1

                        End If


                        cnSQL6.Close()

                    End If


                    If partno >= 1 Then
                        MsgBox("Some of the part numbers are not existing in the Master, please check", vbInformation)
                        Exit Sub
                    End If


                    'end of checking



                    '   If IsDBNull(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) And Me.DataGridViewReceipts.Rows(i).Cells("comboBoxColumn").Value = "Select Item" Then

                    If Len(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) <= 3 And Len(Me.DataGridViewReceipts.Rows(i).Cells("comboBoxColumn").Value) <= 3 Then


                        Iblankc = Iblankc + 1
                    End If

                    If CheckBoxupdate.Checked = False Then

                        If Len(Me.DataGridViewReceipts.Rows(i).Cells("STR").Value) <= 1 Or Len(Me.DataGridViewReceipts.Rows(i).Cells("BIN").Value) <= 1 Then
                            stkbinc = stkbinc + 1
                        End If
                    End If


                    If Iblankc >= 1 Or stkbinc >= 1 Then
                        MsgBox("ItemNumber/Stockroom /Bin should not be blank ", vbInformation)
                        Exit Sub
                    End If
                End If


            Next

            'end of itemnumber blank checking ***************************************************************************************

            'stock room / BIN checking correctness checking *****************************

            Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL3 As SqlCommand
            Dim drSQL3 As SqlDataReader
            Dim strSQL3 As String
            Dim loc As String

            If CheckBoxupdate.Checked = False Then
                For i As Integer = 0 To DataGridViewReceipts.RowCount - 1

                    'Dim Checked As Boolean = CType(DataGridViewReceipts.CurrentCell.Value, Boolean)
                    Dim Checked As Boolean = CType(Me.DataGridViewReceipts.Rows(i).Cells("Sel").Value, Boolean)

                    If Checked Then


                        loc = LTrim(RTrim((Me.DataGridViewReceipts.Rows(i).Cells("STR").Value))) + LTrim(RTrim((Me.DataGridViewReceipts.Rows(i).Cells("BIN").Value)))


                        strSQL3 = "SELECT ltrim(rtrim([Stockroom]))+ ltrim(rtrim([Bin])) FROM [FSDBBR].[dbo].[_NoLock_FS_InventoryLocation] union SELECT ltrim(rtrim([StockRoom]))+ ltrim(rtrim([BIN])) FROM [FSPrograms].[dbo].[TSS_WH_StockRooms] where  ltrim(rtrim([StockRoom]))+ ltrim(rtrim([BIN])) = '" & loc & "' "


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

                Next
                'end of checking

            End If

            Dim strsql2 As String
            Dim cmSQL2 As SqlCommand
            Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)

            'generating regno
            Dim checkrow As Integer
            transmode = "Add"
            nogenerate()

            txtRceiptNo.Text = receiptno

            If txtRceiptNo.Text > 0 Then

                cnSQL2.Open()

                curdate = System.DateTime.Now()



                For i As Integer = 0 To DataGridViewReceipts.RowCount - 1

                    'Dim Checked As Boolean = CType(DataGridViewReceipts.CurrentCell.Value, Boolean)
                    Dim Checked As Boolean = CType(Me.DataGridViewReceipts.Rows(i).Cells("Sel").Value, Boolean)
                    If Checked Then

                        checkrow = checkrow + 1

                        '  If Len(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) < 2 Then
                        If IsDBNull(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) Then

                            item = (Me.DataGridViewReceipts.Rows(i).Cells("comboBoxColumn").Value)

                        ElseIf Len(Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value) < 2 Then

                            item = (Me.DataGridViewReceipts.Rows(i).Cells("comboBoxColumn").Value)

                        Else
                            item = Me.DataGridViewReceipts.Rows(i).Cells("ItemNumber").Value
                        End If




                        Dim MyString As String = Me.DataGridViewReceipts.Rows(i).Cells("Description").Value

                        Dim MyNewString As String = MyString.Replace("'", String.Empty)



                        If CheckBoxupdate.Checked = True Then

                            strsql2 = "insert TSS_WH_Material_Receipt values (" & txtRceiptNo.Text & ",'" & dtpReceiptdt.Value & "', '" & txtGRNNO.Text & "','" & Me.DataGridViewReceipts.Rows(i).Cells("VendorID").Value & "'," & _
                     "'" & Me.DataGridViewReceipts.Rows(i).Cells("PONumber").Value & "','" & Me.DataGridViewReceipts.Rows(i).Cells("PO_LN").Value & "','" & item & "','" & MyNewString & "'," & Me.DataGridViewReceipts.Rows(i).Cells("GRN_Qty").Value & ",'" & Me.DataGridViewReceipts.Rows(i).Cells("UOM").Value & "'," & Me.DataGridViewReceipts.Rows(i).Cells("UnitPrice").Value & ",'" & Me.DataGridViewReceipts.Rows(i).Cells("Type").Value & "', " & _
                     "'" & Me.DataGridViewReceipts.Rows(i).Cells("LotNumber").Value & " ','" & dtpReceiptdt.Value & "', '" & txtStockRoom.Text & "', '" & ComboBoxBINS.Text & "',0, '" & txtRemarks.Text & "','" & Me.DataGridViewReceipts.Rows(i).Cells("HKey").Value & "', '" & username & "'," & _
                     "'" & curdate & "')"


                        Else


                            strsql2 = "insert TSS_WH_Material_Receipt values (" & txtRceiptNo.Text & ",'" & dtpReceiptdt.Value & "', '" & txtGRNNO.Text & "','" & Me.DataGridViewReceipts.Rows(i).Cells("VendorID").Value & "'," & _
                            "'" & Me.DataGridViewReceipts.Rows(i).Cells("PONumber").Value & "','" & Me.DataGridViewReceipts.Rows(i).Cells("PO_LN").Value & "','" & item & "','" & MyNewString & "'," & Me.DataGridViewReceipts.Rows(i).Cells("GRN_Qty").Value & ",'" & Me.DataGridViewReceipts.Rows(i).Cells("UOM").Value & "'," & Me.DataGridViewReceipts.Rows(i).Cells("UnitPrice").Value & ",'" & Me.DataGridViewReceipts.Rows(i).Cells("Type").Value & "', " & _
                            "'" & Me.DataGridViewReceipts.Rows(i).Cells("LotNumber").Value & " ','" & dtpReceiptdt.Value & "', '" & Me.DataGridViewReceipts.Rows(i).Cells("STR").Value & "', '" & Me.DataGridViewReceipts.Rows(i).Cells("BIN").Value & "',0, '" & txtRemarks.Text & "','" & Me.DataGridViewReceipts.Rows(i).Cells("HKey").Value & "', '" & username & "'," & _
                            "'" & curdate & "')"

                        End If


                        cmSQL2 = New SqlCommand(strsql2, cnSQL2)

                        If cmSQL2.ExecuteNonQuery() = 0 Then
                            MsgBox("Cannot Save the Details. " & strsql2, MsgBoxStyle.Exclamation, "Error!")
                            '  txtRegNo.Text = 0
                            'Application.Exit()
                            Exit Sub
                        End If

                    End If

                Next

            End If

            cnSQL2.Close()

            If checkrow = 0 Then
                MsgBox("No records selected for update", vbInformation)
                Exit Sub
            Else

                transmode = "Update"
                nogenerate()

                MsgBox("Receipts updated", vbInformation)
                btnRecAccept.Enabled = False

            End If

        End If

        '  If cmSQL.ExecuteNonQuery() = 0 Then
        'MsgBox("Cannot Save Header Section. " & strsql, MsgBoxStyle.Exclamation, "Error!")
        'txtRegNo.Text = 0
        'Application.Exit()

        'Else
        'MsgBox("Header section saved.", vbInformation)
        'update the regno.back to table
        'btnHeaderSave.Enabled = False

        'mail to finance
    End Sub

    Private Sub DataGridViewReceipts_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewReceipts.CellContentClick

        'Dim s As Integer

        's = DataGridViewReceipts.CurrentCell.ColumnIndex

        ''---

        ''DataGridViewReceipts.Columns.Add("c1", "c1")
        ''DataGridViewReceipts.Columns.Add(New DataGridViewComboBoxColumn)
        ''DataGridViewReceipts.Columns(12).Name = "c2"
        ''DataGridViewReceipts.Columns(12).HeaderText = "c2"
        ''DirectCast(DataGridViewReceipts.Columns(12), DataGridViewComboBoxColumn).DataSource = New String() {"one", "two", "three"}

        ''Dim dt As New DataTable
        ''dt.Columns.AddRange(New DataColumn() {New DataColumn, New DataColumn})

        ''dt.Rows.Add("1", "one")
        ''dt.Rows.Add("2", "two")
        ''dt.Rows.Add("3", "three")

        ''DataGridViewReceipts.AutoGenerateColumns = False
        ''DataGridViewReceipts.Columns(0).DataPropertyName = dt.Columns(0).ColumnName
        ''DataGridViewReceipts.Columns(12).DataPropertyName = dt.Columns(12).ColumnName

        ''DataGridViewReceipts.DataSource = dt





        ''----
        'If s = 0 Then
        '    'Fetch the data from Database.
        '    DataGridViewReceipts.DataSource = Me.GetData("SELECT [CSR],[CSRKey]  FROM [FSDBBR].[dbo].[_NoLock_FS_CSR] WHERE CSR LIKE 'AR%'")

        '    DataGridViewReceipts.AllowUserToAddRows = False

        '    'Add a ComboBox Column to the DataGridView.

        '    ' DataGridViewReceipts.Columns.Add("c1", "c1")
        '    ' DataGridViewReceipts.Columns.Add(New DataGridViewComboBoxColumn)
        '    '   DataGridViewReceipts.Columns(11).Name = "c2"
        '    'DataGridViewReceipts.Columns(1).HeaderText = "c2"

        '    Dim comboBoxColumn As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        '    comboBoxColumn.HeaderText = "Items"
        '    comboBoxColumn.Width = 100
        '    comboBoxColumn.Name = "comboBoxColumn"
        '    DataGridViewReceipts.Columns.Add(comboBoxColumn)

        '    'Loop through the DataGridView Rows.
        '    For Each row As DataGridViewRow In DataGridViewReceipts.Rows

        '        'Reference the ComboBoxCell.
        '        Dim comboBoxCell As DataGridViewComboBoxCell = CType(row.Cells(2), DataGridViewComboBoxCell)

        '        'Insert the Default Item to ComboBoxCell.
        '        comboBoxCell.Items.Add("Select Item")

        '        'Set the Default Value as the Selected Value.
        '        comboBoxCell.Value = "Select Item"

        '        'Fetch the Countries from Database.
        '        Dim dt As DataTable = Me.GetData("SELECT Item_Key ,[Part_Number]FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster]")

        '        'Loop through the DataTable Rows.
        '        For Each drow As DataRow In dt.Rows

        '            'Fetch the CustomerId (Primary Key) value.

        '            Dim Part As String = drow(1).ToString
        '            'drow(0)
        '            'Add the Country value to the ComboBoxCell.
        '            comboBoxCell.Items.Add(drow(1))

        '            'Except for CustomerId #3.
        '            '  If (Part <> "3") Then

        '            'Compare the value of PART
        '            '    If (row.Cells(0).Value.ToString = Part) Then

        '            'Once CustomerId is matched, select the Country in ComboBoxCell.
        '            comboBoxCell.Value = drow(1)
        '            'End If
        '            'End If
        '        Next
        '    Next



        'End If





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



    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub ComboBoxBINS_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxBINS.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If




    End Sub

    Private Sub ComboBoxBINS_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBoxBINS.KeyPress

    End Sub

    Private Sub ComboBoxBINS_MouseClick(sender As Object, e As MouseEventArgs) Handles ComboBoxBINS.MouseClick

        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand


        'dept_type load

        ' strSql = "SELECT Bin,InventoryLocationKey   FROM [FSDBBR].[dbo].[_NoLock_FS_InventoryLocation] WHERE Stockroom = '" & txtStockRoom.Text & " '"

        strSql = "SELECT BIN,ID FROM [FSPrograms].[dbo].[TSS_WH_StockRooms] WITH (NOLOCK) UNION SELECT Bin, InventoryLocationKey FROM FSDBBR.dbo._NoLock_FS_InventoryLocation  WHERE Stockroom = '" & txtStockRoom.Text & " '"



        cmSQL = New SqlCommand(strSql, sqlCon)
        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
        Dim ESource As SqlDataAdapter = New SqlDataAdapter
        ESource.SelectCommand = sqlCmd
        ESource.Fill(source, "eitem")
        With ComboBoxBINS
            .DataSource = source.Tables("eitem")
            .DisplayMember = "BIN"
            .ValueMember = "ID"
            .SelectedIndex = 0
        End With

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxBINS.SelectedIndexChanged






    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click



        txtGRNNO.Text = ""
        txtGRNscan.Text = ""
        txtPONumber.Text = ""
        txtRceiptNo.Text = ""
        txtRemarks.Text = ""
        txtStockRoom.Text = ""
        ComboBoxBINS.Text = ""
        txtVendorID.Text = ""
        CheckBoxupdate.Checked = False

        btnRecAccept.Enabled = True

        DataGridViewReceipts.Columns.Clear()
        ' DataGridViewReceipts.DataSource = DBNull.Value

        btnImageClear.Enabled = True

        CheckBoxSelAll.Visible = True
        txtStockRoom.Visible = True
        ComboBoxBINS.Visible = True
        CheckBoxupdate.Visible = True

        txtGRNNO.Focus()


    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnView.Click


        btnImageClear.Enabled = False
        DataGridViewReceipts.DataSource = DBNull.Value

        DataGridViewReceipts.Columns.Clear()
        txtStockRoom.Text = ""
        ComboBoxBINS.Text = ""
        CheckBoxupdate.Checked = False
        txtRemarks.Text = ""

        CheckBoxSelAll.Visible = False
        txtStockRoom.Visible = False
        ComboBoxBINS.Visible = False
        CheckBoxupdate.Visible = False





        If Len(txtGRNNO.Text) <= 2 Then
            MsgBox("GRN number to be entered", vbInformation)

            Exit Sub
        End If

        If Len(txtGRNNO.Text) > 0 Then
            Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL1 As SqlCommand
            Dim drSQL1 As SqlDataReader
            Dim strSQL1 As String
            Dim msgb As String


            strSQL1 = "SELECT [GRN_NO] FROM [FSPrograms].[dbo].[TSS_WH_Material_Receipt] WHERE [GRN_NO] = '" & txtGRNNO.Text & "'"

            cnSQL1.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
            drSQL1 = cmSQL1.ExecuteReader()

            If drSQL1.Read() Then

                ' If IsDBNull(drSQL3.Item(0)) Then
                ' txtdc.Text = 1

                If Len(drSQL1.Item(0)) > 0 Then

                    msgb = MsgBox("This GRN is already added to the stock, do you want to view ?", vbYesNo)

                    If msgb = vbYes Then
                        'load items, make accept button disable
                        btnRecAccept.Enabled = False
                        DataGridViewReceipts.Visible = True
                        DataGridViewReceipts.Enabled = True
                        Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)
                        Dim strSQL3 As String


                        Dim stockDC3 As DataSet = New DataSet


                        strSQL3 = "select LN as POLN, PartNumber as ItemNumber, Vend_Desc as Description, Qty as 'GRNQty', Type, StockRoom AS STKRm, Bin, LotNumber,VendId AS VendorID, PONumber, Remarks  from [TSS_WH_Material_Receipt] WHERE GRN_NO = '" & txtGRNNO.Text & "' order by LN "
                        'order BY [POLineNumber]"


                        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL3, cnSQL3)
                        Dim stockDAC3 As SqlDataAdapter = New SqlDataAdapter

                        stockDAC3.SelectCommand = sqlCmd
                        cnSQL3.Open()

                        stockDAC3.TableMappings.Add("Table", "Enq")
                        'get data
                        stockDAC3.Fill(stockDC3)


                        DataGridViewReceipts.DataSource = stockDC3.Tables(0)

                        ' Dim data As String = DataGridView1.SelectedRows(0).Cells(5).Value.ToString
                        ' txtRemarks.Text = DataGridViewReceipts.SelectedRows(0).Cells(0).Value.ToString
                        'txtRemarks.Text = DataGridViewReceipts.CurrentRow.Cells(9).Value.ToString

                        txtRemarks.Text = DataGridViewReceipts.CurrentRow.Cells(10).Value.ToString

                        DataGridViewReceipts.Columns("POLN").ReadOnly = True
                        DataGridViewReceipts.Columns("POLN").Width = 45
                        DataGridViewReceipts.Columns("ItemNumber").ReadOnly = True
                        DataGridViewReceipts.Columns("ItemNumber").Width = 145
                        DataGridViewReceipts.Columns("Description").ReadOnly = True
                        DataGridViewReceipts.Columns("Description").Width = 175

                        DataGridViewReceipts.Columns("GRNQty").ReadOnly = True
                        DataGridViewReceipts.Columns("GRNQty").Width = 100
                        DataGridViewReceipts.Columns("GRNQty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        DataGridViewReceipts.Columns("GRNQty").DefaultCellStyle.Format = "N2"

                        DataGridViewReceipts.Columns("Type").ReadOnly = True
                        DataGridViewReceipts.Columns("Type").Width = 40

                        DataGridViewReceipts.Columns("STKRm").Width = 85
                        DataGridViewReceipts.Columns("STKRm").ReadOnly = True
                        DataGridViewReceipts.Columns("BIN").Width = 115
                        DataGridViewReceipts.Columns("BIN").ReadOnly = True

                        DataGridViewReceipts.Columns("LotNumber").ReadOnly = True
                        DataGridViewReceipts.Columns("LotNumber").Width = 135

                        '     DataGridViewReceipts.Columns("GRNDate").ReadOnly = True
                        '    DataGridViewReceipts.Columns("GRNDate").Width = 0

                        DataGridViewReceipts.Columns("VendorID").ReadOnly = True
                        DataGridViewReceipts.Columns("VendorID").Width = 75

                        DataGridViewReceipts.Columns("PONumber").ReadOnly = True
                        DataGridViewReceipts.Columns("PONumber").Width = 75

                        DataGridViewReceipts.Columns("Remarks").ReadOnly = True
                        DataGridViewReceipts.Columns("Remarks").Width = 0

                        cnSQL3.Close()

                    Else
                        Exit Sub
                    End If
                End If
                cnSQL1.Close()
                '  Exit Sub
            End If
            'Exit Sub
        End If

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        btnImageClear.Enabled = False
        DataGridViewReceipts.DataSource = DBNull.Value

        DataGridViewReceipts.Columns.Clear()
        txtStockRoom.Text = ""
        ComboBoxBINS.Text = ""
        CheckBoxupdate.Checked = False
        txtRemarks.Text = ""


        If Len(txtGRNNO.Text) <= 2 Then
            MsgBox("GRN number to be entered", vbInformation)
            Exit Sub
        End If


        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cnSQL3 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim cmSQL1 As SqlCommand
        Dim cmSQL2 As SqlCommand
        Dim cmSQL3 As SqlCommand

        Dim drSQL1 As SqlDataReader
        Dim drSQL2 As SqlDataReader
        'Dim drSQL3 As SqlDataReader

        Dim strSQL1 As String
        Dim strSQL2 As String
        Dim strSQL3 As String

        Dim Msgb As String

        Dim notref As Integer
        Dim ref As Integer
        Dim LOT As String


        Msgb = MsgBox("Deletion is possible, if this GRN is not issued. Are you sure of deleting ?", vbYesNo)

        If Msgb = vbYes Then


            strSQL1 = "SELECT distinct LotNumber  FROM [FSPrograms].[dbo].[TSS_WH_Material_Receipt] where GRN_NO = '" & txtGRNNO.Text & "'"
            cnSQL1.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
            drSQL1 = cmSQL1.ExecuteReader()
            If drSQL1.Read() Then
                Do While drSQL1.Read()

                    LOT = drSQL1.Item(0)
                    strSQL2 = "SELECT LotNumber  FROM [FSPrograms].[dbo].[TSS_WH_MaterialIssueDetail] where LotNumber = '" & LOT & "'"

                    cnSQL2.Close()

                    cnSQL2.Open()
                    cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
                    drSQL2 = cmSQL2.ExecuteReader()


                    If drSQL2.Read() Then

                        If IsDBNull(drSQL2.Item(0)) Then
                            notref = notref + 1
                        Else

                            ref = ref + 1
                        End If
                    Else

                        notref = notref + 1
                    End If


                Loop

            Else
                MsgBox("This GRN is not inwarded ! ", vbInformation)
                Exit Sub
            End If



            If ref = 0 Then

                strSQL3 = "delete  FROM [FSPrograms].[dbo].[TSS_WH_Material_Receipt] where GRN_NO = '" & txtGRNNO.Text & "'"
                cnSQL3.Open()
                cmSQL3 = New SqlCommand(strSQL3, cnSQL3)


                If cmSQL3.ExecuteNonQuery() = 0 Then

                    MsgBox("Error while deleting", vbInformation)
                    Exit Sub
                End If

                MsgBox("GRN deleted", vbInformation)
                Exit Sub
            Else
                MsgBox("This GRN is already referred in Issue, can't be deleted", vbInformation)
                Exit Sub

            End If
            cnSQL1.Close()
            cnSQL2.Close()
            cnSQL3.Close()

        End If


    End Sub

    Private Sub CheckBoxSelAll_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSelAll.CheckedChanged

        If CheckBoxSelAll.Checked = True Then


            For i As Integer = 0 To DataGridViewReceipts.RowCount - 1

                Me.DataGridViewReceipts.Rows(i).Cells("Sel").Value = CheckState.Checked

            Next


        ElseIf CheckBoxSelAll.Checked = False Then


            For i As Integer = 0 To DataGridViewReceipts.RowCount - 1

                Me.DataGridViewReceipts.Rows(i).Cells("Sel").Value = CheckState.Unchecked

            Next

        End If


    End Sub

    Private Sub CheckBoxupdate_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxupdate.CheckedChanged


        If CheckBoxupdate.Checked = True Then

            For i As Integer = 0 To DataGridViewReceipts.RowCount - 1

                Me.DataGridViewReceipts.Rows(i).Cells("STR").Value = txtStockRoom.Text
                Me.DataGridViewReceipts.Rows(i).Cells("BIN").Value = ComboBoxBINS.Text

            Next


        ElseIf CheckBoxupdate.Checked = False Then


            For i As Integer = 0 To DataGridViewReceipts.RowCount - 1

                Me.DataGridViewReceipts.Rows(i).Cells("STR").Value = ""
                Me.DataGridViewReceipts.Rows(i).Cells("BIN").Value = ""

            Next

        End If

    End Sub

    Private Sub ComboBoxMatType_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBoxMatType.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If

    End Sub

    Private Sub ComboBoxMatType_MouseHover(sender As Object, e As EventArgs) Handles ComboBoxMatType.MouseHover

    End Sub

    Private Sub ComboBoxMatType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxMatType.SelectedIndexChanged

    End Sub

    Private Sub CheckBoxLoad_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxLoad.CheckedChanged
        If CheckBoxLoad.Checked = True Then
            '    Dim comboBoxColumn As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
            '   comboBoxColumn.HeaderText = "Items"

            'comboBoxColumn.Width = 150
            'comboBoxColumn.Name = "comboBoxColumn"
            'DataGridViewReceipts.Columns.Add(comboBoxColumn)

            'Loop through the DataGridView Rows.
            '     Dim comboBoxCell As DataGridViewComboBoxCell = CType(row.Cells(14), DataGridViewComboBoxCell)



            For Each row As DataGridViewRow In DataGridViewReceipts.Rows

                'Reference the ComboBoxCell.
                Dim comboBoxCell As DataGridViewComboBoxCell = CType(row.Cells(15), DataGridViewComboBoxCell)

                If Len(comboBoxCell.Value) <= 3 Then



                    Dim dt As DataTable = Me.GetData("SELECT 0 as Item_Key, ' ' as [Part_Number] union SELECT Item_Key ,[Part_Number] FROM [FSPrograms].[dbo].[TSS_WH_ItemMaster]  where Item_Type = '" & ComboBoxMatType.Text & "' order by Item_Key Desc ")
                    '   dt.Columns("Items").DefaultValue = "Select Item"
                    'Loop through the DataTable Rows.
                    For Each drow As DataRow In dt.Rows
                        'Fetch the CustomerId (Primary Key) value.Rows


                        Dim Part As String = drow(1).ToString
                        'drow(0)
                        'Add the Country value to the ComboBoxCell.
                        comboBoxCell.Items.Add(drow(1))


                        comboBoxCell.Value = drow(1)

                    Next
                    '  dt.Columns("Items").DefaultValue = "Selet Item"
                End If

            Next

        End If


    End Sub

    Private Sub GroupBoxEdit_Enter(sender As Object, e As EventArgs) Handles GroupBoxEdit.Enter

    End Sub

    Private Sub txtRemarks_KeyDown(sender As Object, e As KeyEventArgs) Handles txtRemarks.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If
    End Sub

    Private Sub txtRemarks_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtRemarks.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If
    End Sub

    Private Sub txtRemarks_TextChanged(sender As Object, e As EventArgs) Handles txtRemarks.TextChanged

    End Sub

    Private Sub DataGridViewReceipts_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewReceipts.RowHeaderMouseClick

    End Sub

    Private Sub txtStockRoom_KeyDown(sender As Object, e As KeyEventArgs) Handles txtStockRoom.KeyDown
      If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If

    End Sub

    Private Sub txtStockRoom_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtStockRoom.KeyPress
        Dim allowedChars As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,-=+!@#$%^&*()?<>._/{}[]\|" & Chr(Keys.Back) & Chr(Keys.Space)

        If allowedChars.IndexOf(e.KeyChar) = -1 Then
            ' Invalid Character
            e.Handled = True
        End If


    End Sub

    Private Sub txtStockRoom_TextChanged(sender As Object, e As EventArgs) Handles txtStockRoom.TextChanged

    End Sub

    Private Sub CheckBoxupdate_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBoxupdate.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If

    End Sub

    Private Sub CheckBoxLoad_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBoxLoad.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        End If

    End Sub
End Class