'Imports System
'Imports System.Data
Imports System.Data.SqlClient
'Imports System.ComponentModel
Imports System.Configuration
'Imports System.Collections
Imports System.Windows.Forms


Public Class Porv

    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Private Sub Porv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Additional duties upload


        porvidgen()

        ' DTPPorvDate.Value = '" & curdate & "'
        DataGridViewAddCharges.ReadOnly = False


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        strSQL = "sELECT ChargeID as ID,Description as Duty_Tax_Desc, 0.00 as Amount FROM LandedCostModule.dbo.AdditionalCharges"
        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)

        DataGridViewAddCharges.DataSource = stockDC.Tables(0)

        DataGridViewAddCharges.ReadOnly = False

        Dim col0 As DataGridViewColumn = DataGridViewAddCharges.Columns(0) 'ID
        col0.Width = 20
        DataGridViewAddCharges.Columns(0).ReadOnly = True

        Dim col1 As DataGridViewColumn = DataGridViewAddCharges.Columns(1) 'DUTY narration
        col1.Width = 90
        DataGridViewAddCharges.Columns(1).ReadOnly = True

        Dim col2 As DataGridViewColumn = DataGridViewAddCharges.Columns(2) 'Amount
        col2.Width = 90
        DataGridViewAddCharges.Columns(2).ReadOnly = False
        'DataGridViewAddCharges.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridViewAddCharges.Columns.Item(2).ValueType = GetType(Double)
        DataGridViewAddCharges.Columns.Item(2).DefaultCellStyle.Format = "n2"



        cnSQL.Close()

    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub DataGridViewPriceView_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    
    Private Sub Getpodetails()

        DataGridViewPODetails.Columns.Clear()


        Dim Alaram1 As New DataGridViewCheckBoxColumn(False)

        Alaram1.Name = "Sel"
        Alaram1.HeaderText = "Sel"
        Alaram1.ReadOnly = False
        DataGridViewPODetails.Columns.Add(Alaram1)
      


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim strSql As String
        curdate = System.DateTime.Now()
        cnSQL.Open()

        Dim stockDC As DataSet = New DataSet

        strSql = "SELECT FS_POHeader.PONumber, FS_POLine.POLineNumber AS 'Ln', FS_Item.ItemNumber, FS_Item.ItemDescription, " & _
        "FS_Item.ItemUM AS 'UOM', (FS_POLine.LineItemOrderedQuantity-FS_POLine.ReceiptQuantity) AS 'Open Qty', FS_POLineData.ItemUnitCost as 'UnitCost', " & _
        "0 as QtyRecd,0 AS Short, 0 as Amt_INR,0 as CustDuty,0 as FrtCharges,0 as C_And_F, 0 as CST,0 as Packing, 0 as ED_WH,0 as Amt_FC,0 as LnTax_Tot " & _
        "FROM FSDBBR.dbo.FS_Item FS_Item, FSDBBR.dbo.FS_POHeader FS_POHeader, FSDBBR.dbo.FS_POLine FS_POLine, FSDBBR.dbo.FS_POLineData FS_POLineData," & _
        "FSDBBR.dbo.FS_Vendor WHERE FS_Item.ItemKey = FS_POLine.ItemKey And FS_POLine.POHeaderKey = FS_POHeader.POHeaderKey And FS_POHeader.VendorKey = FS_Vendor.VendorKey " & _
        "AND FS_POLine.POLineKey = FS_POLineData.POLineKey AND FS_POLine.POLineStatus <= '4' " & _
        "AND  FS_POHeader.PONumber IN (SELECT a.PONumber FROM FSPrograms.dbo.PORV_POs a where a.PONumber = FS_POHeader.PONumber and a.PORV_ID = " & txtporvid.Text & " ) " & _
        " order by FS_POHeader.PONumber, FS_POLine.POLineNumber"


        '  strSql = "SELECT FS_POHeader.PONumber, FS_POLine.POLineNumber AS 'Ln', FS_Item.ItemNumber, FS_Item.ItemDescription, " & _
        '"FS_Item.ItemUM AS 'UOM', (FS_POLine.LineItemOrderedQuantity-FS_POLine.ReceiptQuantity) AS 'Open Qty', FS_POLineData.ItemUnitCost as 'UnitCost', " & _
        '"0 as QtyRecd,0 AS short,'' as Type,'' as Reason, 0 as Amt_INR,0 as CustDuty,0 as FrtCharges,0 as C_And_F, 0 as CST,0 as Packing, 0 as ED_WH,0 as Amt_FC,0 as LnTax_Tot " & _
        '"FROM FSDBBR.dbo.FS_Item FS_Item, FSDBBR.dbo.FS_POHeader FS_POHeader, FSDBBR.dbo.FS_POLine FS_POLine, FSDBBR.dbo.FS_POLineData FS_POLineData," & _
        '"FSDBBR.dbo.FS_Vendor WHERE FS_Item.ItemKey = FS_POLine.ItemKey And FS_POLine.POHeaderKey = FS_POHeader.POHeaderKey And FS_POHeader.VendorKey = FS_Vendor.VendorKey " & _
        '"AND FS_POLine.POLineKey = FS_POLineData.POLineKey AND FS_POLine.POLineStatus <= '4' " & _
        '"AND  FS_POHeader.PONumber IN (SELECT a.PONumber FROM FSPrograms.dbo.PORV_POs a where a.PONumber = FS_POHeader.PONumber and a.PORV_ID = " & txtporvid.Text & " ) " & _
        '" order by FS_POHeader.PONumber, FS_POLine.POLineNumber"


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        ' cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewPODetails.DataSource = stockDC.Tables(0)

      
        Dim col0 As DataGridViewColumn = DataGridViewPODetails.Columns(0) 'SElection
        col0.Width = 30
        DataGridViewPODetails.Columns(0).ReadOnly = False

        Dim col1 As DataGridViewColumn = DataGridViewPODetails.Columns(1) 'PO
        col1.Width = 85
        DataGridViewPODetails.Columns(1).ReadOnly = True
        DataGridViewPODetails.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Dim col2 As DataGridViewColumn = DataGridViewPODetails.Columns(2) 'LN
        col2.Width = 30
        DataGridViewPODetails.Columns(2).ReadOnly = True

        Dim col3 As DataGridViewColumn = DataGridViewPODetails.Columns(3) 'ItemNumber
        col3.Width = 130
        DataGridViewPODetails.Columns(3).ReadOnly = True
        DataGridViewPODetails.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Dim col4 As DataGridViewColumn = DataGridViewPODetails.Columns(4) 'itemdesc
        col4.Width = 130
        DataGridViewPODetails.Columns(4).ReadOnly = True
        DataGridViewPODetails.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Dim col5 As DataGridViewColumn = DataGridViewPODetails.Columns(5) 'uom
        col5.Width = 40
        DataGridViewPODetails.Columns(5).ReadOnly = True
        DataGridViewPODetails.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Dim col6 As DataGridViewColumn = DataGridViewPODetails.Columns(6) 'open qty
        col6.Width = 68
        DataGridViewPODetails.Columns(6).ReadOnly = True
        DataGridViewPODetails.Columns.Item(6).DefaultCellStyle.Format = "n1"
        DataGridViewPODetails.Columns.Item(6).ValueType = GetType(Double)


        Dim col7 As DataGridViewColumn = DataGridViewPODetails.Columns(7) 'unitcost
        col7.Width = 68
        DataGridViewPODetails.Columns(7).ReadOnly = True
        DataGridViewPODetails.Columns.Item(7).DefaultCellStyle.Format = "n3"
        DataGridViewPODetails.Columns.Item(7).ValueType = GetType(Double)


        Dim col8 As DataGridViewColumn = DataGridViewPODetails.Columns(8) 'qty recd
        col8.Width = 68
        DataGridViewPODetails.Columns(8).ReadOnly = False
        DataGridViewPODetails.Columns.Item(8).DefaultCellStyle.Format = "n1"
        DataGridViewPODetails.Columns.Item(8).ValueType = GetType(Double)
        DataGridViewPODetails.Columns(8).DefaultCellStyle.BackColor = Color.LightSkyBlue

        Dim col9 As DataGridViewColumn = DataGridViewPODetails.Columns(9) 'short
        col9.Width = 50
        DataGridViewPODetails.Columns(9).ReadOnly = False
        DataGridViewPODetails.Columns.Item(9).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(9).ValueType = GetType(Double)
        DataGridViewPODetails.Columns(9).DefaultCellStyle.BackColor = Color.LightSkyBlue


        Dim cmb As New DataGridViewComboBoxColumn()
        cmb.HeaderText = "Type"
        cmb.Name = "cmb"
        cmb.MaxDropDownItems = 4
        cmb.Items.Add("")
        cmb.Items.Add("S")
        cmb.Items.Add("VR")
        'DataGridViewPODetails.Columns.Add(cmb)

        DataGridViewPODetails.Columns.Insert(10, cmb)

        Dim col10 As DataGridViewColumn = DataGridViewPODetails.Columns(10) 'type  
        col10.Width = 55
        DataGridViewPODetails.Columns(10).ReadOnly = False
        DataGridViewPODetails.Columns(10).DefaultCellStyle.BackColor = Color.LightSkyBlue
        DataGridViewPODetails.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft



        Dim Remark As New DataGridViewComboBoxColumn()
        Remark.HeaderText = "Remarks"
        Remark.Name = "Remark"
        Remark.MaxDropDownItems = 4
        Remark.Items.Add("Short Supply")
        Remark.Items.Add("Return to Vendor")
        'DataGridViewPODetails.Columns.Add(cmb)

        DataGridViewPODetails.Columns.Insert(11, Remark)

        Dim col11 As DataGridViewColumn = DataGridViewPODetails.Columns(11) 'reason
        col11.Width = 130
        DataGridViewPODetails.Columns(11).ReadOnly = False
        DataGridViewPODetails.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridViewPODetails.Columns(11).DefaultCellStyle.BackColor = Color.LightSkyBlue


        'DataGridViewPODetails.Columns.Item(10).DefaultCellStyle.Format = "n2"
        'DataGridViewPODetails.Columns.Item(7).ValueType = GetType(Double)
        'DataGridViewPODetails.Columns(7).DefaultCellStyle.BackColor = Color.LightSkyBlue

        Dim col12 As DataGridViewColumn = DataGridViewPODetails.Columns(12) '
        col12.Width = 68
        DataGridViewPODetails.Columns(12).ReadOnly = True
        DataGridViewPODetails.Columns.Item(12).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(12).ValueType = GetType(Double)


        Dim col13 As DataGridViewColumn = DataGridViewPODetails.Columns(13) 'recd amt
        col13.Width = 68
        DataGridViewPODetails.Columns(13).ReadOnly = True
        DataGridViewPODetails.Columns.Item(13).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(13).ValueType = GetType(Double)

        Dim col14 As DataGridViewColumn = DataGridViewPODetails.Columns(14) 'custom dy
        col14.Width = 68
        DataGridViewPODetails.Columns(14).ReadOnly = True
        DataGridViewPODetails.Columns.Item(14).DefaultCellStyle.Format = "n1"
        DataGridViewPODetails.Columns.Item(14).ValueType = GetType(Double)


        Dim col15 As DataGridViewColumn = DataGridViewPODetails.Columns(15) 'frt
        col15.Width = 68
        DataGridViewPODetails.Columns(15).ReadOnly = True
        DataGridViewPODetails.Columns.Item(15).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(15).ValueType = GetType(Double)


        Dim col16 As DataGridViewColumn = DataGridViewPODetails.Columns(16) 'cand f
        col16.Width = 68
        DataGridViewPODetails.Columns(16).ReadOnly = True
        DataGridViewPODetails.Columns.Item(16).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(16).ValueType = GetType(Double)


        Dim col17 As DataGridViewColumn = DataGridViewPODetails.Columns(17) 'cst
        col17.Width = 68
        DataGridViewPODetails.Columns(17).ReadOnly = True
        DataGridViewPODetails.Columns.Item(17).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(17).ValueType = GetType(Double)


        Dim col18 As DataGridViewColumn = DataGridViewPODetails.Columns(18) 'packing
        col18.Width = 68
        DataGridViewPODetails.Columns(18).ReadOnly = True
        DataGridViewPODetails.Columns.Item(18).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(18).ValueType = GetType(Double)


        Dim col19 As DataGridViewColumn = DataGridViewPODetails.Columns(19) 'ed-wh
        col19.Width = 68
        DataGridViewPODetails.Columns(19).ReadOnly = True
        DataGridViewPODetails.Columns.Item(19).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(19).ValueType = GetType(Double)

        Dim col20 As DataGridViewColumn = DataGridViewPODetails.Columns(20) 'line total
        col20.Width = 68
        DataGridViewPODetails.Columns(20).ReadOnly = True
        DataGridViewPODetails.Columns.Item(20).DefaultCellStyle.Format = "n2"
        DataGridViewPODetails.Columns.Item(20).ValueType = GetType(Double)


    End Sub

    
    Private Sub TextBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVendId.Leave

        If Len(txtVendId.Text) > 1 Then

            Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL1 As SqlCommand
            Dim drSQL1 As SqlDataReader
            Dim strSQL1 As String

            txtVendId.Text = UCase(txtVendId.Text)

            strSQL1 = "SELECT VendorID,VendorName,VendorCountry,VendorCurrencyCode FROM FSDBBR.dbo._NoLock_FS_Vendor WHERE VendorID = '" & txtVendId.Text & "'"
            cnSQL1.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
            drSQL1 = cmSQL1.ExecuteReader()


            If drSQL1.Read() Then

                If IsDBNull(drSQL1.Item(0)) Then
                    MsgBox("Invalid vendor id", vbInformation)
                    Exit Sub
                Else
                    txtvendName.Text = drSQL1.Item(1)
                    txtVendCountry.Text = drSQL1.Item(2)

                    If drSQL1.Item(3) = "00000" Then
                        txtvendcurrency.Text = "INR"
                    Else
                        txtvendcurrency.Text = drSQL1.Item(3)
                    End If
                End If


            Else
                MsgBox("Invalid vendor id", vbInformation)
                Exit Sub

            End If

        End If


    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVendId.TextChanged

    End Sub

    Private Sub BtnVendorId_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnVendorId.Click


        GroupBoxVendorList.Visible = True
        DataGridViewVendorList.Visible = True

        ' GroupBoxVendorList.Location = New System.Drawing.Point(329, 26)


        DataGridViewVendorList.Width = 708
        DataGridViewVendorList.Height = 362

        GroupBoxVendorList.Width = 757
        GroupBoxVendorList.Height = 401


        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet

        txtVendId.Text = txtVendId.Text + "%"

        strSQL = "SELECT VendorID,VendorName,VendorCountry,VendorCurrencyCode FROM FSDBBR.dbo._NoLock_FS_Vendor WHERE VendorID LIKE '" & txtVendId.Text & "' order by VendorID"
        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewVendorList.DataSource = stockDC.Tables(0)

        'cnSQL.Close()

        Dim col0 As DataGridViewColumn = DataGridViewVendorList.Columns(0) 'vendor id
        col0.Width = 60
        
        Dim col1 As DataGridViewColumn = DataGridViewVendorList.Columns(1) 'vendorname
        col1.Width = 300


        Dim col2 As DataGridViewColumn = DataGridViewVendorList.Columns(2) 'PO
        col2.Width = 180

        Dim col3 As DataGridViewColumn = DataGridViewVendorList.Columns(3) 'PO
        col3.Width = 100

    End Sub

    Private Sub DataGridViewVendorList_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridViewVendorList_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        txtVendId.Text = DataGridViewVendorList.CurrentRow.Cells(0).Value.ToString
        txtvendName.Text = DataGridViewVendorList.CurrentRow.Cells(1).Value.ToString
        txtVendCountry.Text = DataGridViewVendorList.CurrentRow.Cells(2).Value.ToString

        If DataGridViewVendorList.CurrentRow.Cells(3).Value.ToString = "00000" Then
            txtvendcurrency.Text = "INR"
        Else
            txtvendcurrency.Text = DataGridViewVendorList.CurrentRow.Cells(3).Value.ToString
        End If


        DataGridViewVendorList.Visible = False


    End Sub

    Public Sub porvidgen()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(PORV_ID)from PORV_POs"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtporvid.Text = 1
            Else
                txtporvid.Text = drSQL1.Item(0) + 1
            End If
        Else
            txtporvid.Text = 1

        End If
        cnSQL1.Close()

    End Sub

    Public Sub porvsubidgen()

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        strSQL1 = "select max(PORV_SUB_ID)from PORV_PorvItemDetails"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                txtporvsubid.Text = 1
            Else
                txtporvsubid.Text = (drSQL1.Item(0) + 1)
            End If
        Else
            txtporvsubid.Text = 1

        End If
        cnSQL1.Close()

    End Sub


    Private Sub ButtonOpenPOs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' DataGridViewAddCharges.ReadOnly = False
        ' DataGridViewAddCharges.Columns(0).ReadOnly = True





        'Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim strSQL As String

        'DataGridViewPOs.Columns.Clear()

        'Dim AlarmColumn1 As New DataGridViewCheckBoxColumn(False)

        'AlarmColumn1.Name = "Sel"
        'AlarmColumn1.HeaderText = "Selection"
        'AlarmColumn1.ReadOnly = False


        'DataGridViewPOs.Columns.Add(AlarmColumn1)
        'DataGridViewPOs.ReadOnly = False




        'Dim stockDC As DataSet = New DataSet

        'strSQL = "Select DISTINCT(FS_POHeader.PONumber)FROM FSDBBR.dbo.FS_POHeader FS_POHeader, FSDBBR.dbo.FS_POLine FS_POLine, " & _
        '"FSDBBR.dbo.FS_Vendor WHERE  FS_POLine.POHeaderKey = FS_POHeader.POHeaderKey And FS_POHeader.VendorKey = FS_Vendor.VendorKey " & _
        '"AND FS_POLine.POLineStatus <= '4' AND FS_Vendor.VendorID = '" & txtVendId.Text & "' order  by FS_POHeader.PONumber "


        'Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        'Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        'stockDAC.SelectCommand = sqlCmd
        'cnSQL.Open()

        'stockDAC.TableMappings.Add("Table", "Enq")
        ''get data
        'stockDAC.Fill(stockDC)


        'DataGridViewPOs.DataSource = stockDC.Tables(0)
        ''DataGridViewPOs.Columns(0).ReadOnly = True

        'cnSQL.Close()

    End Sub

    Private Sub ButtonOpenPOs_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOpenPOs.Click
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        If txtVendId.Text = "" Then
            MsgBox("Vendor should be selected", vbInformation)
            Exit Sub
        End If




        DataGridViewPOs.Columns.Clear()

        Dim AlarmColumn1 As New DataGridViewCheckBoxColumn(False)

        AlarmColumn1.Name = "S"
        AlarmColumn1.HeaderText = "Sel"
        AlarmColumn1.ReadOnly = False


        DataGridViewPOs.Columns.Add(AlarmColumn1)
        DataGridViewPOs.ReadOnly = False

        Dim stockDC As DataSet = New DataSet

        strSQL = "Select DISTINCT(FS_POHeader.PONumber)FROM FSDBBR.dbo.FS_POHeader FS_POHeader, FSDBBR.dbo.FS_POLine FS_POLine, " & _
        "FSDBBR.dbo.FS_Vendor WHERE  FS_POLine.POHeaderKey = FS_POHeader.POHeaderKey And FS_POHeader.VendorKey = FS_Vendor.VendorKey " & _
        "AND FS_POLine.POLineStatus <= '4' AND FS_Vendor.VendorID = '" & txtVendId.Text & "' order  by FS_POHeader.PONumber "


        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewPOs.DataSource = stockDC.Tables(0)

        DataGridViewPOs.ReadOnly = False

        DataGridViewPOs.Columns(0).ReadOnly = False
        DataGridViewPOs.Columns(1).ReadOnly = True

        cnSQL.Close()

    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click

        Dim msgb As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim strSql As String
        Dim cmSQL As SqlCommand

        TextBoxTOT.Text = 0


        For i As Integer = 0 To DataGridViewAddCharges.RowCount - 1

            If DataGridViewAddCharges.Rows(i).Cells("Amount").Value > 0 Then
                TextBoxTOT.Text = Val(TextBoxTOT.Text) + Val(DataGridViewAddCharges.Rows(i).Cells("Amount").Value)

            End If

        Next


        If Val(TextBoxTOT.Text) > 0 Then


            msgb = MsgBox("Pl check the details entered, are you sure of saving ?", vbYesNo)

            If msgb = vbYes Then


                GroupboxPorvDetails.Visible = False

                GroupBoxPOItemDetails.Location = New Point(12, 59)
                GroupBoxPOItemDetails.Visible = True
                GroupBoxPOItemDetails.Width = 1581
                GroupBoxPOItemDetails.Height = 726


                curdate = System.DateTime.Now()
                cnSQL.Open()

                For i As Integer = 0 To DataGridViewPOs.RowCount - 1

                    If DataGridViewPOs.Rows(i).Cells("S").Value = True Then

                        strSql = "insert PORV_POs  values(" & txtporvid.Text & ",'" & txtVendId.Text & "', '" & curdate & "','" & username & "','" & DataGridViewPOs.Rows(i).Cells("PONumber").Value & "','P')"

                        cmSQL = New SqlCommand(strSql, cnSQL)

                        If cmSQL.ExecuteNonQuery() = 0 Then
                            MsgBox("Cannot save po details" & strSql, MsgBoxStyle.Exclamation, "Error!")
                            Application.Exit()

                        End If
                    End If


                Next

                For i As Integer = 0 To DataGridViewAddCharges.RowCount - 1

                    If DataGridViewAddCharges.Rows(i).Cells("Amount").Value > 0 Then


                        strSql = "insert PORV_DutyDetails values(" & txtporvid.Text & ",'" & txtVendId.Text & "', '" & curdate & "','" & username & "','" & DataGridViewAddCharges.Rows(i).Cells("ID").Value & "','" & DataGridViewAddCharges.Rows(i).Cells("Duty_Tax_Desc").Value & "'," & DataGridViewAddCharges.Rows(i).Cells("Amount").Value & ",'" & txtBOE.Text & "'," & txtInvoiceValue.Text & "," & txtAssVal.Text & ",'P')"

                        cmSQL = New SqlCommand(strSql, cnSQL)

                        If cmSQL.ExecuteNonQuery() = 0 Then
                            MsgBox("Cannot save duty details" & strSql, MsgBoxStyle.Exclamation, "Error!")
                            Application.Exit()

                        End If
                    End If

                Next

            Else
                Exit Sub
            End If

            GroupBoxPOItemDetails.Enabled = True
            GroupBoxHeaderDeails.Enabled = False
            GroupBoxOpenPos.Enabled = False
            GroupBoxDutyDetails.Enabled = False

            Getpodetails()
        Else
            MsgBox("Click on Total text box , to get duty total", vbInformation)
            Exit Sub

        End If

    End Sub

    Private Sub ButtonVendClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonVendClose.Click
        GroupBoxVendorList.Visible = False
    End Sub

    Private Sub TextBoxTOT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxTOT.GotFocus


    End Sub

    Private Sub TextBoxTOT_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBoxTOT.MouseClick

        TextBoxTOT.Text = 0


        For i As Integer = 0 To DataGridViewAddCharges.RowCount - 1

            If DataGridViewAddCharges.Rows(i).Cells("Amount").Value > 0 Then
                TextBoxTOT.Text = Val(TextBoxTOT.Text) + Val(DataGridViewAddCharges.Rows(i).Cells("Amount").Value)

            End If
         
        Next


    End Sub

    Private Sub TextBoxTOT_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxTOT.TextChanged

    End Sub

    Private Sub DataGridViewPOs_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewPOs.CellContentClick

    End Sub

    Private Sub ButtonPODetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxPOItemDetails.Enter

    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQtyTot.TextChanged

    End Sub

    Private Sub GroupBoxHeaderDeails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxHeaderDeails.Enter

    End Sub

    Private Sub DataGridViewPODetails_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub ButtonCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalculate.Click
        Totaltxtclear()
        ButtonLCMSaving.Enabled = True
        Dim a As Integer
        a = 0
        txtMoreCount.Text = 0

        If Val(txtExchangeRate.Text) > 0 Then

            For i As Integer = 0 To DataGridViewPODetails.RowCount - 1

                ' DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value = ((Val(txtExchangeRate.Text) * DataGridViewPODetails.Rows(i).Cells("UnitCost").Value) * DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value)

                If DataGridViewPODetails.Rows(i).Cells("Short").Value > 0 Then
                    'If DataGridViewPODetails.Rows(i).Cells("cmb").ToString <> "V" Or DataGridViewPODetails.Rows(i).Cells("cmb").ToString <> "S" Then

                    'If (DataGridViewPODetails.Rows(i).Cells("cmb").ToString = "S") Or (DataGridViewPODetails.Rows(i).Cells("cmb").ToString = "V") Then
                    If Len(DataGridViewPODetails.Rows(i).Cells("cmb").Value) >= 1 Then

                    Else

                        MsgBox("Type needs to be selected, when short qty is more than zero", vbInformation)
                        Exit Sub

                    End If

                    If Len(DataGridViewPODetails.Rows(i).Cells("Remark").Value) < 3 Then
                        MsgBox("Reason needs to be selected, when short qty is more than zero", vbInformation)
                        Exit Sub


                    End If
                End If


                txtQtyTot.Text = Val(txtQtyTot.Text) + Val(DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value)


                txtShortTotal.Text = Val(txtShortTotal.Text) + Val(DataGridViewPODetails.Rows(i).Cells("Short").Value)



                DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value = ((Val(txtExchangeRate.Text) * DataGridViewPODetails.Rows(i).Cells("UnitCost").Value) * (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value + DataGridViewPODetails.Rows(i).Cells("Short").Value))
                txtTotAmt.Text = Val(txtTotAmt.Text) + Val(DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value)


            Next

            If Val(txtQtyTot.Text) > 0 Then


                For i As Integer = 0 To DataGridViewPODetails.RowCount - 1

                    DataGridViewPODetails.Rows(i).Cells("CustDuty").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(0).Cells("Amount").Value) / Val(txtTotAmt.Text)
                    txtCustTot.Text = Val(txtCustTot.Text) + Val(DataGridViewPODetails.Rows(i).Cells("CustDuty").Value)

                    DataGridViewPODetails.Rows(i).Cells("FrtCharges").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(1).Cells("Amount").Value) / Val(txtTotAmt.Text)
                    txtFrtTot.Text = Val(txtFrtTot.Text) + Val(DataGridViewPODetails.Rows(i).Cells("FrtCharges").Value)

                    DataGridViewPODetails.Rows(i).Cells("C_And_F").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(2).Cells("Amount").Value) / Val(txtTotAmt.Text)
                    TXTC_F.Text = Val(TXTC_F.Text) + Val(DataGridViewPODetails.Rows(i).Cells("C_And_F").Value)

                    DataGridViewPODetails.Rows(i).Cells("CST").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(3).Cells("Amount").Value) / Val(txtTotAmt.Text)
                    txtCST.Text = Val(txtCST.Text) + Val(DataGridViewPODetails.Rows(i).Cells("CST").Value)

                    DataGridViewPODetails.Rows(i).Cells("Packing").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(4).Cells("Amount").Value) / Val(txtTotAmt.Text)
                    TXTPacking.Text = Val(TXTPacking.Text) + Val(DataGridViewPODetails.Rows(i).Cells("Packing").Value)

                    DataGridViewPODetails.Rows(i).Cells("ED_WH").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(5).Cells("Amount").Value) / Val(txtTotAmt.Text)
                    txtED.Text = Val(txtED.Text) + Val(DataGridViewPODetails.Rows(i).Cells("ED_WH").Value)

                    DataGridViewPODetails.Rows(i).Cells("Amt_FC").Value = (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value * DataGridViewPODetails.Rows(i).Cells("UnitCost").Value)
                    txtInvTotFC.Text = Val(txtInvTotFC.Text) + Val(DataGridViewPODetails.Rows(i).Cells("Amt_FC").Value)

                    DataGridViewPODetails.Rows(i).Cells("LnTax_Tot").Value = DataGridViewPODetails.Rows(i).Cells("CustDuty").Value + DataGridViewPODetails.Rows(i).Cells("FrtCharges").Value + DataGridViewPODetails.Rows(i).Cells("C_And_F").Value + DataGridViewPODetails.Rows(i).Cells("CST").Value + DataGridViewPODetails.Rows(i).Cells("Packing").Value + DataGridViewPODetails.Rows(i).Cells("ED_WH").Value
                    txtLineTaxTot.Text = Val(txtLineTaxTot.Text) + Val(DataGridViewPODetails.Rows(i).Cells("LnTax_Tot").Value)

                    'For i As Integer = 0 To DataGridViewProjectMasterList.Rows.Count - 1
                    If DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value + DataGridViewPODetails.Rows(i).Cells("Short").Value > DataGridViewPODetails.Rows(i).Cells("Open Qty").Value Then
                        DataGridViewPODetails.Rows(i).Cells("QtyRecd").Style.BackColor = Color.Red
                        a = a + 1
                    End If
                    'Next

                    If (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value + DataGridViewPODetails.Rows(i).Cells("Short").Value) < DataGridViewPODetails.Rows(i).Cells("Open Qty").Value And (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value) <> 0 Then
                        DataGridViewPODetails.Rows(i).Cells("QtyRecd").Style.BackColor = Color.Yellow

                    End If

                    If (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value + DataGridViewPODetails.Rows(i).Cells("Short").Value) = DataGridViewPODetails.Rows(i).Cells("Open Qty").Value And (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value) <> 0 Then
                        DataGridViewPODetails.Rows(i).Cells("QtyRecd").Style.BackColor = Color.LightSkyBlue

                    End If

                    If (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value) = 0 Then
                        DataGridViewPODetails.Rows(i).Cells("QtyRecd").Style.BackColor = Color.LightSkyBlue

                    End If

                    txtMoreCount.Text = a
                Next

            Else
                MsgBox("Qty Recd to be entered ", vbInformation)
                Exit Sub
            End If

        Else
            MsgBox(" Pl enter exchange rate", vbInformation)
            Exit Sub

        End If
        ButtonLCMSaving.Enabled = True


    End Sub

    Private Sub DataGridViewVendorList_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewVendorList.CellContentClick

    End Sub

    Private Sub DataGridViewVendorList_RowHeaderMouseClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridViewVendorList.RowHeaderMouseClick

        txtVendId.Text = DataGridViewVendorList.CurrentRow.Cells(0).Value.ToString()

        txtvendName.Text = DataGridViewVendorList.CurrentRow.Cells(1).Value.ToString()

        txtVendCountry.Text = DataGridViewVendorList.CurrentRow.Cells(2).Value.ToString()

        txtvendcurrency.Text = DataGridViewVendorList.CurrentRow.Cells(3).Value.ToString()

        GroupBoxVendorList.Visible = False
        'DataGBridViewVendorList.Visible = False


    End Sub

    'Private Sub BtnLCM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'GroupboxPorvDetails.Visible = False


    'GroupBoxPOItemDetails.Location = New Point(276, 59)
    'GroupBoxPOItemDetails.Visible = True
    'GroupBoxPOItemDetails.Width = 1319
    'GroupBoxPOItemDetails.Height = 726


    'End Sub

    Private Sub BtnPorv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'GroupBoxPOItemDetails.Visible = False

        'GroupboxPorvDetails.Location = New Point(276, 59)
        'GroupboxPorvDetails.Visible = True
        'GroupboxPorvDetails.Width = 1319
        'GroupboxPorvDetails.Height = 726

    End Sub

    Private Sub BtnLCM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLCM.Click

        GroupboxPorvDetails.Visible = False
        GroupBoxPOItemDetails.Location = New Point(12, 59)
        GroupBoxPOItemDetails.Visible = True
        GroupBoxPOItemDetails.Width = 1581
        GroupBoxPOItemDetails.Height = 726
        GroupBoxPOItemDetails.Enabled = True

    End Sub

    Private Sub BtnPorv_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPorv.Click
        'check lcm saved or not
        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String

        BtnLCM.Enabled = True


        strSQL1 = "SELECT PORV_ID from PORV_LCMDetails  WHERE PORV_ID = '" & txtporvid.Text & "'"
        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                MsgBox("LCM records to be saved before moving to PROV !! ", vbInformation)
                Exit Sub
            Else
               
            End If


        Else
            MsgBox("LCM records to be saved before moving to PROV !! ", vbInformation)
            Exit Sub

        End If
        'end of check lcm saved or not
        cnSQL1.Close()


        Dim count1 As Integer = 0
       
        'If CheckBoxSelAll.Checked = False Then
        'count1 = 1
        'End If

        'check any items are checked

        '  If DataGridViewPODetails.Columns(0).Name = "Sel" Then
        '   Dim count1 As Integer = 0
        ' For Each row As DataGridViewRow In DataGridViewPODetails.Rows
        porvsubidgen()
        cnSQL1.Open()
        For i As Integer = 0 To DataGridViewPODetails.RowCount - 1

            If Not IsDBNull(DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value) Then

                If DataGridViewPODetails.Rows(i).Cells("Sel").Value IsNot Nothing And DataGridViewPODetails.Rows(i).Cells("Sel").Value = True And DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value > 0 Then

                    count1 += 1
                    'saving records to porv
                    'DataGridViewPODetails
                    strSQL1 = "insert PORV_PorvItemDetails  values(" & txtporvid.Text & ",'" & DTPPorvDate.Value & "'," & _
                                        "" & txtporvsubid.Text & ", '" & curdate & "'," & _
                                        "'" & DataGridViewPODetails.Rows(i).Cells("PONumber").Value & "', " & _
                                        " " & DataGridViewPODetails.Rows(i).Cells("Ln").Value & ", " & _
                                        "'', " & _
                                        "'" & DataGridViewPODetails.Rows(i).Cells("ItemNumber").Value & "', " & _
                                        "0, " & _
                                        " " & DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value & ", " & _
                                        " '','','',''," & _
                                        "  0 , " & _
                                        " '','','',''," & _
                                        " 0, " & _
                                        " '','','','', " & _
                                        " '',0,'','','',''," & _
                                        " 0,0,0,0,0, " & _
                                        " ''," & _
                                        " '" & username & "', " & _
                                        "'" & curdate & "', " & _
                                        "'" & curdate & "',0) "

                    ' cnSQL1.Open()
                    cmSQL1 = New SqlCommand(strSQL1, cnSQL1)

                    If cmSQL1.ExecuteNonQuery() = 0 Then
                        MsgBox("Cannot save line details !" & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                        Exit Sub
                    End If

                    '--------------------------------------
                End If
            End If

        Next

        'End If

        If count1 = 0 Then
            MsgBox("You need to select the records before  moving to PROV !! ", vbInformation)
            Exit Sub

        End If

        'End If



        GroupBoxPOItemDetails.Visible = False

        GroupboxPorvDetails.Location = New Point(12, 59)
        GroupboxPorvDetails.Visible = True
        GroupboxPorvDetails.Enabled = True
        GroupboxPorvDetails.Width = 1581
        GroupboxPorvDetails.Height = 726
        DataGridViewPorvDetails.ScrollBars = ScrollBars.Both

        PORVDetails()


    End Sub

    'Private Sub DataGridViewPODetails_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewPODetails.CellContentClick

    'End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PORVDetails()
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim strSql As String
        '   Dim cmSQL As SqlCommand
        curdate = System.DateTime.Now()
        cnSQL.Open()

        Dim stockDC As DataSet = New DataSet

       
        strSql = "SELECT PONumber, Ln,ItemNumber, 0 as AssVal,PorvQty1 AS Qty1,'' AS BIN1," & _
        "0 as Qty2,'' AS BIN2,0 as Qty3,'' AS BIN3,''as Sup_Inv_Dt,'' as Sup_Inv_No,'' as Lot_No,'' as CureDt, TariffNo as Tariff_No,10 as 'D%', " & _
        "0 as BCD,0 as 'CVD-12%', 0 AS 'EC-2%', 0 AS 'HEC-1%', 0 as 'SAD-4%' from TSS_PORV_ItemDetails_Tariff WHERE PORV_ID = " & txtporvid.Text & " AND PORV_SUB_ID = " & txtporvsubid.Text & " order by PONumber, Ln "

        'lin1 -5 items
        'line2 - 15 items
        'line3 - 20


        Dim sqlCmd As SqlCommand = New SqlCommand(strSql, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd

        stockDAC.TableMappings.Add("Table", "Enq")
        stockDAC.Fill(stockDC)

        DataGridViewPorvDetails.DataSource = stockDC.Tables(0)
        cnSQL.Close()

        Dim col0 As DataGridViewColumn = DataGridViewPorvDetails.Columns(0) 'PO
        col0.Width = 80
        DataGridViewPorvDetails.Columns(0).ReadOnly = True
        DataGridViewPorvDetails.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Dim col1 As DataGridViewColumn = DataGridViewPorvDetails.Columns(1) 'LN
        col1.Width = 25
        DataGridViewPorvDetails.Columns(1).ReadOnly = True


        Dim RT As New DataGridViewComboBoxColumn()  'RTYPE
        RT.HeaderText = "RT"
        RT.Name = "RT"
        'stk.MaxDropDownItems = 4
        RT.Items.Add("R")
        RT.Items.Add("V")
        DataGridViewPorvDetails.Columns.Insert(2, RT)
        RT.DefaultCellStyle.NullValue = RT.Items(0)
        Dim col2 As DataGridViewColumn = DataGridViewPorvDetails.Columns(2) 'rtype
        col2.Width = 40
        DataGridViewPorvDetails.Columns(2).ReadOnly = False


        Dim col3 As DataGridViewColumn = DataGridViewPorvDetails.Columns(3) 'ItemNumber
        col3.Width = 125
        DataGridViewPorvDetails.Columns(3).ReadOnly = True
        DataGridViewPorvDetails.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Dim col4 As DataGridViewColumn = DataGridViewPorvDetails.Columns(4) 'AssValue
        col4.Width = 90
        DataGridViewPorvDetails.Columns(4).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(4).DefaultCellStyle.Format = "n2"
        '        DataGridViewPorvDetails.Columns.Item(4).ValueType = GetType(Double)

        Dim col5 As DataGridViewColumn = DataGridViewPorvDetails.Columns(5) 'Qty1
        col5.Width = 60
        DataGridViewPorvDetails.Columns(5).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(5).DefaultCellStyle.Format = "n1"
        DataGridViewPorvDetails.Columns.Item(5).ValueType = GetType(Double)



        Dim stk1 As New DataGridViewComboBoxColumn() 'STK1
        stk1.HeaderText = "SK1"
        stk1.Name = "SK1"
        ' stk1.MaxDropDownItems = 4
        stk1.Items.Add("WH")
        stk1.Items.Add("K6")
        stk1.Items.Add("K7")
        DataGridViewPorvDetails.Columns.Insert(6, stk1)
        Dim col6 As DataGridViewColumn = DataGridViewPorvDetails.Columns(6) 'stk1
        col6.Width = 50
        DataGridViewPorvDetails.Columns(6).ReadOnly = False
        DataGridViewPorvDetails.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Dim col7 As DataGridViewColumn = DataGridViewPorvDetails.Columns(7) 'BIN1
        col7.Width = 60
        DataGridViewPorvDetails.Columns(7).ReadOnly = False
        DataGridViewPorvDetails.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Dim IC1 As New DataGridViewComboBoxColumn() 'IC1
        IC1.HeaderText = "IC1"
        IC1.Name = "IC1"

        ' stk1.MaxDropDownItems = 4
        IC1.Items.Add("O")
        IC1.Items.Add("I")
        IC1.Items.Add("H")
        DataGridViewPorvDetails.Columns.Insert(8, IC1)
        IC1.DefaultCellStyle.NullValue = IC1.Items(0)
        Dim col8 As DataGridViewColumn = DataGridViewPorvDetails.Columns(8) 'IC1
        col8.Width = 40
        DataGridViewPorvDetails.Columns(8).ReadOnly = False

        '---ADDING COMBO BOX AND SELECTING DEFAULT VALUE IMPORTANT CODE
        'Dim comboBox As New DataGridViewComboBoxColumn()

        ''Add some stuff to the combobox
        'comboBox.Items.Add("FirstItem")
        'comboBox.Items.Add("SecondItem")

        ''Select the first item
        'comboBox.DefaultCellStyle.NullValue = comboBox.Items(0)

        ''Now add the whole combobox to the DataGridView
        'dgvItems.Columns.Add(comboBox)

        '---

        Dim C1 As New DataGridViewComboBoxColumn() 'C1
        C1.HeaderText = "C1"
        C1.Name = "C1"
        ' stk1.MaxDropDownItems = 4
        C1.Items.Add("G")
        C1.Items.Add("R")
        DataGridViewPorvDetails.Columns.Insert(9, C1)
        C1.DefaultCellStyle.NullValue = C1.Items(0)
        Dim col9 As DataGridViewColumn = DataGridViewPorvDetails.Columns(9) 'code1
        col9.Width = 40
        DataGridViewPorvDetails.Columns(9).ReadOnly = False


        Dim col10 As DataGridViewColumn = DataGridViewPorvDetails.Columns(10) 'qty2
        col10.HeaderText = "QTY2"
        col10.Width = 60
        DataGridViewPorvDetails.Columns(10).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(10).DefaultCellStyle.Format = "n2"
        DataGridViewPorvDetails.Columns.Item(10).ValueType = GetType(Double)


        Dim stk2 As New DataGridViewComboBoxColumn() 'STK2
        stk2.HeaderText = "SK2"
        stk2.Name = "SK2"
        ' stk1.MaxDropDownItems = 4
        stk2.Items.Add("WH")
        stk2.Items.Add("K6")
        stk2.Items.Add("K7")
        DataGridViewPorvDetails.Columns.Insert(11, stk2)
        Dim col11 As DataGridViewColumn = DataGridViewPorvDetails.Columns(11) 'stk2
        col11.Width = 50
        DataGridViewPorvDetails.Columns(11).ReadOnly = False
        DataGridViewPorvDetails.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Dim col12 As DataGridViewColumn = DataGridViewPorvDetails.Columns(12) 'bin2
        col12.Width = 60
        DataGridViewPorvDetails.Columns(12).ReadOnly = False


        Dim IC2 As New DataGridViewComboBoxColumn() 'IC2
        IC2.HeaderText = "IC2"
        IC2.Name = "IC2"
        ' stk1.MaxDropDownItems = 4
        IC2.Items.Add("O")
        IC2.Items.Add("I")
        IC2.Items.Add("H")
        DataGridViewPorvDetails.Columns.Insert(13, IC2)

        Dim col13 As DataGridViewColumn = DataGridViewPorvDetails.Columns(13) 'IC2
        col13.Width = 40
        DataGridViewPorvDetails.Columns(13).ReadOnly = False

        Dim C2 As New DataGridViewComboBoxColumn() 'C2
        C2.HeaderText = "C2"
        C2.Name = "C2"
        ' stk1.MaxDropDownItems = 4E
        C2.Items.Add("G")
        C2.Items.Add("R")
        DataGridViewPorvDetails.Columns.Insert(14, C2)
        Dim col14 As DataGridViewColumn = DataGridViewPorvDetails.Columns(14) 'code1
        col14.Width = 40
        DataGridViewPorvDetails.Columns(14).ReadOnly = False


        Dim col15 As DataGridViewColumn = DataGridViewPorvDetails.Columns(15) 'qty3
        col15.Width = 30
        DataGridViewPorvDetails.Columns(15).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(15).DefaultCellStyle.Format = "n1"
        DataGridViewPorvDetails.Columns.Item(15).ValueType = GetType(Double)


        Dim stk3 As New DataGridViewComboBoxColumn()
        stk3.HeaderText = "SK3"
        stk3.Name = "SK3"
        stk3.Items.Add("WH")
        stk3.Items.Add("K6")
        stk3.Items.Add("K7")
        DataGridViewPorvDetails.Columns.Insert(16, stk3)
        Dim col16 As DataGridViewColumn = DataGridViewPorvDetails.Columns(16) 'stk3
        col16.Width = 50
        DataGridViewPorvDetails.Columns(16).ReadOnly = False


        Dim col17 As DataGridViewColumn = DataGridViewPorvDetails.Columns(17) 'bin3
        col17.Width = 50
        DataGridViewPorvDetails.Columns(17).ReadOnly = False

        Dim IC3 As New DataGridViewComboBoxColumn() 'IC3
        IC3.HeaderText = "IC3"
        IC3.Name = "IC3"
        ' stk1.MaxDropDownItems = 4
        IC3.Items.Add("O")
        IC3.Items.Add("I")
        IC3.Items.Add("H")
        DataGridViewPorvDetails.Columns.Insert(18, IC3)
        Dim col18 As DataGridViewColumn = DataGridViewPorvDetails.Columns(18) 'IC3
        col18.Width = 40
        DataGridViewPorvDetails.Columns(18).ReadOnly = False

        Dim C3 As New DataGridViewComboBoxColumn() 'C3
        C3.HeaderText = "C3"
        C3.Name = "C3"
        C3.Items.Add("G")
        C3.Items.Add("R")
        DataGridViewPorvDetails.Columns.Insert(19, C3)
        Dim col19 As DataGridViewColumn = DataGridViewPorvDetails.Columns(19) 'IC3
        col19.Width = 40
        DataGridViewPorvDetails.Columns(19).ReadOnly = False



        'Dim col20 As DataGridViewColumn = DataGridViewPorvDetails.Columns(20) 'tariffno
        'col20.Width = 100
        'DataGridViewPorvDetails.Columns(20).ReadOnly = False
        'DataGridViewPorvDetails.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        'Dim col21 As DataGridViewColumn = DataGridViewPorvDetails.Columns(21) 'duty%
        'col21.Width = 35
        'DataGridViewPorvDetails.Columns(21).ReadOnly = False
        'DataGridViewPorvDetails.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'DataGridViewPorvDetails.Columns.Item(21).DefaultCellStyle.Format = "n1"
        'DataGridViewPorvDetails.Columns.Item(21).ValueType = GetType(Double)


        Dim col20 As DataGridViewColumn = DataGridViewPorvDetails.Columns(20) 'sup inv date
        col20.Width = 70
        DataGridViewPorvDetails.Columns(20).ReadOnly = False
        DataGridViewPorvDetails.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Dim col21 As DataGridViewColumn = DataGridViewPorvDetails.Columns(21) 'sup inv no
        col21.Width = 100
        DataGridViewPorvDetails.Columns(21).ReadOnly = False
        DataGridViewPorvDetails.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Dim col22 As DataGridViewColumn = DataGridViewPorvDetails.Columns(22) 'lot no
        col22.Width = 110
        DataGridViewPorvDetails.Columns(22).ReadOnly = False
        DataGridViewPorvDetails.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Dim col23 As DataGridViewColumn = DataGridViewPorvDetails.Columns(23) 'cure date
        col23.Width = 50
        DataGridViewPorvDetails.Columns(23).ReadOnly = False
        DataGridViewPorvDetails.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Dim col24 As DataGridViewColumn = DataGridViewPorvDetails.Columns(24) 'tariffno
        col24.Width = 100
        DataGridViewPorvDetails.Columns(24).ReadOnly = False
        DataGridViewPorvDetails.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Dim col25 As DataGridViewColumn = DataGridViewPorvDetails.Columns(25) 'duty%
        col25.Width = 35
        DataGridViewPorvDetails.Columns(25).ReadOnly = False
        DataGridViewPorvDetails.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridViewPorvDetails.Columns.Item(25).DefaultCellStyle.Format = "n1"
        DataGridViewPorvDetails.Columns.Item(25).ValueType = GetType(Double)


        Dim col26 As DataGridViewColumn = DataGridViewPorvDetails.Columns(26) 'bcd
        col26.Width = 90
        DataGridViewPorvDetails.Columns(26).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(26).DefaultCellStyle.Format = "n2"
        DataGridViewPorvDetails.Columns.Item(26).ValueType = GetType(Double)


        Dim col27 As DataGridViewColumn = DataGridViewPorvDetails.Columns(27) 'cvd
        col27.Width = 90
        DataGridViewPorvDetails.Columns(27).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(27).DefaultCellStyle.Format = "n2"
        DataGridViewPorvDetails.Columns.Item(27).ValueType = GetType(Double)


        Dim col28 As DataGridViewColumn = DataGridViewPorvDetails.Columns(28) 'ec
        col28.Width = 90
        DataGridViewPorvDetails.Columns(28).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(28).DefaultCellStyle.Format = "n2"
        DataGridViewPorvDetails.Columns.Item(28).ValueType = GetType(Double)


        Dim col29 As DataGridViewColumn = DataGridViewPorvDetails.Columns(29) 'hec
        col29.Width = 90
        DataGridViewPorvDetails.Columns(29).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(29).DefaultCellStyle.Format = "n2"
        DataGridViewPorvDetails.Columns.Item(29).ValueType = GetType(Double)


        Dim col30 As DataGridViewColumn = DataGridViewPorvDetails.Columns(30) 'sad
        col30.Width = 90
        DataGridViewPorvDetails.Columns(30).ReadOnly = False
        DataGridViewPorvDetails.Columns.Item(30).DefaultCellStyle.Format = "n2"
        DataGridViewPorvDetails.Columns.Item(30).ValueType = GetType(Double)


        Dim LC As New DataGridViewComboBoxColumn() 'Line closure
        LC.HeaderText = "LC"
        LC.Name = "LC"
        LC.Items.Add("N")
        LC.Items.Add("Y")
        DataGridViewPorvDetails.Columns.Insert(31, LC)

        LC.DefaultCellStyle.NullValue = LC.Items(0)

        Dim col31 As DataGridViewColumn = DataGridViewPorvDetails.Columns(31) 'LINE CLOSURE
        col31.Width = 40
        DataGridViewPorvDetails.Columns(31).ReadOnly = False

        DataGridViewPorvDetails.Columns(5).Frozen = True


        Porvdutycal() 'TSS_PORV_TariffNos





    End Sub

    'Private Sub ButtonLcmSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLcmSave.Click

    'Dim msgb As String
    'Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
    'Dim strSql As String
    'Dim cmSQL As SqlCommand


    'BtnPorv.Enabled = True


    'If Val(txtQtyTot.Text) > 0 Then


    '    msgb = MsgBox("Pl check the details entered, are you sure of saving ?", vbYesNo)
    '    If msgb = vbYes Then
    '        cnSQL.Open()
    '        For i As Integer = 0 To DataGridViewPODetails.RowCount - 1

    '            If DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value > 0 Then

    '                '             strSql = "SELECT FS_POHeader.PONumber, FS_POLine.POLineNumber AS 'Ln', FS_Item.ItemNumber, FS_Item.ItemDescription, " & _
    '                '"FS_Item.ItemUM AS 'UOM', (FS_POLine.LineItemOrderedQuantity-FS_POLine.ReceiptQuantity) AS 'Open Qty', FS_POLineData.ItemUnitCost as 'UnitCost', " & _
    '                '"0 as QtyRecd, 0 as Amt_INR,0 as CustDuty,0 as FrtCharges,0 as C_And_F, 0 as CST,0 as Packing, 0 as ED_WH,0 as Amt_FC,0 as LnTax_Tot " & _


    '                strSql = "insert PORV_LCMDetails  values(" & txtporvid.Text & ",'" & txtVendId.Text & "'," & _
    '                "'" & DataGridViewPODetails.Rows(i).Cells("PONumber").Value & "', " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("Ln").Value & ", " & _
    '                "'" & DataGridViewPODetails.Rows(i).Cells("ItemNumber").Value & "', " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("open Qty").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("Short").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("Type").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("Reason").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("UnitCost").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("CustDuty").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("FrtCharges").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("C_And_F").Value & " ," & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("CST").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("Packing").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("ED_WH").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value & ", " & _
    '                " " & DataGridViewPODetails.Rows(i).Cells("Amt_FC").Value & ", " & _
    '                " '" & username & "', " & _
    '                "'" & curdate & "', " & _
    '                "'" & curdate & "',0) "

    '                cmSQL = New SqlCommand(strSql, cnSQL)

    '                If cmSQL.ExecuteNonQuery() = 0 Then
    '                    MsgBox("Cannot save po details !" & strSql, MsgBoxStyle.Exclamation, "Error!")
    '                    Application.Exit()

    '                End If
    '            End If

    '        Next

    '        MsgBox("LCM deatils are saved", vbInformation)
    '        Exit Sub

    '    Else
    '        Exit Sub
    '    End If
    'Else
    '    MsgBox("Qunatities received to be entered before saving !", vbInformation)
    '    Exit Sub
    'End If


    'End Sub

    Private Sub txtInvoiceValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceValue.TextChanged

    End Sub

    Private Sub txtBOE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBOE.TextChanged

    End Sub

    Private Sub txtAssVal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAssVal.TextChanged

    End Sub

    'Private Sub ButtonLCMClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLCMClose.Click
    '    GroupBoxPOItemDetails.Visible = False

    'End Sub

    Private Sub ButtonLCMClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLCMClose.Click

    End Sub


    Private Sub DataGridViewPODetails_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewPODetails.CellContentClick

    End Sub

    Private Sub ButtonLcmSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ButtonLCMSaving_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLCMSaving.Click
        Dim msgb As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSql As String
        Dim cmSQL As SqlCommand

        Dim taxdiff As Double


        ' taxdiff = Val(TextBoxTOT.Text) - (Val(txtCustTot.Text) + Val(txtFrtTot.Text) + Val(TXTC_F.Text) + Val(txtCST.Text) + (TXTPacking.Text))

        taxdiff = ((100 * (Val(TextBoxTOT.Text) - (Val(txtCustTot.Text) + Val(txtFrtTot.Text) + Val(TXTC_F.Text) + Val(txtCST.Text) + (TXTPacking.Text)))) / TextBoxTOT.Text)


        If taxdiff > 0.001 Then


            MsgBox("Mismatch in Duty, Pl check again", vbInformation)
            Exit Sub

        End If




        If Val(txtMoreCount.Text) > 0 Then
            MsgBox("Recd Qty should not be more than open Qty !", vbInformation)
            Exit Sub

        Else
            BtnPorv.Enabled = True


            If Val(txtQtyTot.Text) > 0 Then


                msgb = MsgBox("Pl check the details entered, are you sure of saving ?", vbYesNo)
                If msgb = vbYes Then
                    cnSQL.Open()
                    For i As Integer = 0 To DataGridViewPODetails.RowCount - 1

                        If (Not IsDBNull(DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value)) And (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value) > 0 Then

                            '             strSql = "SELECT FS_POHeader.PONumber, FS_POLine.POLineNumber AS 'Ln', FS_Item.ItemNumber, FS_Item.ItemDescription, " & _
                            '"FS_Item.ItemUM AS 'UOM', (FS_POLine.LineItemOrderedQuantity-FS_POLine.ReceiptQuantity) AS 'Open Qty', FS_POLineData.ItemUnitCost as 'UnitCost', " & _
                            '"0 as QtyRecd, 0 as Amt_INR,0 as CustDuty,0 as FrtCharges,0 as C_And_F, 0 as CST,0 as Packing, 0 as ED_WH,0 as Amt_FC,0 as LnTax_Tot " & _


                            strSql = "insert PORV_LCMDetails  values(" & txtporvid.Text & ",'" & txtVendId.Text & "'," & _
                            "'" & DataGridViewPODetails.Rows(i).Cells("PONumber").Value & "', " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("Ln").Value & ", " & _
                            "'" & DataGridViewPODetails.Rows(i).Cells("ItemNumber").Value & "', " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("open Qty").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("Short").Value & ", " & _
                            " ' " & DataGridViewPODetails.Rows(i).Cells("cmb").Value & "', " & _
                            " '" & DataGridViewPODetails.Rows(i).Cells("Remark").Value & "', " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("UnitCost").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("CustDuty").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("FrtCharges").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("C_And_F").Value & " ," & _
                            " " & DataGridViewPODetails.Rows(i).Cells("CST").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("Packing").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("ED_WH").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value & ", " & _
                            " " & DataGridViewPODetails.Rows(i).Cells("Amt_FC").Value & ", " & _
                            " '" & username & "', " & _
                            "'" & curdate & "', " & _
                            "'" & curdate & "',0) "

                            cmSQL = New SqlCommand(strSql, cnSQL)

                            If cmSQL.ExecuteNonQuery() = 0 Then
                                MsgBox("Cannot save po details !" & strSql, MsgBoxStyle.Exclamation, "Error!")
                                Application.Exit()

                            End If
                        End If

                    Next

                    MsgBox("LCM deatils are saved", vbInformation)
                    BtnPorv.Enabled = True

                    Exit Sub

                Else
                    Exit Sub
                End If
            Else
                MsgBox("Qunatities received to be entered before saving !", vbInformation)
                Exit Sub
            End If
        End If

    End Sub

    Private Sub txtPurchaseInvoiceNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPurchaseInvoiceNumber.TextChanged

    End Sub

    Private Sub ButtonPORVSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPORVSave.Click

    End Sub

    Private Sub GroupBoxHeaderDeails_HandleDestroyed(ByVal sender As Object, ByVal e As System.EventArgs) Handles GroupBoxHeaderDeails.HandleDestroyed

    End Sub

    Private Sub ChkPoCopy_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkPoCopy.CheckedChanged

        If ChkPoCopy.Checked = True Then
            Dim i As Integer = 0
            Dim lotcount As Integer = 1

            For i = 0 To DataGridViewPorvDetails.RowCount - 2
                DataGridViewPorvDetails.Rows(i).Cells(21).Value = txtPurchaseInvoiceNumber.Text
                DataGridViewPorvDetails.Columns(20).DefaultCellStyle.Format = "mm/dd/yy"
                DataGridViewPorvDetails.Rows(i).Cells(20).Value = dtpPurInvDate.Value
                'DataGridViewPorvDetails.Columns(20).DefaultCellStyle.Format = "mm/dd/yy"
                DataGridViewPorvDetails.Rows(i).Cells(22).Value = txtPurchaseInvoiceNumber.Text & -(lotcount)
                lotcount = lotcount + 1

            Next

        ElseIf ChkPoCopy.Checked = False Then
            Dim i As Integer = 0
            For i = 0 To DataGridViewPorvDetails.RowCount - 2
                DataGridViewPorvDetails.Rows(i).Cells(21).Value = ""
                DataGridViewPorvDetails.Rows(i).Cells(20).Value = ""
                DataGridViewPorvDetails.Rows(i).Cells(22).Value = ""
            Next
        End If


    End Sub

    Private Sub Porvdutycal()

        'ButtonLCMSaving.Enabled = True
        'Dim a As Integer
        'a = 0
        'txtMoreCount.Text = 0

        'If Val(txtExchangeRate.Text) > 0 Then

        '    For i As Integer = 0 To DataGridViewPODetails.RowCount - 1


        '        If DataGridViewPODetails.Rows(i).Cells("Short").Value > 0 Then
        '            If Len(DataGridViewPODetails.Rows(i).Cells("cmb").Value) >= 1 Then

        '            Else

        '                MsgBox("Type needs to be selected, when short qty is more than zero", vbInformation)
        '                Exit Sub

        '            End If

        '            If Len(DataGridViewPODetails.Rows(i).Cells("Remark").Value) < 3 Then
        '                MsgBox("Reason needs to be selected, when short qty is more than zero", vbInformation)
        '                Exit Sub


        '            End If
        '        End If


        'txtQtyTot.Text = Val(txtQtyTot.Text) + Val(DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value)


        'txtShortTotal.Text = Val(txtShortTotal.Text) + Val(DataGridViewPODetails.Rows(i).Cells("Short").Value)



        'DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value = ((Val(txtExchangeRate.Text) * DataGridViewPODetails.Rows(i).Cells("UnitCost").Value) * (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value + DataGridViewPODetails.Rows(i).Cells("Short").Value))
        'txtTotAmt.Text = Val(txtTotAmt.Text) + Val(DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value)


        '    Next




        'If Val(txtQtyTot.Text) > 0 Then


        '    For i As Integer = 0 To DataGridViewPODetails.RowCount - 1

        '        DataGridViewPODetails.Rows(i).Cells("CustDuty").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(0).Cells("Amount").Value) / Val(txtTotAmt.Text)
        '        txtCustTot.Text = Val(txtCustTot.Text) + Val(DataGridViewPODetails.Rows(i).Cells("CustDuty").Value)

        '        DataGridViewPODetails.Rows(i).Cells("FrtCharges").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(1).Cells("Amount").Value) / Val(txtTotAmt.Text)
        '        txtFrtTot.Text = Val(txtFrtTot.Text) + Val(DataGridViewPODetails.Rows(i).Cells("FrtCharges").Value)

        '        DataGridViewPODetails.Rows(i).Cells("C_And_F").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(2).Cells("Amount").Value) / Val(txtTotAmt.Text)
        '        TXTC_F.Text = Val(TXTC_F.Text) + Val(DataGridViewPODetails.Rows(i).Cells("C_And_F").Value)

        '        DataGridViewPODetails.Rows(i).Cells("CST").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(3).Cells("Amount").Value) / Val(txtTotAmt.Text)
        '        txtCST.Text = Val(txtCST.Text) + Val(DataGridViewPODetails.Rows(i).Cells("CST").Value)

        '        DataGridViewPODetails.Rows(i).Cells("Packing").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(4).Cells("Amount").Value) / Val(txtTotAmt.Text)
        '        TXTPacking.Text = Val(TXTPacking.Text) + Val(DataGridViewPODetails.Rows(i).Cells("Packing").Value)

        '        DataGridViewPODetails.Rows(i).Cells("ED_WH").Value = (DataGridViewPODetails.Rows(i).Cells("Amt_INR").Value * DataGridViewAddCharges.Rows(5).Cells("Amount").Value) / Val(txtTotAmt.Text)
        '        txtED.Text = Val(txtED.Text) + Val(DataGridViewPODetails.Rows(i).Cells("ED_WH").Value)

        '        DataGridViewPODetails.Rows(i).Cells("Amt_FC").Value = (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value * DataGridViewPODetails.Rows(i).Cells("UnitCost").Value)
        '        txtInvTotFC.Text = Val(txtInvTotFC.Text) + Val(DataGridViewPODetails.Rows(i).Cells("Amt_FC").Value)

        '        DataGridViewPODetails.Rows(i).Cells("LnTax_Tot").Value = DataGridViewPODetails.Rows(i).Cells("CustDuty").Value + DataGridViewPODetails.Rows(i).Cells("FrtCharges").Value + DataGridViewPODetails.Rows(i).Cells("C_And_F").Value + DataGridViewPODetails.Rows(i).Cells("CST").Value + DataGridViewPODetails.Rows(i).Cells("Packing").Value + DataGridViewPODetails.Rows(i).Cells("ED_WH").Value
        '        txtLineTaxTot.Text = Val(txtLineTaxTot.Text) + Val(DataGridViewPODetails.Rows(i).Cells("LnTax_Tot").Value)

        '        'For i As Integer = 0 To DataGridViewProjectMasterList.Rows.Count - 1
        '        If DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value + DataGridViewPODetails.Rows(i).Cells("Short").Value > DataGridViewPODetails.Rows(i).Cells("Open Qty").Value Then
        '            DataGridViewPODetails.Rows(i).Cells("QtyRecd").Style.BackColor = Color.Red
        '            a = a + 1
        '        End If
        '        'Next

        '        If (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value + DataGridViewPODetails.Rows(i).Cells("Short").Value) < DataGridViewPODetails.Rows(i).Cells("Open Qty").Value And (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value) <> 0 Then
        '            DataGridViewPODetails.Rows(i).Cells("QtyRecd").Style.BackColor = Color.Yellow

        '        End If

        '        If (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value + DataGridViewPODetails.Rows(i).Cells("Short").Value) = DataGridViewPODetails.Rows(i).Cells("Open Qty").Value And (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value) <> 0 Then
        '            DataGridViewPODetails.Rows(i).Cells("QtyRecd").Style.BackColor = Color.LightSkyBlue

        '        End If

        '        If (DataGridViewPODetails.Rows(i).Cells("QtyRecd").Value) = 0 Then
        '            DataGridViewPODetails.Rows(i).Cells("QtyRecd").Style.BackColor = Color.LightSkyBlue

        '        End If

        '        txtMoreCount.Text = a
        '    Next

        'Else
        '    MsgBox("Qty Recd to be entered ", vbInformation)
        '    Exit Sub
        'End If

        'Else
        'MsgBox(" Pl enter exchange rate", vbInformation)
        'Exit Sub

        'End If
        'ButtonLCMSaving.Enabled = True


    End Sub

    Private Sub CheckBoxSelAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxSelAll.CheckedChanged
        If CheckBoxSelAll.Checked = True Then
            Dim i As Integer = 0
            For i = 0 To DataGridViewPODetails.RowCount - 1
                DataGridViewPODetails.Rows(i).Cells(0).Value = True
            Next

        Else
            Dim i As Integer = 0
            For i = 0 To DataGridViewPODetails.RowCount - 1
                DataGridViewPODetails.Rows(i).Cells(0).Value = False
            Next
        End If


    End Sub
    Private Sub Totaltxtclear()
        txtLineTaxTot.Text = 0
        txtInvTotFC.Text = 0
        txtED.Text = 0
        TXTPacking.Text = 0
        txtCST.Text = 0
        TXTC_F.Text = 0
        txtFrtTot.Text = 0
        txtCustTot.Text = 0
        txtTotAmt.Text = 0
        txtShortTotal.Text = 0
        txtQtyTot.Text = 0
    End Sub

    Private Sub ButtonDutyCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDutyCal.Click

        '        strSql = "SELECT PONumber, Ln,ItemNumber, 0 as AssVal,PorvQty1 AS Qty1,'' AS BIN1," & _
        '     "0 as Qty2,'' AS BIN2,0 as Qty3,'' AS BIN3, TariffNo as Tariff_No,10 as 'D%',''as Sup_Inv_Dt,'' as Sup_Inv_No,'' as Lot_No,'' as CureDt, " & _
        '    "0 as BCD,0 as CVD, 0 AS EC, 0 AS HEC, 0 as SAD from TSS_PORV_ItemDetails_Tariff WHERE PORV_ID = " & txtporvid.Text & " AND PORV_SUB_ID = " & txtporvsubid.Text & " order by PONumber, Ln "


        Dim i As Integer = 0

        For i = 0 To DataGridViewPorvDetails.RowCount - 1
            DataGridViewPorvDetails.Rows(i).Cells(26).Value = (DataGridViewPorvDetails.Rows(i).Cells(4).Value * DataGridViewPorvDetails.Rows(i).Cells(25).Value) / 100
            DataGridViewPorvDetails.Rows(i).Cells(26).Value = (Math.Round(Int(DataGridViewPorvDetails.Rows(i).Cells(26).Value * 100) / 10, 0) * 0.1)

            DataGridViewPorvDetails.Rows(i).Cells(27).Value = ((DataGridViewPorvDetails.Rows(i).Cells(26).Value + DataGridViewPorvDetails.Rows(i).Cells(4).Value) * 12) / 100
            DataGridViewPorvDetails.Rows(i).Cells(27).Value = Math.Round(Int(DataGridViewPorvDetails.Rows(i).Cells(27).Value * 100) / 10, 0) * 0.1

            DataGridViewPorvDetails.Rows(i).Cells(28).Value = ((DataGridViewPorvDetails.Rows(i).Cells(26).Value + DataGridViewPorvDetails.Rows(i).Cells(27).Value) * 2) / 100
            DataGridViewPorvDetails.Rows(i).Cells(28).Value = Math.Round(Int(DataGridViewPorvDetails.Rows(i).Cells(28).Value * 100) / 10, 0) * 0.1

            DataGridViewPorvDetails.Rows(i).Cells(29).Value = ((DataGridViewPorvDetails.Rows(i).Cells(26).Value + DataGridViewPorvDetails.Rows(i).Cells(27).Value) * 1) / 100
            DataGridViewPorvDetails.Rows(i).Cells(29).Value = Math.Round(Int(DataGridViewPorvDetails.Rows(i).Cells(29).Value * 100) / 10, 0) * 0.1

            DataGridViewPorvDetails.Rows(i).Cells(30).Value = ((DataGridViewPorvDetails.Rows(i).Cells(4).Value + DataGridViewPorvDetails.Rows(i).Cells(26).Value + DataGridViewPorvDetails.Rows(i).Cells(27).Value + DataGridViewPorvDetails.Rows(i).Cells(28).Value + DataGridViewPorvDetails.Rows(i).Cells(29).Value) * 4) / 100
            DataGridViewPorvDetails.Rows(i).Cells(30).Value = Math.Round(Int(DataGridViewPorvDetails.Rows(i).Cells(30).Value * 100) / 10, 0) * 0.1


        Next



    End Sub

    Private Sub DataGridViewAddCharges_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewAddCharges.CellContentClick

    End Sub
End Class