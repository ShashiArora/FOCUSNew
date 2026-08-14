Imports System.Data.SqlClient
'Imports SoftionBrands.FourthShift.Transaction
'Imports Microsoft.Office.Interop.Outlook
'Imports CrystalDecisions.CrystalReports.Engine





Public Class SUBDC
    Dim dupchk As Boolean
    ' Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Private Sub RBNewDC_CheckedChanged(sender As Object, e As EventArgs) Handles RBNewDC.CheckedChanged
        loadvendors()
        txtPONumber.Text = ""
        txtchkcount.Text = ""
        txtdc.Text = ""
        txtRemark2.Text = ""
        txtRemarks1.Text = ""


    End Sub

    Sub dataGridViewscdc_CurrentCellDirtyStateChanged( _
    ByVal sender As Object, ByVal e As EventArgs) _
    Handles DataGridViewSCDC.CurrentCellDirtyStateChanged

        If DataGridViewSCDC.IsCurrentCellDirty Then
            DataGridViewSCDC.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Public Sub dataGridViewscdc_CellValueChanged(ByVal sender As Object, _
    ByVal e As DataGridViewCellEventArgs) _
    Handles DataGridViewSCDC.CellValueChanged

        If DataGridViewSCDC.Columns(e.ColumnIndex).Name = "checkBoxColumn" Then
            Dim count1 As Integer = 0
            For Each row As DataGridViewRow In DataGridViewSCDC.Rows
                If row.Cells("checkBoxColumn").Value IsNot Nothing And row.Cells("checkBoxColumn").Value = True Then
                    count1 += 1
                End If
            Next

            txtchkcount.Text = count1
        End If
    End Sub




    Private Sub SUBDC_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Loadvendors()
    End Sub
    Private Sub loadvendors()


        Dim sqlCon As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim sqlCon1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        Dim strSql As String
        Dim source As DataSet = New DataSet
        Dim cmSQL As SqlCommand

        Dim drSQL As SqlDataReader

        '   strSql = "Select b.VendorKey , a.VendorID + ' - ' +  b.VendorName as 'Vendor Name'  from FSDBBR.dbo.TSS_SubDcOpenPOS a inner join FSDBBR.dbo._NoLock_FS_Vendor b on a.VendorID = b.VendorID  group by a.VendorID + ' - ' +  b.VendorName,b.VendorKey "

        strSql = "Select b.VendorKey ,  b.VendorID + '-' + b.VendorName as 'VendorName' from FSDBBR.dbo.TSS_SubDcOpenPOS a inner join FSDBBR.dbo._NoLock_FS_Vendor b on a.VendorID = b.VendorID  group by  b.VendorID + '-' + b.VendorName,b.VendorKey "

        'modifications  by Indira  on 5TH Aug 25
        'when there are no records, selectedindex=0 is giving error.  So this validation added.

        sqlCon1.Open()
        cmSQL = New SqlCommand(strSql, sqlCon1)
        drSQL = cmSQL.ExecuteReader()

        If drSQL.Read() Then

            If IsDBNull(drSQL.Item(0)) Then

                MsgBox("No data found for DC creation")
                Exit Sub
            Else
                'end of modifications 5th Aug 25

                cmSQL = New SqlCommand(strSql, sqlCon)
                Dim sqlCmd As SqlCommand = New SqlCommand(strSql, sqlCon)
                Dim ESource As SqlDataAdapter = New SqlDataAdapter
                ESource.SelectCommand = sqlCmd
                ESource.Fill(source, "eSource")
                With ComboBoxVendors
                    .DataSource = source.Tables("eSource")
                    .DisplayMember = "VendorName"
                    .ValueMember = "VendorKey"

                    .SelectedIndex = 0

                End With
            End If
        End If


    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs) Handles BtnOK.Click

        If Val(txtchkcount.Text) > 0 Then
            DataGridViewSCDC.Rows.Clear()
        End If


        dupchk = False
        BTNDCSAVE.Enabled = True
        Dim mystr As String
        mystr = ComboBoxVendors.SelectedValue.ToString
        'Dim cut_at As String = '-'
        '   Dim x As Integer = InStr(mystr, "-")



        'Dim vend As String = mystr.Substring(0, x - 1)

        'Dim string_after As String = mystr.Substring(x + cut_at.Length - 1)

        DataGridViewSCDC.Enabled = True


        Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
        checkBoxColumn.HeaderText = ""
        checkBoxColumn.Width = 30
        checkBoxColumn.Name = "checkBoxColumn"
        DataGridViewSCDC.Columns.Insert(0, checkBoxColumn)




        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim strSQL As String


        Dim stockDC As DataSet = New DataSet


        '    strSQL = "SELECT a.RegNo,  a.[Reg.Date], a.CustomerID, a.CustomerName, a.City, a.Class,a.Class1, a.Cust_Exist_New as Exist_Cust, a.CSR, a.TSSISeg, a.TSSSeg,a.MarketType, a.Enq_Ref_no, a.Enq_Ref_date, " & _
        '        "a.Enq_Source, a.Enq_Recd_date,a.Doc_upload, a.Doc_Details,a.Special_instructions, a.Enq_Int_code as Key1,a.Cust_IntCode from TSS_Enq_Pending_Project_Aproval a where " & _
        '      "a.RegNo not in (select b.Enq_Reg_No from ENQ_Project_Approval_Status b where (b.Status <> 'MoreInfoRequired') and a.RegNo = b.Enq_Reg_No) order by a.RegNo"

        strSQL = "Select VendorID, PONumber,POCreatedDate,POLineNumber, ItemNumber, LineItemOrderedQuantity  as Qty, ItemUnitCost, POLineKey, VendorKey  from FSDBBR.dbo.TSS_SubDcOpenPOS  where VendorKey =    " & mystr & ""



        'Enq_Int_code as Key1, Cust_IntCode 

        Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
        Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

        stockDAC.SelectCommand = sqlCmd
        cnSQL.Open()

        stockDAC.TableMappings.Add("Table", "Enq")
        'get data
        stockDAC.Fill(stockDC)


        DataGridViewSCDC.DataSource = stockDC.Tables(0)
        cnSQL.Close()


    End Sub

    Private Sub BTN_Click(sender As Object, e As EventArgs) Handles BTNDCSAVE.Click


        'Generate dc number

        Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String


        Dim cnSQL2 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL2 As SqlCommand
        Dim drSQL2 As SqlDataReader
        Dim strSQL2 As String



        Dim YE As Date = Today
        'Dim FINYE As String
        Dim FINYEAR As Decimal

        'Dim YEA As Integer = Year(YE)
        '        Dim YEA1 As Integer = Year(YE) + 1
        'FINYE = Trim(Str(YEA)) + Trim(Str(YEA1))
        'FINYEAR = Convert.ToDouble(FINYE)


        'strSQL1 = "select max(DCNUMBER)from FSPrograms.dbo.TSS_SubDc_Config where YEAR = " & YEA & " and ENABLED = 'Y'"
        ' strSQL1 = "select max(DCNUMBER)from FSPrograms.dbo.TSS_SubDc_Config where YEAR = " & FINYEAR & " and ENABLED = 'Y'"



        strSQL1 = "select DCNUMBER,YEAR from FSPrograms.dbo.TSS_SubDc_Config where  ENABLED = 'Y'"

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        drSQL1 = cmSQL1.ExecuteReader()


        '  cmSQL2 = New SqlCommand(strSQL1, cnSQL1)
        ' drSQL2 = cmSQL2.ExecuteReader()

        If drSQL1.Read() Then

            If IsDBNull(drSQL1.Item(0)) Then
                ' txtdc.Text = 1

                MsgBox("Data not found in TSS_SubDc_Config, please contact IT department, vbinformation")
                Exit Sub
            Else


                txtdc.Text = drSQL1.Item(0) + 1
                FINYEAR = drSQL1.Item(1)


            End If

        End If
        cnSQL1.Close()
        'end of dc number generation


        'genereate dc header key




        'strSQL2 = "select max([DC_HeaderKey])  FROM [FSPrograms].[dbo].[TSS_SubDc_Details]"
        'cnSQL2.Open()
        'cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
        'drSQL2 = cmSQL2.ExecuteReader()


        'If drSQL2.Read() Then

        '    If IsDBNull(drSQL2.Item(0)) Then
        '        txtHDKEY.Text = 1
        '    Else


        '        txtHDKEY.Text = drSQL2.Item(0) + 1


        '    End If

        'End If
        'cnSQL2.Close()
        'end of hader key




        'saving to the table TSS_SubDc_Details

        Dim strsql As String
        Dim cmSQL As SqlCommand
        ' Dim st As String
        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        cnSQL.Open()

        For i As Integer = 0 To Me.DataGridViewSCDC.RowCount - 1
            ' For i As Integer = 0 To Val(txtchkcount.Text)



            If CBool(Me.DataGridViewSCDC.Rows(i).Cells(0).Value) = True Then
                ' MessageBox.Show(Me.DataGridViewSCDC.Rows(i).Cells(1).Value.ToString())


                'genereate dc header key

                strSQL2 = "select max([DC_HeaderKey])  FROM [FSPrograms].[dbo].[TSS_SubDc_Details]"
                cnSQL2.Open()
                cmSQL2 = New SqlCommand(strSQL2, cnSQL2)
                drSQL2 = cmSQL2.ExecuteReader()


                If drSQL2.Read() Then

                    If IsDBNull(drSQL2.Item(0)) Then
                        txtHDKEY.Text = 1
                    Else


                        txtHDKEY.Text = drSQL2.Item(0) + 1


                    End If

                End If
                cnSQL2.Close()

                'end of hader key

                strsql = "insert FSPrograms.dbo.TSS_SubDc_Details values(" & txtHDKEY.Text & ", " & txtdc.Text & ",'" & YE & "','" & Me.DataGridViewSCDC.Rows(i).Cells(2).Value.ToString() & "'," & _
                    "" & Me.DataGridViewSCDC.Rows(i).Cells(4).Value.ToString() & "," & Me.DataGridViewSCDC.Rows(i).Cells(8).Value.ToString() & ",'" & Me.DataGridViewSCDC.Rows(i).Cells(1).Value.ToString() & "','" & username & "','" & txtRemarks1.Text & "', '" & txtRemark2.Text & "', " & FINYEAR & ")"


                cmSQL = New SqlCommand(strsql, cnSQL)
                'i = i + 1

                If cmSQL.ExecuteNonQuery() = 0 Then
                    MsgBox("Cannot Save dc Details. " & strsql, MsgBoxStyle.Exclamation, "Error!")

                    Exit Sub

                End If





            Else
                ' Exit Sub
                'Exit Sub
            End If

        Next
        cnSQL1.Close()


        strSQL1 = "update FSPrograms.dbo.TSS_SubDc_Config set DCNUMBER = '" & txtdc.Text & "' WHERE YEAR = " & FINYEAR & " and ENABLED = 'Y'"
        ' strSQL1 = "update FSPrograms.dbo.TSS_SubDc_Config set DCNUMBER = '" & txtdc.Text & "' WHERE ENABLED = 'Y'"

        cnSQL1.Open()
        cmSQL1 = New SqlCommand(strSQL1, cnSQL1)
        If cmSQL1.ExecuteNonQuery() = 0 Then
            MsgBox("DC Number not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
            Exit Sub

        Else

            MsgBox("DC Generated.", vbInformation)
            RBView.Checked = True
            BTNDCSAVE.Enabled = False
            Exit Sub

        End If

        'end of update

    End Sub

    Private Sub DataGridViewSCDC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewSCDC.CellContentClick
        'two dcs should not be selected 

        ' Dim i As Integer

        'If CBool(Me.DataGridViewSCDC.Rows(i).Cells(0).Value) = True Then
        'MessageBox.Show(Me.DataGridViewSCDC.Rows(i).Cells(1).Value.ToString())

        'End If


        'txtPONumber.Text = Me.DataGridViewSCDC.Rows(i).Cells(1).Value.ToString()


    End Sub

    Private Sub DataGridViewSCDCSelectAll_CurrentCellDirtyStateChanged(
        ByVal sender As Object,
        ByVal e As EventArgs) Handles DataGridViewSCDC.CurrentCellDirtyStateChanged

        RemoveHandler DataGridViewSCDC.CurrentCellDirtyStateChanged,
            AddressOf DataGridViewSCDCSelectAll_CurrentCellDirtyStateChanged

        If TypeOf DataGridViewSCDC.CurrentCell Is DataGridViewCheckBoxCell Then
            DataGridViewSCDC.EndEdit()
            Dim Checked As Boolean = CType(DataGridViewSCDC.CurrentCell.Value, Boolean)
            If Checked Then
                If Len(txtPONumber.Text) = 0 Then
                    txtPONumber.Text = Me.DataGridViewSCDC.CurrentRow.Cells(2).Value.ToString()
                End If

                If txtPONumber.Text <> Me.DataGridViewSCDC.CurrentRow.Cells(2).Value.ToString() Then

                    MessageBox.Show("Two purchase orders can't be selected for one DC")
                    dupchk = True
                    Call uncheck()
                    dupchk = False
                    '    Me.DataGridViewSCDC.CurrentRow.Cells(1).Value = False

                    '  DataGridViewSCDC.CurrentRow.Selected = True

                    '      Dim val1 As DataGridViewCheckBoxCell = DataGridViewSCDC(DataGridViewSCDC.CurrentCell.ColumnIndex, DataGridViewSCDC.CurrentCell.RowIndex)
                    '     val1.Value = False

                    'Dim allchecked As Boolean = False
                    '    If allchecked Then
                    'For Each row As DataGridViewRow In DataGridViewSCDC.Rows
                    'row.Cells(1).Value = False
                    'allchecked = False
                    'Next
                    'End If

                    ' MessageBox.Show("You have checked")
                End If

            Else
                ' MessageBox.Show("You have un-checked")
                If Val(txtchkcount.Text) = 0 Then
                    txtPONumber.Text = ""
                End If

            End If
        End If

        AddHandler DataGridViewSCDC.CurrentCellDirtyStateChanged,
            AddressOf DataGridViewSCDCSelectAll_CurrentCellDirtyStateChanged


    End Sub
    Private Sub uncheck()
        If dupchk = True Then
            Dim val1 As DataGridViewCheckBoxCell = DataGridViewSCDC(DataGridViewSCDC.CurrentCell.ColumnIndex, DataGridViewSCDC.CurrentCell.RowIndex)
            val1.Value = False


        End If
    End Sub



    '  Dim allchecked As Boolean = False
    ' Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    If allchecked Then
    '       For Each row As DataGridViewRow In DataGridView1.Rows
    '          row.Cells(x).Value = False
    '         allchecked = False
    '    Next
    'Else
    '   For Each row As DataGridViewRow In DataGridView1.Rows
    '      row.Cells(x).Value = True
    '     allchecked = True
    'Next
    'End If
    'End Sub








    Private Sub DataGridViewSCDC_MouseUp(sender As Object, e As MouseEventArgs) Handles DataGridViewSCDC.MouseUp

        'If dupchk = True Then
        '    Dim val1 As DataGridViewCheckBoxCell = DataGridViewSCDC(DataGridViewSCDC.CurrentCell.ColumnIndex, DataGridViewSCDC.CurrentCell.RowIndex)
        '    val1.Value = False
        '    dupchk = False

        'End If
    End Sub

    Private Sub RBView_CheckedChanged(sender As Object, e As EventArgs) Handles RBView.CheckedChanged
        MsgBox("http://tssblrfsh102/Reports/Pages/Report.aspx?ItemPath=%2fReports%2fPurchasing+Reports%2fSub-JobWork", vbInformation)
        Exit Sub

        'Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)

        'dlgPrintPreview = new frmPrintPreview();

        'sPrintType = queryAdapter.GetPrintType(sInvoiceNumber);

        'dlgPrintPreview.ServerURL = new Uri(ConfigurationSettings.AppSettings["ReportServer"]);

        'if (sPrintType == "EXP")
        'dlgPrintPreview.ReportPath = ConfigurationSettings.AppSettings["ExportInvoice"];
        '       Else
        'dlgPrintPreview.ReportPath = ConfigurationSettings.AppSettings["DomesticInvoice"];

        '// Create the inoice report parameter
        'ReportParameter[] reportParams = new ReportParameter[1];

        'reportParams[0] = new ReportParameter("InvoiceNumber", sInvoiceNumber);

        'dlgPrintPreview.ReportParameters = reportParams;

        'dlgPrintPreview.ShowDialog(this);

        'dlgPrintPreview.Dispose();





    End Sub

    Private Sub RBDelDC_CheckedChanged(sender As Object, e As EventArgs) Handles RBDelDC.CheckedChanged
        MsgBox("not yet ready", vbInformation)
        Exit Sub

    End Sub

    Private Sub RBCancel_CheckedChanged(sender As Object, e As EventArgs) Handles RBCancel.CheckedChanged
        MsgBox("not yet ready", vbInformation)
        Exit Sub
    End Sub

    Private Sub ComboBoxVendors_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxVendors.SelectedIndexChanged

    End Sub

    Private Sub GroupBoxPODetails_Enter(sender As Object, e As EventArgs) Handles GroupBoxPODetails.Enter

    End Sub
End Class