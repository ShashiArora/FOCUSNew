Imports System.Drawing.Printing
Imports System.Drawing.Imaging
Imports System.IO
Imports Microsoft.Reporting.WinForms
Imports System
Imports System.Collections.Generic
Imports System.Text

Imports System.Configuration

Imports System.Data.SqlClient
'Imports CrystalDecisions.CrystalReports.Engine
Imports System.Collections
Imports System.Data
Imports System.Math







Public Class ShipLabelDomestic
    Dim pages As New List(Of Metafile)
    Dim pageIndex As Integer = 0
    Dim doc As New Printing.PrintDocument()
    Dim ReportViewer1 As New ReportViewer
    'Protected Const CONNECTION_STRING As String = "Server=10.56.40.5;Database=FSPrograms;User ID=sa;Password=Trelleborg123"

    Public Report As String
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"
    Public Shared ConnectionString As String = ConnectionStringNew




    Private Sub PrintPageHandler(ByVal sender As Object, _
        ByVal e As PrintPageEventArgs)
        Dim page As Metafile = pages(pageIndex)
        pageIndex += 1
        Dim pWidth As Integer = 827
        Dim pHeight As Integer = 1100
        e.Graphics.DrawImage(page, 0, 0, pWidth, pHeight)
        e.HasMorePages = pageIndex < pages.Count
    End Sub

    '   Private Sub Button1_Click(ByVal sender As System.Object, _
    'ByVal e As System.EventArgs) Handles Button1.Click
    'Dim doc As New Printing.PrintDocument()
    'doc = New Printing.PrintDocument()
    'AddHandler doc.PrintPage, AddressOf PrintPageHandler
    'Dim dialog As New PrintDialog()
    'dialog.Document = doc
    'Dim print As DialogResult
    'print = dialog.ShowDialog()

    'doc.PrinterSettings = dialog.PrinterSettings

    'Dim deviceInfo As String = _
    '"<DeviceInfo>" & _
    '"<OutputFormat>emf</OutputFormat>" & _
    '"  <PageWidth>8.5in</PageWidth>" & _
    '"  <PageHeight>11in</PageHeight>" & _
    '"  <MarginTop>0.25in</MarginTop>" & _
    '"  <MarginLeft>0.25in</MarginLeft>" & _
    '"  <MarginRight>0.25in</MarginRight>" & _
    '"  <MarginBottom>0.25in</MarginBottom>" & _
    '"</DeviceInfo>"

    'Dim warnings() As Microsoft.Reporting.WinForms.Warning
    'Dim streamids() As String
    'Dim mimeType, encoding, filenameExtension, path As String
    'mimeType = "" : encoding = "" : filenameExtension = ""

    ''Input parameter report
    ''  Dim DateFrom As Date = CDate("4/15/2015")
    '' Dim DateTo As Date = CDate("4/15/2015")

    ''Dim parmDateFrom As New ReportParameter("DateFrom", DateFrom)
    ''Dim parmDateTo As New ReportParameter("DateTo", DateTo)
    ''Dim parmSO1(1) As ReportParameter
    ''parmSO1(0) = parmDateFrom
    ''parmSO1(1) = parmDateTo

    'Dim data() As Byte
    ''ReportViewer1.ServerReport.SetParameters(parmSO1)
    'data = ReportViewer1.ServerReport.Render("Image", _
    '       deviceInfo, mimeType, encoding, filenameExtension, _
    '       streamids, warnings)
    'pages.Add(New Metafile(New MemoryStream(data)))

    'For Each pageName As String In streamids
    '    data = ReportViewer1.ServerReport.RenderStream("Image", _
    '           pageName, deviceInfo, mimeType, encoding)
    '    pages.Add(New Metafile(New MemoryStream(data)))
    'Next
    'doc.Print()
    'Me.ReportViewer1.RefreshReport()
    '  End Sub



    Private Sub Form5_Load(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles MyBase.Load

        '        Button1.Text = "Print"
        'With ReportViewer1
        '.Visible = False
        '.ProcessingMode = ProcessingMode.Remote
        '.ServerReport.ReportPath = "/Reports/IT-PendingForTesting/Modified Reports/LabelPrint"
        '.ServerReport.ReportServerUrl = New  _
        'Uri("http://tssblrfsh101/reportserver")
        'End With
        BtnPrint.Enabled = False

        Me.Controls.Add(ReportViewer1)
        ReportViewer1.Hide()

    End Sub

    Private Sub BtnShipOK_Click(sender As Object, e As EventArgs) Handles BtnShipOK.Click
        '  Dim report As String
        BtnPrint.Enabled = True
        'Report format selection
        With ReportViewer1
            .Visible = False
            .ProcessingMode = ProcessingMode.Remote


            If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
                report = "A"
                .ServerReport.ReportPath = "/Reports/SCM/LabelPrint"

            ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
                report = "B"
                .ServerReport.ReportPath = "/Reports/SCM/LabelPrintInv"

            ElseIf ComboBox1.Text = Trim("C)Box Label Pre-Invoice") Then
                report = "C"
                .ServerReport.ReportPath = "/Reports/SCM/BoxLabel"

            ElseIf ComboBox1.Text = Trim("D)Box Label Post-Invoice") Then
                Report = "D"
                .ServerReport.ReportPath = "/Reports/SCM/BoxLabel"
            End If

            '   .ServerReport.ReportPath = "/Reports/IT-PendingForTesting/Modified Reports/LabelPrint"
            .ServerReport.ReportServerUrl = New  _
                     Uri("http://tssblrfsh101/reportserver")
        End With
        'end of report format selection


        '   btnshipprint.Enabled = True


        '  Dim cnSQL As SqlConnection
        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String




        'If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
        '    report = "A"
        '    '  shipenable()

        'ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
        '    report = "B"
        '    'shipenable()

        'ElseIf ComboBox1.Text = Trim("C)Box Label Pre-Invoice") Then
        '    report = "C"
        '    'shipdisable()
        'ElseIf ComboBox1.Text = Trim("D)Box Label Post Invoice") Then
        '    report = "D"
        'End If


        If txtInvoiceNo.Text.Length = 0 And (report = "A" Or report = "C") Then
            MsgBox("Please, enter Shipment no.!", MsgBoxStyle.Critical, "Error!")
            txtInvoiceNo.Focus()
            Exit Sub
        End If

        If txtInvoiceNo.Text.Length = 0 And (report = "B" Or report = "D") Then
            MsgBox("Please, enter Invoice No.!", MsgBoxStyle.Critical, "Error!")
            txtInvoiceNo.Focus()
            Exit Sub
        End If



        '  If report = "A" Then

        DataGridViewShipLabel.Columns.Clear()

        'adding check box
        'Dim checkCol As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
        'checkCol.HeaderText = "Del"
        'DataGridViewProjectMasterEdit.Columns.Add(checkCol)
        'end of adding check box

        DataGridViewShipLabel.ReadOnly = False

        Dim AlarmColumn1 As New DataGridViewCheckBoxColumn(False)

        AlarmColumn1.Name = "Sel"
        AlarmColumn1.HeaderText = "Selection"
        AlarmColumn1.ReadOnly = False


        DataGridViewShipLabel.Columns.Add(AlarmColumn1)
        DataGridViewShipLabel.ReadOnly = False


        '  ButtonProjDetailsDelete.Enabled = True

        'GroupBoxEdit.Visible = True
        DataGridViewShipLabel.Visible = True
        DataGridViewShipLabel.Enabled = True

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)


        Dim stockDC As DataSet = New DataSet

        If report = "A" Then
            strSQL = "Select 0 as LabelQty, CONumber,COLineNumber,ItemNumber,LotNumber, ShippedQuantity, CustomerID, ShipmentDate, HistoryShipmentKey " & _
                    "from FSDBBR.dbo.TSS_Item_label_History_Shipment " & _
                    "where ShippedQuantity > 0 and ShipmentNumber  = " & txtInvoiceNo.Text & " and " & _
                    "COLineNumber like '" & TxtLineNo.Text & "' order by COLineNumber "

        ElseIf report = "B" Then
            strSQL = "Select 0 as LabelQty, a.CONumber,a.COLineNumber,a.ItemNumber,a.LotNumber, a.ShippedQuantity, a.CustomerID, a.ShipmentDate, a.HistoryShipmentKey " & _
                     "FROM   FSDBBR.dbo.TSS_Item_label_History_Shipment a LEFT OUTER JOIN FSDBBR.dbo.TSS_ItemLabel_InvoiceHistory_FSPrograms b ON a.ShipmentNumV = b.ShipmentNo AND a.ItemNumber = b.ItemNumber " & _
                     "where  b.InvoiceNo = '" & txtInvoiceNo.Text & "' And COLineNumber like '" & TxtLineNo.Text & "' order by COLineNumber "


            '   strSQL = "Select 0 as LabelQty, a.CONumber,a.COLineNumber,a.ItemNumber,a.LotNumber, a.ShippedQuantity, a.CustomerID, a.ShipmentDate, a.HistoryShipmentKey " & _
            '"FROM   dbo.TSS_Item_label_History_Shipment a LEFT OUTER JOIN dbo.TSS_ItemLabel_InvoiceHistory_FSPrograms b ON a.ShipmentNumV = b.ShipmentNo " & _
            '"where (Len(b.InvoiceNo) > 0)  And b.InvoiceNo = " & txtInvoiceNo.Text & " And " & _
            '"COLineNumber like '" & TxtLineNo.Text & "' order by COLineNumber "


        ElseIf report = "C" Then

            strSQL = "SELECT  a.ShipmentNumber,0 as InvoiceNumber, a.CustomerID, b.CustomerName FROM [FSDBBR].[dbo].[_NoLock_FS_ShipmentHeader] a " & _
             "LEFT OUTER JOIN FSDBBR.dbo.[_NoLock_FS_Customer] b ON a.CustomerID = b.CustomerID  " & _
             "where a.ShipmentNumber = '" & txtInvoiceNo.Text & "'"

            'SELECT  a.ShipmentNumber,a.CustomerID, b.CustomerName FROM [FSDBBR].[dbo].[_NoLock_FS_ShipmentHeader] a
            '            LEFT OUTER JOIN FSDBBR.dbo.[_NoLock_FS_Customer] b ON a.CustomerID = b.CustomerID 
            '          where a.ShipmentNumber = '16902'


        ElseIf Report = "D" Then
            strSQL = " SELECT a.InvoiceNumber, right(a.CONumber,6) AS ShipmentNumber ,a.CustomerID, b.CustomerName FROM [FSDBBR].[dbo].[_NoLock_FS_ARInvoiceHeader] a " & _
             "LEFT OUTER JOIN FSDBBR.dbo.[_NoLock_FS_Customer] b ON a.CustomerID = b.CustomerID  where a.InvoiceNumber = '" & txtInvoiceNo.Text & "'"


        Else
            MsgBox("Pl select the Line", vbInformation)
            Exit Sub
        End If


        If Report = "C" Or Report = "D" Then
            Dim cnSQL1 As SqlConnection = New SqlConnection(ConnectionStringNew)
            Dim cmSQL1 As SqlCommand
            Dim drSQL1 As SqlDataReader
            ' Dim strSQL1 As String


            cnSQL1.Open()

            cmSQL1 = New SqlCommand(strSQL, cnSQL1)
            drSQL1 = cmSQL1.ExecuteReader()

            If drSQL1.Read() Then
                txtCustomer.Text = drSQL1.Item(3)
            End If
            cnSQL1.Close()
        Else






            Dim sqlCmd As SqlCommand = New SqlCommand(strSQL, cnSQL)
            Dim stockDAC As SqlDataAdapter = New SqlDataAdapter

            stockDAC.SelectCommand = sqlCmd
            cnSQL.Open()



            stockDAC.TableMappings.Add("Table", "Enq")
            'get data
            stockDAC.Fill(stockDC)
            '  txtCustomer.Text = stockDAC.SelectCommand(3)


            DataGridViewShipLabel.DataSource = stockDC.Tables(0)



            'DataGridViewProjectMasterEdit.Columns.Add(checkCol)

            DataGridViewShipLabel.Columns("CONumber").ReadOnly = True

            DataGridViewShipLabel.Columns("COLineNumber").ReadOnly = True

            DataGridViewShipLabel.Columns("ItemNumber").ReadOnly = True

            DataGridViewShipLabel.Columns("LotNumber").ReadOnly = True

            DataGridViewShipLabel.Columns("ShippedQuantity").ReadOnly = True
            DataGridViewShipLabel.Columns("CustomerID").ReadOnly = True
            DataGridViewShipLabel.Columns("ShipmentDate").ReadOnly = True
            DataGridViewShipLabel.Columns("HistoryShipmentKey").ReadOnly = True

            cnSQL.Close()







            Exit Sub

            MsgBox(Err.Description)

            '  cnSQL.Close()
            '   End Try

            'Catch
            '    MsgBox("Wrong data entered! Check the Invoice number! ", MsgBoxStyle.Exclamation, "Error!")
            '    ClearAll()
            'End Try

        End If
        'BtnPrint.Enabled = True
    End Sub

    Private Sub DataGridViewShipLabel_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewShipLabel.CellContentClick

        Dim ColumnName1 As String = DataGridViewShipLabel.Columns(e.ColumnIndex).Name

        If ColumnName1 = "Sel" Then
            Dim CellCheckBox1 As DataGridViewCheckBoxCell = _
                CType(DataGridViewShipLabel.Rows(e.RowIndex).Cells(ColumnName1), DataGridViewCheckBoxCell)

            Dim CellCheckBoxState1 As String = CellCheckBox1.EditingCellFormattedValue.ToString


            DataGridViewShipLabel.Rows(CellCheckBox1.RowIndex).Cells(0).Value = True
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click

        If Report = "A" Or Report = "B" Then


            Dim Lblqty As Integer
            Dim NoOfLbl As Integer
            Dim C As Integer
            Dim D As Integer

            For Each RW As DataGridViewRow In Me.DataGridViewShipLabel.Rows
                If RW.Cells(0).Value = True Then

                    C = Val(RW.Cells(6).Value) Mod Val(RW.Cells(1).Value)

                    D = Val(RW.Cells(6).Value) - Val(C)
                    NoOfLbl = D / Val(RW.Cells(1).Value)

                    If C > 0 Then

                        NoOfLbl = NoOfLbl + 1

                    End If

                    '     End If

                    Dim CTR As Integer
                    '    Dim LABELCOUNT As Integer
                    '    LABELCOUNT = Val(txtshiplabel.Text)
                    '    Dim lastlabelqty As Integer

                    'lastlabelqty = Val(TXTLOTQTYPORV.Text) - ((Val(txtNoofLabels.Text) - 1) * Val(TextlblqtyPORV.Text))
                    'lastlabelqty = Val(txtshipqty.Text) - ((Val(txtshiplabel.Text) - 1) * Val(txtshipqpl.Text))


                    CTR = 1

                    Do While CTR <= NoOfLbl
                        If CTR = NoOfLbl And C > 0 Then

                            Lblqty = C
                        Else
                            Lblqty = RW.Cells(1).Value
                        End If

                        'printing coding and parameter passing
                        'history ship key
                        'lblqty

                        Dim doc As New Printing.PrintDocument()
                        doc = New Printing.PrintDocument()
                        AddHandler doc.PrintPage, AddressOf PrintPageHandler
                        Dim dialog As New PrintDialog()
                        dialog.Document = doc
                        Dim print As DialogResult
                        print = dialog.ShowDialog()

                        doc.PrinterSettings = dialog.PrinterSettings

                        Dim deviceInfo As String = _
                        "<DeviceInfo>" & _
                        "<OutputFormat>emf</OutputFormat>" & _
                        "  <PageWidth>8.5in</PageWidth>" & _
                        "  <PageHeight>11in</PageHeight>" & _
                        "  <MarginTop>0.25in</MarginTop>" & _
                        "  <MarginLeft>0.25in</MarginLeft>" & _
                        "  <MarginRight>0.25in</MarginRight>" & _
                        "  <MarginBottom>0.25in</MarginBottom>" & _
                        "</DeviceInfo>"

                        Dim warnings() As Microsoft.Reporting.WinForms.Warning
                        Dim streamids() As String
                        Dim mimeType, encoding, filenameExtension, path As String
                        mimeType = "" : encoding = "" : filenameExtension = ""

                        'Input parameter report
                        Dim HistoryShipmentKey As Integer  'CDate("4/15/2015")
                        Dim Quantity As Integer  'CDate("4/15/2015")
                        HistoryShipmentKey = RW.Cells(9).Value
                        Quantity = Lblqty
                        '                Dim parmphk As New ReportParameter("phk", phk)
                        '               Dim parmlqty As New ReportParameter("lqty", lqty)

                        Dim phk As New ReportParameter("HistoryShipmentKey", HistoryShipmentKey)
                        Dim lqty As New ReportParameter("Quantity", Quantity)


                        Dim parmSO1(1) As ReportParameter
                        parmSO1(0) = phk
                        parmSO1(1) = lqty

                        Dim data() As Byte
                        ReportViewer1.ServerReport.SetParameters(parmSO1)
                        data = ReportViewer1.ServerReport.Render("Image", _
                               deviceInfo, mimeType, encoding, filenameExtension, _
                               streamids, warnings)
                        pages.Add(New Metafile(New MemoryStream(data)))

                        For Each pageName As String In streamids
                            data = ReportViewer1.ServerReport.RenderStream("Image", _
                                   pageName, deviceInfo, mimeType, encoding)
                            pages.Add(New Metafile(New MemoryStream(data)))
                        Next
                        doc.Print()
                        Me.ReportViewer1.RefreshReport()


                        'end of printing coding

                        CTR = CTR + 1
                    Loop




                    '     TextBox1.Text = RW.Cells(2).Value.ToString
                    '    TextBox2.Text = RW.Cells(1).Value.ToString
                End If

            Next
        ElseIf Report = "C" Or Report = "D" Then
            Dim NoOfLbl As Integer
            Dim CTR As Integer

            NoOfLbl = Val(TxtLineNo.Text)


            CTR = 1

            Do While CTR <= NoOfLbl


                Dim doc As New Printing.PrintDocument()
                doc = New Printing.PrintDocument()
                AddHandler doc.PrintPage, AddressOf PrintPageHandler
                Dim dialog As New PrintDialog()
                dialog.Document = doc
                Dim print As DialogResult
                print = dialog.ShowDialog()

                doc.PrinterSettings = dialog.PrinterSettings

                Dim deviceInfo As String = _
                "<DeviceInfo>" & _
                "<OutputFormat>emf</OutputFormat>" & _
                "  <PageWidth>8.5in</PageWidth>" & _
                "  <PageHeight>11in</PageHeight>" & _
                "  <MarginTop>0.25in</MarginTop>" & _
                "  <MarginLeft>0.25in</MarginLeft>" & _
                "  <MarginRight>0.25in</MarginRight>" & _
                "  <MarginBottom>0.25in</MarginBottom>" & _
                "</DeviceInfo>"

                Dim warnings() As Microsoft.Reporting.WinForms.Warning
                Dim streamids() As String
                Dim mimeType, encoding, filenameExtension, path As String
                mimeType = "" : encoding = "" : filenameExtension = ""

                'Input parameter report
                Dim InvoiceNumber As String
                ' Dim ModeOfDespatch As String
                Dim RepType As String
                Dim modd As String

                InvoiceNumber = txtInvoiceNo.Text
                'ModeOfDespatch = txtModeofDespatch.Text
                RepType = Report
                modd = txtModeofDespatch.Text

                Dim inv As New ReportParameter("InvoiceNumber", InvoiceNumber)
                '  Dim mode As New ReportParameter("ModeOfDespatch", ModeOfDespatch)
                Dim rept As New ReportParameter("RepType", RepType)
                Dim moddesp As New ReportParameter("modd", modd)


                Dim parmSO1(2) As ReportParameter
                parmSO1(0) = inv
                '    parmSO1(1) = mode
                parmSO1(1) = rept
                parmSO1(2) = moddesp

                Dim data() As Byte
                ReportViewer1.ServerReport.SetParameters(parmSO1)
                data = ReportViewer1.ServerReport.Render("Image", _
                       deviceInfo, mimeType, encoding, filenameExtension, _
                       streamids, warnings)
                pages.Add(New Metafile(New MemoryStream(data)))

                For Each pageName As String In streamids
                    data = ReportViewer1.ServerReport.RenderStream("Image", _
                           pageName, deviceInfo, mimeType, encoding)
                    pages.Add(New Metafile(New MemoryStream(data)))
                Next
                doc.Print()
                Me.ReportViewer1.RefreshReport()


                'end of printing coding

                CTR = CTR + 1
            Loop



        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_Leave(sender As Object, e As EventArgs) Handles ComboBox1.Leave
        txtCustomer.ReadOnly = True

        If ComboBox1.Text = Trim("A)Item Labels Pre-Invoice") Then
            LblInvoice.Text = "Shipment No."
            LblLineNo.Text = "Line No."
            TxtLineNo.Text = "%"
            lblModeofdespatch.Visible = False
            txtModeofDespatch.Visible = False


            lblCustomer.Visible = False
            txtCustomer.Visible = False
        ElseIf ComboBox1.Text = Trim("B)Item Labels Post-Invoice") Then
            LblInvoice.Text = "Invoice No."
            LblLineNo.Text = "Line No."
            TxtLineNo.Text = "%"
            lblModeofdespatch.Visible = False
            txtModeofDespatch.Visible = False
           
            lblCustomer.Visible = False
            txtCustomer.Visible = False
        ElseIf ComboBox1.Text = Trim("C)Box Label Pre-Invoice") Then
            LblInvoice.Text = "Shipment No."
            lblModeofdespatch.Visible = True
            txtModeofDespatch.Visible = True
            LblLineNo.Text = "No.of Labels"
            TxtLineNo.Text = "1"
            lblCustomer.Visible = True
            txtCustomer.Visible = True


        ElseIf ComboBox1.Text = Trim("D)Box Label Post-Invoice") Then
            ' LblInvoice.Text = ""
            LblInvoice.Text = "Invoice No."
            lblModeofdespatch.Visible = True
            txtModeofDespatch.Visible = True
            LblLineNo.Text = "No.of Labels"
            TxtLineNo.Text = "1"
            lblCustomer.Visible = True
            txtCustomer.Visible = True
        End If
        CLEAR()


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub btnclear_Click(sender As Object, e As EventArgs) Handles btnclear.Click
        ' DataGridViewShipLabel.DataSource = Nothing
        'DataGridViewShipLabel.Rows.Clear()
        'DataGridViewShipLabel.Columns.Clear()
        'txtCustomer.Text = ""
        'txtModeofDespatch.Text = ""
        'txtInvoiceNo.Text = ""
        'TxtLineNo.Text = ""
        CLEAR()
   
    End Sub
    Private Sub CLEAR()
        DataGridViewShipLabel.DataSource = Nothing
        DataGridViewShipLabel.Rows.Clear()
        DataGridViewShipLabel.Columns.Clear()
        txtCustomer.Text = ""
        txtModeofDespatch.Text = ""
        txtInvoiceNo.Text = ""
        '  TxtLineNo.Text = "1"
    End Sub

    Private Sub BtnPrint_Leave(sender As Object, e As EventArgs) Handles BtnPrint.Leave

    End Sub
End Class