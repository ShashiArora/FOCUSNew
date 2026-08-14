'Imports System
'Imports System.Data
Imports System.Data.SqlClient
'Imports System.ComponentModel
Imports System.Configuration

Module Module1

    Public username As String
    Public usertype As String

    Public qtycheck As Boolean
    Public mode As String
    Public curdate As Date
    Public screentype As String

    Public receiptno As Integer
    Public reqno As Integer
    Public transtype As String
    Public issno As Integer
    Public toolretno As Integer
    Public receiptmfg As Integer

    Private ConnectionString As String
    Public stockDS As DataSet = New DataSet
    Public check As Integer
    Public transmode As String
    Public rfqcomp As String
    Public datatype As String
    Public parthistory As String


    Public ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"

    Public Sub nogenerate()

        Dim cnSQL5 As SqlConnection = New SqlConnection(ConnectionStringNew)
        Dim cmSQL1 As SqlCommand
        Dim drSQL1 As SqlDataReader
        Dim strSQL1 As String
        'Material Receipt ******
        If transtype = "Receipts" And transmode = "Add" Then
            strSQL1 = "Select Last_Used_No from TSS_WH_Trans_NumberControl where TransactionType = 'Receipts'"
            cnSQL5.Open()

            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            drSQL1 = cmSQL1.ExecuteReader()

            If drSQL1.Read() Then
                receiptno = drSQL1.Item(0) + 1
            End If
            cnSQL5.Close()
        ElseIf transtype = "Receipts" And transmode = "Update" Then


            strSQL1 = "Update TSS_WH_Trans_NumberControl set Last_Used_No = " & receiptno & " where TransactionType = 'Receipts'"
            cnSQL5.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            If cmSQL1.ExecuteNonQuery() = 0 Then
                MsgBox("Receipt Number Not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If
            cnSQL5.Close()
        End If
        'Material Request ****************
        If transtype = "MatReq" And transmode = "Add" Then
            strSQL1 = "Select Last_Used_No from TSS_WH_Trans_NumberControl where TransactionType = 'MatReq'"
            cnSQL5.Open()

            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            drSQL1 = cmSQL1.ExecuteReader()

            If drSQL1.Read() Then
                reqno = drSQL1.Item(0) + 1
            End If
            cnSQL5.Close()
        ElseIf transtype = "MatReq" And transmode = "Update" Then


            strSQL1 = "Update TSS_WH_Trans_NumberControl set Last_Used_No = " & reqno & " where TransactionType = 'MatReq'"
            cnSQL5.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            If cmSQL1.ExecuteNonQuery() = 0 Then
                MsgBox("Request Number Not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If
            cnSQL5.Close()
        End If
   
        'Material Issue ***********************************

        'Material Request *********************************
        If transtype = "MatIss" And transmode = "Add" Then
            strSQL1 = "Select Last_Used_No from TSS_WH_Trans_NumberControl where TransactionType = 'MatIss'"
            cnSQL5.Open()

            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            drSQL1 = cmSQL1.ExecuteReader()

            If drSQL1.Read() Then
                issno = drSQL1.Item(0) + 1
            End If
            cnSQL5.Close()
        ElseIf transtype = "MatIss" And transmode = "Update" Then


            strSQL1 = "Update TSS_WH_Trans_NumberControl set Last_Used_No = " & issno & " where TransactionType = 'MatIss'"
            cnSQL5.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            If cmSQL1.ExecuteNonQuery() = 0 Then
                MsgBox("Issue Number Not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If
            cnSQL5.Close()
        End If
        '******************************************

        'Tool Return  *********************************
        If transtype = "ToolRet" And transmode = "Add" Then
            strSQL1 = "Select Last_Used_No from TSS_WH_Trans_NumberControl where TransactionType = 'ToolRet'"
            cnSQL5.Open()

            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            drSQL1 = cmSQL1.ExecuteReader()

            If drSQL1.Read() Then
                toolretno = drSQL1.Item(0) + 1
            End If
            cnSQL5.Close()
        ElseIf transtype = "ToolRet" And transmode = "Update" Then


            strSQL1 = "Update TSS_WH_Trans_NumberControl set Last_Used_No = " & toolretno & " where TransactionType = 'ToolRet'"
            cnSQL5.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            If cmSQL1.ExecuteNonQuery() = 0 Then
                MsgBox("Tool Return Number is not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If
            cnSQL5.Close()
        End If
        '******************************************

        'Material Receipt at Mfg building  *********************************

        If transtype = "ReceiptsMfg" And transmode = "Add" Then
            strSQL1 = "Select Last_Used_No from TSS_WH_Trans_NumberControl where TransactionType = 'ReceiptsMfg'"
            cnSQL5.Open()

            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            drSQL1 = cmSQL1.ExecuteReader()

            If drSQL1.Read() Then
                receiptmfg = drSQL1.Item(0) + 1
            End If
            cnSQL5.Close()
        ElseIf transtype = "ReceiptsMfg" And transmode = "Update" Then


            strSQL1 = "Update TSS_WH_Trans_NumberControl set Last_Used_No = " & receiptmfg & " where TransactionType = 'ReceiptsMfg'"
            cnSQL5.Open()
            cmSQL1 = New SqlCommand(strSQL1, cnSQL5)
            If cmSQL1.ExecuteNonQuery() = 0 Then
                MsgBox("Receipt number at mfg location is not updated." & strSQL1, MsgBoxStyle.Exclamation, "Error!")
                Exit Sub
            End If
            cnSQL5.Close()
        End If
        '***********************************************


    End Sub


End Module
