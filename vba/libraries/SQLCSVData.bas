Attribute VB_Name = "SQLCSVData"
Option Explicit


Function CSVImportToArray(ByVal csv_full_path As String, ByVal sql As String) As Variant

    If Dir(csv_full_path) = "" Then Exit Function

    Dim file_name As String
    Dim folder_path As String

    file_name = Dir(csv_full_path)
    folder_path = Replace(csv_full_path, file_name, "")

    Dim ado_connection As Object
    Set ado_connection = CreateObject("ADODB.connection")

    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "TEXT;HDR=YES;FMT=Delimited"
        .Open folder_path
    End With

    Dim ado_recordset As Object
    Set ado_recordset = CreateObject("ADODB.recordset")
    Set ado_recordset = ado_connection.Execute(sql)

    If ado_recordset.EOF = True Then
        CSVImportToArray = Empty
    Else
        CSVImportToArray = ado_recordset.GetRows
    End If

    ado_connection.Close

End Function


Sub CSVImportToSheet(ByVal csv_full_path As String, _
    ByVal sql As String, ByVal paste_start_range As Range)

    If Dir(csv_full_path) = "" Then Exit Sub

    Dim file_name As String
    Dim folder_path As String

    file_name = Dir(csv_full_path)
    folder_path = Replace(csv_full_path, file_name, "")

    Dim ado_connection As Object
    Set ado_connection = CreateObject("ADODB.connection")

    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "TEXT;HDR=YES;FMT=Delimited"
        .Open folder_path
    End With

    Dim ado_recordset As Object
    Set ado_recordset = CreateObject("ADODB.recordset")
    Set ado_recordset = ado_connection.Execute(sql)

    paste_start_range.CopyFromRecordset ado_recordset

    ado_connection.Close

End Sub
