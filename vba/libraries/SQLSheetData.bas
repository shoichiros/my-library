Attribute VB_Name = "SQLSheetData"
Option Explicit


Function sheetImportToArray(ByVal sql As String) As Variant

    Dim db_path As String
    db_path = ThisWorkbook.Path & "\" & ThisWorkbook.Name

    Dim ado_connection As Object
    Set ado_connection = CreateObject("ADODB.connection")

    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open db_path
    End With

    Dim ado_recordset As Object
    Set ado_recordset = CreateObject("ADODB.recordset")
    Set ado_recordset = ado_connection.Execute(sql)

    If ado_recordset.EOF = True Then
        SheetImportToArray = Empty
    Else
        SheetImportToArray = ado_recordset.GetRows
    End If

    ado_connection.Close

End Function


Sub sheetImportToSheet(ByVal sql As String, ByVal paste_start_range As Range)

    Dim db_path As String
    db_path = ThisWorkbook.Path & "\" & ThisWorkbook.Name

    Dim ado_connection As Object
    Set ado_connection = CreateObject("ADODB.connection")

    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open db_path
    End With

    Dim ado_recordset As Object
    Set ado_recordset = CreateObject("ADODB.recordset")
    Set ado_recordset = ado_connection.Execute(sql)

    paste_start_range.CopyFromRecordset ado_recordset

    ado_connection.Close

End Sub
