Attribute VB_Name = "CSVOperation"
Option Explicit

' Required Microsoft ActiveX Data Objects 6.1 Library
' HDR=YES is field name header
' HDR=NO is field name F1, F2, F3....etc

Sub CSVImport(csv_file_path As String)

    If csv_file_path = "False" Then
        MsgBox "CSV file does not exist."
        Exit Sub
    Else
        Dim ado_connection As New ADODB.connection
        
        With ado_connection
            .Provider = "Microsoft.ACE.OLEDB.16.0"
            .Properties("Extended Properties") = "TEXT;HDR=YES;FMT=Delimited"
            .Open ThisWorkbook.Path & "\"
        End With
        
        Dim file_name As Variant
        Dim sql As String
        Dim ado_recodeset As New ADODB.Recordset
        
        file_name = Split(csv_file_path, "\")
        sql = "SELECT *" _
            & " FROM " & file_name(UBound(file_name))
        
        Set ado_recodeset = ado_connection.Execute(sql)
        
        Cells.Clear
        Cells(1, 1).CopyFromRecordset ado_recodeset
        
        ado_connection.Close
        
        Set ado_recodeset = Nothing
        Set ado_connection = Nothing
    End If
        
End Sub
