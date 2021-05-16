Attribute VB_Name = "CSVOperation"
Option Explicit

' Required Microsoft ActiveX Data Objects 6.1 Library
' HDR=YES is field name header
' HDR=NO is field name F1, F2, F3....etc
'
' # For example
'     Dim csv_full_path As String
'     Dim sql As String
'     Dim file_name As String
'
'     csv_full_path = Application.GetOpenFilename("CSV(*.csv), *.csv", , "csv")
'     file_name = Dir(csv_full_path)
'     sql = "SELECT *" _
'        & " FROM [" & file_name & "]"

Function CSVImportToArray(ByVal csv_full_path As String, ByVal sql As String) As Variant

    If Dir(csv_full_path) = "" Then Exit Function
    
    Dim file_name As String
    Dim folder_path As String
    
    file_name = Dir(csv_full_path)
    folder_path = Replace(csv_full_path, file_name, "")
    
    Dim ado_connection As New ADODB.connection
        
    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "TEXT;HDR=YES;FMT=Delimited"
        .Open folder_path
    End With
    
    Dim ado_recordset As New ADODB.Recordset
    Set ado_recordset = ado_connection.Execute(sql)
    
    If ado_recordset.EOF = True Then
        CSVImportToArray = Empty
    Else
        CSVImportToArray = ado_recordset.GetRows
    End If
    
    ado_connection.Close
        
End Function

' Required Microsoft ActiveX Data Objects 6.1 Library
' "ADO recordset GetRows" exclusive use CSV export Sub

Sub getRowsExportToCSV(ByVal base_data As Variant)

    If IsArray(base_data) = False Then Exit Sub
    
    Dim ado_stream As New ADODB.Stream
    Dim i As Long
    Dim j As Long
    Dim data_row As String
    
    With ado_stream
        .Open
        
        For i = LBound(base_data, 2) To UBound(base_data, 2)
            data_row = ""
            
            For j = LBound(base_data) To UBound(base_data)
                data_row = data_row & base_data(j, i) & ","
            Next j
            
            .WriteText Left(data_row, Len(data_row) - 1), adWriteLine
        Next i
        
        .SaveToFile "test.csv", adSaveCreateOverWrite
        .Close
    End With
    
End Sub

' Required Microsoft ActiveX Data Objects 6.1 Library
' Nomal array or multiple array Sub

Sub arrayExportToCSV(ByVal data_lists As Variant)

    If IsArray(data_lists) = False Then Exit Sub

    Dim ado_stream As New ADODB.Stream
    Dim i As Long
    Dim j As Long
    Dim data_row As String
    
    With ado_stream
        .Open
        
        For i = LBound(data_lists) To UBound(data_lists)
            data_row = ""
            
            For j = LBound(data_lists, 2) To UBound(data_lists, 2)
                data_row = data_row & data_lists(i, j) & ","
            Next j
            
            .WriteText Left(data_row, Len(data_row) - 1), adWriteLine
        Next i
        
        .SaveToFile "test.csv", adSaveCreateOverWrite
        .Close
    End With

End Sub

' Required Microsoft ActiveX Data Objects 6.1 Library
' HDR=YES is field name header
' HDR=NO is field name F1, F2, F3....etc
'
' # For example
'     Dim csv_full_path As String
'     Dim sql As String
'     Dim file_name As String
'     Dim paste_start_range As Range
'
'     csv_full_path = Application.GetOpenFilename("CSV(*.csv), *.csv", , "csv")
'     file_name = Dir(csv_full_path)
'     sql = "SELECT *" _
'        & " FROM [" & file_name & "]"
'     Set paste_start_range = Sheet1.Range("A2")

Sub CSVImportToSheet(ByVal csv_full_path As String, _
    ByVal sql As String, ByVal paste_start_range As Range)

    If Dir(csv_full_path) = "" Then Exit Sub
    
    Dim file_name As String
    Dim folder_path As String
    
    file_name = Dir(csv_full_path)
    folder_path = Replace(csv_full_path, file_name, "")
    
    Dim ado_connection As New ADODB.connection
        
    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "TEXT;HDR=YES;FMT=Delimited"
        .Open folder_path
    End With
    
    Dim ado_recordset As New ADODB.Recordset
    Set ado_recordset = ado_connection.Execute(sql)
    
    paste_start_range.CopyFromRecordset ado_recordset
    
    ado_connection.Close
        
End Sub

' Required Microsoft ActiveX Data Objects 6.1 Library
' # For example
'     Dim sheet_name As String
'     Dim sql As String
'
'     sheet_name = Sheet1.Name
'     sql = "SELECT *" _
'         & " FROM [" & sheet_name & "$]"

Function SheetImportToArray(ByVal sql As String) As Variant
    
    Dim db_path As String
    db_path = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    
    Dim ado_connection As New ADODB.connection
        
    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open db_path
    End With
    
    Dim ado_recordset As New ADODB.Recordset
    Set ado_recordset = ado_connection.Execute(sql)
    
    If ado_recordset.EOF = True Then
        SheetImportToArray = Empty
    Else
        SheetImportToArray = ado_recordset.GetRows
    End If
    
    ado_connection.Close
        
End Function

' Required Microsoft ActiveX Data Objects 6.1 Library
' # For example
'     Dim sheet_name As String
'     Dim sql As String
'     Dim paste_start_range As Range
'
'     sheet_name = Sheet1.Name
'     sql = "SELECT *" _
'         & " FROM [" & sheet_name & "$]"
'     Set paste_start_range = Sheet1.Range("A2")

Function SheetImportToSheet(ByVal sql As String, ByVal paste_start_range As Range) As Variant
    
    Dim db_path As String
    db_path = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    
    Dim ado_connection As New ADODB.connection
        
    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open db_path
    End With
    
    Dim ado_recordset As New ADODB.Recordset
    Set ado_recordset = ado_connection.Execute(sql)
    
    paste_start_range.CopyFromRecordset ado_recordset
    
    ado_connection.Close
        
End Function
