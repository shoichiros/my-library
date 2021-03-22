Attribute VB_Name = "CSVOperation"
Option Explicit

' Required Microsoft ActiveX Data Objects 6.1 Library
' HDR=YES is field name header
' HDR=NO is field name F1, F2, F3....etc
'
' ## For example ##
'     Dim csv_file_path as String
'     Dim sql As String
'     Dim file_name As String
'
'     csv_file_path = Application.GetOpenFilename("CSV(*.csv), *.csv", , "csv")
'     file_name = Dir(csv_file_path)
'     sql = "SELECT *" _
'        & " FROM " & file_name

Function CSVImportToArray(csv_file_path As String, sql As String)

    If Dir(csv_file_path) = "" Then Exit Function
    
    Dim file_name As String
    Dim folder_path As String
    
    file_name = Dir(csv_file_path)
    folder_path = Replace(csv_file_path, file_name, "")
    
    Dim ado_connection As New ADODB.connection
        
    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "TEXT;HDR=YES;FMT=Delimited"
        .Open folder_path
    End With
    
    Dim ado_recodeset As New ADODB.Recordset
    Dim lists As Variant
    
    Set ado_recodeset = ado_connection.Execute(sql)
    lists = ado_recodeset.GetRows
    
    If IsArray(lists) = True Then
        CSVImportToArray = WorksheetFunction.Transpose(lists)
    Else
        CSVImportToArray = Empty
    End If
    
    ado_connection.Close
        
End Function

' Required Microsoft ActiveX Data Objects 6.1 Library

' ## For example ##
' base_data is multiple array
' full path is "C:\Users\username\Documents\test_file.txt"
' Olso "C:\Users\username\Documents\test_file.csv"

Sub outputToCSV(ByVal base_data As Variant, ByVal full_path As String)
    
    Dim ado_stream As New ADODB.Stream
    Dim i As Long
    Dim j As Long
    Dim data As String
    
    With ado_stream
        .Open
        
        For i = LBound(base_data) To UBound(base_data)
            data = ""
            
            For j = LBound(base_data, 2) To UBound(base_data, 2)
                data = data & base_data(i, j) & ","
            Next j
        
            .WriteText Left(data, Len(data) - 1), 1
        Next i
        
        .SaveToFile full_path, 2
        .Close
    End With
    
End Sub

