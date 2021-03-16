Attribute VB_Name = "CSVOperation"
Option Explicit

' Required Microsoft ActiveX Data Objects 6.1 Library
' HDR=YES is field name header
' HDR=NO is field name F1, F2, F3....etc
'
' ## For example ##
'     Dim sql As String
'     Dim file_name As String
'
'     file_name = Dir(csv_file_path)
'     sql = "SELECT *" _
'        & " FROM " & file_name

Function CSVImport(csv_file_path As String, sql As String)
    
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
    Set ado_recodeset = ado_connection.Execute(sql)
        
    Dim lists As Variant
    lists = ado_recodeset.GetRows
    
    If IsArray(lists) = True Then
        CSVImport = WorksheetFunction.Transpose(lists)
    Else
        CSVImport = Empty
    End If
    
    ado_connection.Close
        
    Set ado_recodeset = Nothing
    Set ado_connection = Nothing
        
End Function


Sub outputToCSV(sheet_name As String, folder_name As String)
    
    Dim folder_path As String
    folder_path = ThisWorkbook.Path & "\" & folder_name & "\"
    
    Application.ScreenUpdating = False
    
    If Dir(folder_path, vbDirectory) = "" Then MkDir folder_path
    
    Worksheets(sheet_name).Copy
    ActiveSheet.SaveAs Filename:=folder_path & sheet_name & ".csv", FileFormat:=xlCSV
    ActiveSheet.Close
    
    Application.ScreenUpdating = True
    
End Sub

