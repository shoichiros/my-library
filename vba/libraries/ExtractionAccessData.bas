Attribute VB_Name = "ExtractionAccessData"
Option Explicit


' -------------------------------------------------------------------------------------
' ## From Access Table or Query to Excel Sheet, import SQL executed data. ##
'
' sql --- SQL code, "FROM" is Table or Query name
' db_path --- Target Access file full path
' paste_sheet --- Objectsheet name
' is_table --- Create Data table in Excel
'
' -- For Example --
' sql = "SELECT name, age FROM [dataTable]"
' db_path = "C:\Users\{Your Username}\Desktop\AccessData.accdb"
' paste_sheet = DataSheet
' is_table = True

Sub importAccessToTableSheet(ByVal sql As String, ByVal db_path As String, _
    ByVal paste_sheet As Worksheet, ByVal is_table As Boolean)
    
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")

    With connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Open db_path
    End With

    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    Dim i As Long
    
    Set recordset = connection.Execute(sql)
    
    With paste_sheet
        .Cells.Clear
        
        For i = 0 To recordset.Fields.Count - 1
            .Cells(1, i + 1) = recordset.Fields(i).Name
        Next i
        
        .Range("A2").CopyFromRecordset recordset
        
        If is_table = True Then .ListObjects.Add
        
    End With
    
    connection.Close

End Sub
