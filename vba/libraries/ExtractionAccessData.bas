Attribute VB_Name = "ExtractionAccessData"
Option Explicit


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
