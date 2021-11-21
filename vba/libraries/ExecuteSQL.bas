Attribute VB_Name = "ExecuteSQL"
Option Explicit


Sub thisworkbookExecuteSQL(ByVal sql As String)
    
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")

    With connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open ThisWorkbook.Path & "\" & ThisWorkbook.Name
    End With

    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    Set recordset = connection.Execute(sql)
    
    connection.Close

End Sub
