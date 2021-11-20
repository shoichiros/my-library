Attribute VB_Name = "ExecuteSQL"
Option Explicit


' -------------------------------------------------------------------------------------
' ## ThisWorkbook SQL execute. ##
'
' sql --- SQL code, "UPDATE" "INSERT INTO" is sheet name + $ or Objectsheet.Name + $ or Data Table name
'
' -- For Example --
' sql = "UPDATE [excelSheetName$] SET name = 'Joy' WHERE name = 'j'"
' sql = "INSERT INTO [dataTableName] VALUES ('joy', 26)"

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
