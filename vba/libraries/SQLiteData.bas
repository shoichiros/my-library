Attribute VB_Name = "SQLiteData"
Option Explicit


Function SQLiteImportToArray(ByVal db_path As String, ByVal sql As String) As Variant
    
    Const CONNECT_DRIVER As String = "DRIVER=SQLite3 ODBC Driver;"
    Dim data_source As String: data_source = CONNECT_DRIVER & "Database=" & db_path & ";"
    
    Dim ado_connection As Object
    Set ado_connection = CreateObject("ADODB.connection")
    ado_connection.Open data_source

    Dim ado_recordset As Object: Set ado_recordset = CreateObject("ADODB.recordset")
    Set ado_recordset = ado_connection.Execute(sql)

    If ado_recordset.EOF = True Then
        SQLiteImportToArray = Empty
    Else
        SQLiteImportToArray = ado_recordset.GetRows
    End If

End Function


Sub SQLiteImportToSheet(ByVal db_path As String, _
    ByVal sql As String, ByVal paste_start_range As Range)

    Const CONNECT_DRIVER As String = "DRIVER=SQLite3 ODBC Driver;"
    Dim data_source As String: data_source = CONNECT_DRIVER & "Database=" & db_path & ";"
    
    Dim ado_connection As Object
    Set ado_connection = CreateObject("ADODB.connection")
    ado_connection.Open data_source

    Dim ado_recordset As Object: Set ado_recordset = CreateObject("ADODB.recordset")
    Set ado_recordset = ado_connection.Execute(sql)

    paste_start_range.CopyFromRecordset ado_recordset

End Sub


Sub SQLiteExecution(ByVal db_path As String, ByVal sql As String)

    Const CONNECT_DRIVER As String = "DRIVER=SQLite3 ODBC Driver;"
    Dim data_source As String: data_source = CONNECT_DRIVER & "Database=" & db_path & ";"
    
    Dim ado_connection As Object
    Set ado_connection = CreateObject("ADODB.connection")
    ado_connection.Open data_source

    Dim ado_recordset As Object: Set ado_recordset = CreateObject("ADODB.recordset")
    Set ado_recordset = ado_connection.Execute(sql)

End Sub

