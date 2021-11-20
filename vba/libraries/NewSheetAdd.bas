Attribute VB_Name = "NewSheetAdd"
Option Explicit

Sub newSheetsAdd()

    Dim sheet_name As String
    sheet_name = InputBox("シートの名前は？", Title:="シート名の入力")
    
    If sheet_name = "" Then
        MsgBox "シート作成をキャンセルしました。"
        Exit Sub
    Else
        Dim make_sheets_number As Long
        make_sheets_number = InputBox("何枚分のシートを作りますか？", Title:="シートの枚数")
        
        If make_sheets_number = 0 Then
            MsgBox "シート作成をキャンセルしました。"
            Exit Sub
        End If
    End If
    
    Dim i As Long
    
    For i = 1 To make_sheets_number
        Sheets.Add.Name = sheet_name & i
    Next i

End Sub

