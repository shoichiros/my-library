Attribute VB_Name = "NewSheetAdd"
Option Explicit

Sub newSheetsAdd()

    Dim sheet_name As String
    sheet_name = InputBox("�V�[�g�̖��O�́H", Title:="�V�[�g���̓���")
    
    If sheet_name = "" Then
        MsgBox "�V�[�g�쐬���L�����Z�����܂����B"
        Exit Sub
    Else
        Dim make_sheets_number As Long
        make_sheets_number = InputBox("�������̃V�[�g�����܂����H", Title:="�V�[�g�̖���")
        
        If make_sheets_number = 0 Then
            MsgBox "�V�[�g�쐬���L�����Z�����܂����B"
            Exit Sub
        End If
    End If
    
    Dim i As Long
    
    For i = 1 To make_sheets_number
        Sheets.Add.Name = sheet_name & i
    Next i

End Sub

