Attribute VB_Name = "DirOperation"
Option Explicit

'# For example
'    Dim move_folder_name As String
'    Dim before_folder As String
'    Dim after_folder As String
'
'    move_folder_name = "test_folder" ' Target folder name
'    before_folder = "C:\Users\[username]\Desktop\test\" ' Until before target folder path
'    after_folder = "C:\Users\[username]\Desktop\" ' Folder path you want to move
'
'    Call moveFolder(move_folder_name, before_folder, after_folder)

Sub moveFolder(ByVal move_folder_name As String, _
    ByVal before_folder As String, ByVal after_folder As String)
    
    Dim move_before_folder As String
    Dim move_after_folder As String
    
    move_before_folder = before_folder & move_folder_name
    move_after_folder = after_folder & move_folder_name
    
    If Dir(move_after_folder, vbDirectory) <> "" Then
        MsgBox "The destination folder has the same folder name." & vbLf _
            & "Ends the process without moving the folder."
        End
    End If
    
    Name move_before_folder As move_after_folder

End Sub

