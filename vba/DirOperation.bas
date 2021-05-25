Attribute VB_Name = "DirOperation"
Option Explicit

'# Required library "Microsoft Scripting Runtime"
'# For example
'
'    Dim before_folder_path As String
'    Dim after_folder_path As String
'
'    before_folder_path = "C:\Users\[username]\Desktop\test_folder"
'    after_folder_path = "C:\Users\[username]\Desktop\test\"
'
'    Call moveFolders(before_folder_path, after_folder_path)
'
Sub moveFolders(ByVal before_folder As String, _
    ByVal after_folder As String)

    Dim file_system As New FileSystemObject

    If file_system.FolderExists(before_folder) Then
        file_system.moveFolder before_folder, after_folder
    Else
        MsgBox "Does not exist folder"
        Exit Sub
    End If

End Sub
