Attribute VB_Name = "DirOperation"
Option Explicit


Sub moveFolders(ByVal before_folder As String, _
    ByVal after_folder As String)

    Dim file_system As Object
    Set file_system = CreateObject("Scripting.FileSystemObject")

    If file_system.FolderExists(before_folder) Then
        file_system.moveFolder before_folder, after_folder
    Else
        MsgBox "Does not exist folder"
        Exit Sub
    End If

End Sub
