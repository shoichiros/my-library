Attribute VB_Name = "MultipleMkdir"
Option Explicit


Sub multipleLayersMkDir(ByVal output_folder_path As String)

    Dim file_system As New FileSystemObject
    Dim folder_lists As Variant
    Dim i As Long
    Dim folder_path As String

    folder_lists = Split(output_folder_path, "\")

    For i = LBound(folder_lists) To UBound(folder_lists)
        folder_path = folder_path + folder_lists(i) & "\"

        If folder_path = "\" Then
            folder_path = folder_path
        ElseIf folder_path = "\\" Then
            i = i + 1
            folder_path = folder_path + folder_lists(i) & "\"
        ElseIf file_system.FolderExists(folder_path) = False Then
            MkDir folder_path
        End If

    Next i

End Sub

