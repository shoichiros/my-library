Attribute VB_Name = "MultipleMkdir"
Option Explicit


Sub multipleLayersMkDir(ByVal output_folder_path As String)

    Dim folder_lists As Variant
    Dim i As Long
    Dim folder_path As String
    
    folder_lists = Split(output_folder_path, "\")
    
    For i = LBound(folder_lists) To UBound(folder_lists)
        folder_path = folder_path + folder_lists(i) & "\"
    
        If Dir(folder_path, vbDirectory) = "" Then MkDir folder_path
        
    Next i
    
End Sub
