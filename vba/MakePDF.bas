Attribute VB_Name = "makePDF"
Option Explicit

' is_make_name_folder = True is file_name folder include
' Add "\" to the end of the folder_path

Sub makePDFFile(base_sheet As Worksheet, file_name As String, folder_path As String, is_make_name_folder As Boolean)
    
    If is_make_name_folder = True Then
        MkDir main_folder_path & file_name
        base_sheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=folder_path & file_name & "\" & file_name & ".pdf", _
            Quality:=xlQualityStandard
    Else
        base_sheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=folder_path & file_name & ".pdf", _
            Quality:=xlQualityStandard
    End If

End Sub

