Attribute VB_Name = "makePDF"
Option Explicit

' is_make_name_folder = True is file_name folder include

Sub makePDFFile(base_sheet As Worksheet, file_name As String, is_make_name_folder As Boolean)
    
    Dim main_folder_path As String
    
    main_folder_path = ThisWorkbook.Path & "\"
    
    If is_make_name_folder = True Then
        MkDir main_folder_path & file_name
        base_sheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=main_folder_path & file_name & "\" & file_name & ".pdf", _
            Quality:=xlQualityStandard
    Else
        base_sheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=main_folder_path & file_name & ".pdf", _
            Quality:=xlQualityStandard
    End If

End Sub

