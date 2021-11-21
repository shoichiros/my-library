Attribute VB_Name = "MakePDF"
Option Explicit


Sub makePDFFile(Byval sheets_name_array As Variant, Byval output_folder_name As String)
    
    Dim main_folder_path As String
    main_folder_path = ThisWorkbook.Path & "\" & output_folder_name & "\"
    
    If Dir(main_folder_path, vbDirectory) = "" Then MkDir main_folder_path
    
    Dim base_sheet As Worksheet
            
    For Each base_sheet In Worksheets(sheets_name_array)
        base_sheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=main_folder_path & base_sheet.Name & ".pdf", _
            Quality:=xlQualityStandard
    Next base_sheet

End Sub


Sub makePDFFileAll(Byval output_file_name As String, Byval output_folder_name As String)
    
    Dim main_folder_path As String
    main_folder_path = ThisWorkbook.Path & "\" & output_folder_name & "\"
    
    If Dir(main_folder_path, vbDirectory) = "" Then MkDir main_folder_path
    
    ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=main_folder_path & output_file_name & ".pdf", _
        Quality:=xlQualityStandard

End Sub


Sub makePDFWrapSheets(Byval sheet_name_array As Variant, _
    Byval output_file_name As String, Byval output_folder_name As String)

    Dim main_folder_path As String
    main_folder_path = ThisWorkbook.Path & "\" & output_folder_name & "\"
    
    If Dir(main_folder_path, vbDirectory) = "" Then MkDir main_folder_path
    
    Worksheets(sheet_name_array).Select
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=main_folder_path & output_file_name & ".pdf", _
        Quality:=xlQualityStandard

End Sub

