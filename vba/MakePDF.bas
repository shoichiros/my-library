Attribute VB_Name = "makePDF"
Option Explicit


Sub makePDFFile(base_sheet_array As Variant, folder_name As String)
    
    Dim main_folder_path As String
    main_folder_path = ThisWorkbook.Path & "\" & folder_name
    
    If Dir(main_folder_path) = "" Then MkDir main_folder_path
    
    Dim base_sheet As Worksheet
            
    For Each base_sheet In Worksheets(base_sheet_array)
        base_sheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=main_folder_path & "\" & base_sheet.Name & ".pdf", _
            Quality:=xlQualityStandard
    Next base_sheet

End Sub

