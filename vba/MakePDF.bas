Attribute VB_Name = "makePDF"
Option Explicit

' Args is sheets_name_array = Array("sheet_name1", "sheet_name2")

Sub makePDFFile(sheets_name_array As Variant, folder_name As String)
    
    Dim main_folder_path As String
    main_folder_path = ThisWorkbook.Path & "\" & folder_name
    
    If Dir(main_folder_path, vbDirectory) = "" Then MkDir main_folder_path
    
    Dim base_sheet As Worksheet
            
    For Each base_sheet In Worksheets(sheets_name_array)
        base_sheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=main_folder_path & "\" & base_sheet.Name & ".pdf", _
            Quality:=xlQualityStandard
    Next base_sheet

End Sub

