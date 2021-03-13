Attribute VB_Name = "makePDF"
Option Explicit


' ## makePDFFile ##
' Some sheet choice, make PDF files

' Args is sheets_name_array = Array("sheet_name1", "sheet_name2")
' Also sheets_name_array = Array(1, 2)

Sub makePDFFile(sheets_name_array As Variant, output_folder_name As String)
    
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


' ## makePDFFileAll ##
' ThisWorkbook all sheets, make only one PDF file

Sub makePDFFileAll(output_file_name As String, output_folder_name As String)
    
    Dim main_folder_path As String
    main_folder_path = ThisWorkbook.Path & "\" & output_folder_name & "\"
    
    If Dir(main_folder_path, vbDirectory) = "" Then MkDir main_folder_path
    
    ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=main_folder_path & output_file_name & ".pdf", _
        Quality:=xlQualityStandard

End Sub


' ## makePDFWrapSheets ##
' From some sheets, make only one PDF file

' Args is sheets_name_array = Array("sheet_name1", "sheet_name2")
' Also sheets_name_array = Array(1, 2)

Sub makePDFWrapSheets(sheet_name_array As Variant, _
    output_file_name As String, output_folder_name As String)

    Dim main_folder_path As String
    main_folder_path = ThisWorkbook.Path & "\" & output_folder_name & "\"
    
    If Dir(main_folder_path, vbDirectory) = "" Then MkDir main_folder_path
    
    Worksheets(sheet_name_array).Select
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=main_folder_path & output_file_name & ".pdf", _
        Quality:=xlQualityStandard

End Sub

