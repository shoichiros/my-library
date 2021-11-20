Attribute VB_Name = "WrapPrintOutSheet"
Option Explicit


Sub wrapPrintOutSheets()
    
    Dim i As Long
    Dim sheet_counts As Long
    Dim sheet_container() As Variant
    
    sheet_counts = Sheets.Count
    ReDim sheet_container(1 To sheet_counts)
    
    For i = LBound(sheet_container) To UBound(sheet_container)
        sheet_container(i) = Sheets(i).Name
    Next i
    
    Worksheets(sheet_container).PrintPreview
    
End Sub
