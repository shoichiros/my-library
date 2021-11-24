Attribute VB_Name = "ImportText"
Option Explicit


Sub importCSVFullData(ByVal import_sheet As Worksheet)
    
    Dim file_full_path As String
    file_full_path = Application.GetOpenFilename("CSV(*.csv), *.csv", , "csv")

    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    Workbooks.Open file_full_path
    
    With import_sheet
        Dim data_lists As Variant
        Dim last_row As Long: last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        Dim last_column As Long: last_column = .Cells(1, .Columns.Count).End(xlToLeft).Column

        data_lists = .Range(.Cells(1, 1), .Cells(last_row, last_column))
        ActiveWorkbook.Close
        
        import_sheet.Range("A1").Resize(last_row, last_column) = data_lists
    End With
    
    Application.ScreenUpdating = True: Application.DisplayAlerts = True
    
End Sub
