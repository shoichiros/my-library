Attribute VB_Name = "CellsMerge"
Option Explicit


Sub cellsMerge(ByVal target_sheet As Worksheet, ByVal base_column As Long, _
    ByVal start_row As Long, ByVal target_column As Long, ByVal is_sum As Boolean)

    Dim merge_target As Range
    Dim last_low As Long
    Dim i As Long
    
    Set merge_target = target_sheet.Cells(start_row, target_column)
    last_low = target_sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = start_row To last_low
        
        If target_sheet.Cells(i, base_column) = target_sheet.Cells(i, base_column).Offset(1, 0) Then
            Set merge_target = Union(merge_target, target_sheet.Cells(i, target_column).Offset(1, 0))
        Else
            Application.DisplayAlerts = False
            
            If is_sum = True Then merge_target = WorksheetFunction.Sum(merge_target)
            
            merge_target.Merge
            Set merge_target = merge_target.Offset(1, 0)
        End If
        
    Next i

End Sub

