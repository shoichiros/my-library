Attribute VB_Name = "SheetLastRow"
Option Explicit


Function sheetLastRow(byval target_sheet As Worksheet, Optional check_column As Long = 1) As Long
  
    sheetLastRow = target_sheet.Cells(1, check_column).End(xlDown).Row

End Function
