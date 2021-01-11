Attribute VB_Name = "SheetLastRowFunction"
Option Explicit

Function sheetLastRow(ws As Worksheet, Optional CheckCol As Long = 1) As Long
  
    sheetLastRow = ws.Cells(1, CheckCol).End(xlDown).Row

End Function
