Attribute VB_Name = "Protection"
Option Explicit


Sub sheetProtection(ByVal is_protect As Boolean, _
    ByVal password As String, ByVal target_sheets As Variant)

    Dim sheet As Variant
    
    For Each sheet In target_sheets
    
        If is_protect = True Then
            sheet.Protect password
        Else
            sheet.Unprotect password
        End If
    
    Next sheet

End Sub
