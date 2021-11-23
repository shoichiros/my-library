Attribute VB_Name = "TrimData"
Option Explicit


Function trimBetweenStringData(ByVal base_string As String, _
    ByVal trim_start_char As String, ByVal trim_end_char As String) As String
    
    Dim trimmed_string As String
    
    trimmed_string = Right(base_string, Len(base_string) - InStr(base_string, trim_start_char))
    trimmed_string = Left(trimmed_string, InStr(trimmed_string, trim_end_char) - 1)
    
    trimBetweenStringData = trimmed_string

End Function
