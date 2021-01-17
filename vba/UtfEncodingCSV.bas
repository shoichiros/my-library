Attribute VB_Name = "UtfEncodingCSV"
Option Explicit
' Import library, Microsoft ActiveX Data Objects 6.1 library
Sub utfEncodingCSV()

    Dim csv_file As String
    Dim encoded_csv_file As String
    Dim utf_stream As New ADODB.Stream
    
    csv_file = Application.GetOpenFilename("CSV Files(*.csv),*.csv", , "CSVファイルを選択")
    encoded_csv_file = "encoded_file.csv"
    
    ' Encoding CSV as utf-8
    With utf_stream
        .Open
        .LoadFromFile csv_file
        .Type = adTypeText
        .Charset = "utf-8"
        .SaveToFile encoded_csv_file
        .Close
    End With
    
End Sub
