Attribute VB_Name = "UtfEncodingCSV"
Option Explicit
' Import library, Microsoft ActiveX Data Objects 6.1 library
Sub utfEncodingCSV()

    Dim csv_file As String

    csv_file = Application.GetOpenFilename("CSV Files(*.csv),*.csv", , "CSVファイルを選択")
    
    If csv_file = "" Then Exit Sub

    Dim encoded_csv_file As String
    Dim utf_stream As New ADODB.Stream

    encoded_csv_file = "encoded_file.csv"

    ' Encoding CSV as utf-8
    With utf_stream
        .Open
        .LoadFromFile csv_file
        .Type = adTypeText
        .Charset = "utf-8"
        .SaveToFile encoded_csv_file, 2
        .Close
    End With

End Sub
