Attribute VB_Name = "EncodeImportCSVFunction"
Option Explicit

Function encodedImportCSV(encode As String)

    Dim csv_file As String
    csv_file = Application.GetOpenFilename("CSV Files(*.csv),*.csv", , "CSVファイルを選択")

    If csv_file = "" Then
        MsgBox "キャンセルしました。"
    End If

    Dim encode_stream As New ADODB.Stream

    ' Encoding CSV file into text
    With encode_stream
        .Open
        .LoadFromFile csv_file
        .Type = adTypeText
        .Charset = encode
        encodedImportCSV = .ReadText
        .Close
    End With

End Function
