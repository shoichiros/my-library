Attribute VB_Name = "EncodeCSV"
Option Explicit


Function encodedImportCSV(Byval encode As String)

    Dim csv_file As String
    csv_file = Application.GetOpenFilename("CSV Files(*.csv),*.csv", , "csv")

    If csv_file = "" Then: MsgBox "Process canceled.": End

    Dim encode_stream As Object
    Set encode_stream = CreateObject("ADODB.Stream")

    With encode_stream
        .Open
        .LoadFromFile csv_file
        .Type = adTypeText
        .Charset = encode
        encodedImportCSV = .ReadText
        .Close
    End With

End Function
