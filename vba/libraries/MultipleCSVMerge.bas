Attribute VB_Name = "MultipleCSVMerge"
Option Explicit


Sub csvFilesMerge(ByVal target_folder As String, ByVal output_folder As String)

    Dim file_system As Object
    Set file_system = CreateObject("Scripting.FileSystemObject")

    Dim text_stream As TextStream
    Dim text_stream_output As TextStream
    Dim results As String
    Dim i As Long
    Dim file_lists As Variant
    
    file_lists = getfileLists(target_folder)
    
    For i = LBound(file_lists) To UBound(file_lists)
        Set text_stream = file_system.OpenTextFile(target_folder & file_lists(i))
        
        If i = 0 Then
            Set text_stream_output = file_system.OpenTextFile(output_folder & "output.csv", ForWriting, True)
            
            results = text_stream.ReadAll
            text_stream_output.Write results
        Else
            Set text_stream_output = file_system.OpenTextFile(output_folder & "output.csv", ForAppending, True)
            
            text_stream.SkipLine
            results = text_stream.ReadAll
            text_stream_output.Write results
        End If
    
        text_stream.Close
        text_stream_output.Close
    Next i
    
End Sub


Private Function getfileLists(ByVal folder_path As String) As Variant

    Dim target_file As String
    Dim files As String
    
    target_file = Dir(folder_path & "*.csv")
    
    Do While target_file <> ""
        files = files + target_file & ","
        target_file = Dir()
    Loop
    
    files = Left(files, Len(files) - 1)
    getfileLists = Split(files, ",")
    
End Function
