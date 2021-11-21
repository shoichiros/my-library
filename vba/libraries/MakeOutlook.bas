Attribute VB_Name = "MakeOutlook"
Option Explicit


Sub makeOutlookMail(Byval address As String, Byval subject As String, _
    Byval body As String, Byval is_attach As Boolean, Optional Byval attach_file_path_array As Variant)

    Dim outlook_app As New Outlook.Application
    Dim outlook_mail As Outlook.MailItem
    
    Set outlook_mail = outlook_app.CreateItem(olMailItem)
    
    With outlook_mail
        .BodyFormat = olFormatPlain
        .To = address
        .subject = subject
        .body = body
        
        If is_attach = True Then
            
            If Dir(attach_file_path_array(0)) <> "" Then
                Dim attach_file As Variant
            
                For Each attach_file In attach_file_path_array
                    .Attachments.Add attach_file, olByValue
                Next attach_file
                
            Else
                MsgBox "Does not exist Attachments file"
                Exit Sub
            End If
            
        End If
        ' Saved Outlook draft
        .Save
    End With
    
    Set outlook_mail = Nothing
    Set outlook_app = Nothing
    
End Sub

