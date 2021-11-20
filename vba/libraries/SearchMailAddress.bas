Attribute VB_Name = "SearchMailAddress"
Option Explicit

' --------------------------------------------------
' # excel data table example(sheet name: address_list)
'
' targetName | sendMethod | mailAddress
' --------------------------------------------
' shouTanaka | to                | example@example.com
' shouSuzuki  | cc                | examples_cc1@example.com
' kanaAbe      | cc                | examples_cc2@example.com
'
' Output:
' to: example@example.com
' cc: examples_cc1@example.com;examples_cc2@example.com
'
' --------------------------------------------------

' --------------------------------------------------
' # sheetDataToArray #
' ## Args for example:
' sql = "SELECT *" _
'   & " FROM [sheet_name$] "
' --------------------------------------------------

Private Function sheetDataToArray(ByVal sql As String) As Variant
    
    Dim ado_connection As New ADODB.Connection
    Dim sheet_db As String
    
    sheet_db = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    
    With ado_connection
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open sheet_db
    End With
    
    Dim ado_recordset As ADODB.Recordset
    Set ado_recordset = ado_connection.Execute(sql)
    
    sheetDataToArray = ado_recordset.GetRows
       
    ado_connection.Close

End Function

' --------------------------------------------------
' # getMailAddressLists #
' First data extraction
' --------------------------------------------------

Function getMailAddressLists(ByVal base_sheet_name As String, _
    ByVal target_name As String) As Variant

    Dim sql As String
    Dim mail_address_lists As Variant
    
    sql = "SELECT *" _
        & " FROM [" & base_sheet_name & "$]" _
        & " WHERE targetName = '" & target_name & "'"
    
    getMailAddressLists = sheetDataToArray(sql)
    
End Function

' --------------------------------------------------
' # sortMailAddress #
' ## Args for example:
' mail_address_lists = getMailAddressLists
' --------------------------------------------------

Function sortMailAddress(ByVal mail_address_lists As Variant, _
    ByVal is_cc_address As Boolean) As String
    
    Dim i As Long
    Dim mail_address As String
    
    For i = LBound(mail_address_lists, 2) To UBound(mail_address_lists, 2)
        
        If is_cc_address = True And mail_address_lists(1, i) = "cc" Then
            mail_address = mail_address & mail_address_lists(2, i) & ";"
        Else
            mail_address = mail_address_lists(2, i)
        End If
        
    Next i
    
    If is_cc_address = True Then
        mail_address = Left(mail_address, Len(mail_address) - 1)
    End If
    
    sortMailAddress = mail_address

End Function


' test case
Sub testAddressExtraction()

    Dim mail_address_lists As Variant
    mail_address_lists = getMailAddressLists("address_list", "target_name")
    
    Dim mail_address As String
    mail_address = sortMailAddress(mail_address_lists, True)
    Stop
    
End Sub
