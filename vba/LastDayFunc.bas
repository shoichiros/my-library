Attribute VB_Name = "LastDayFunc"
Option Explicit

' is_character = True Where to use, the ID, File name and Folder name
' is_chatacter = True is "2021-03-31" as String
' is_character = False is "2021/03/31" as Date

Function thisMonthLastDay(is_character As Boolean)
    
    Dim last_day As Date
    Dim character_last_day As String
    
    last_day = DateSerial(Year(Date), Month(Date) + 1, 1) - 1
    
    If is_character = True Then
        character_last_day = Replace(last_day, "/", "-")
        thisMonthLastDay = character_last_day
    Else
        thisMonthLastDay = last_day
    End If

End Function


Function nextMonthLastDay(is_character As Boolean)
    
    Dim last_day As Date
    Dim character_last_day As String
    
    last_day = DateSerial(Year(Date), Month(Date) + 2, 1) - 1
    
    If is_character = True Then
        character_last_day = Replace(last_day, "/", "-")
        nextMonthLastDay = character_last_day
    Else
        nextMonthLastDay = last_day
    End If

End Function
