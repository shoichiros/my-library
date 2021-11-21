Attribute VB_Name = "LastDayFunc"
Option Explicit


Function monthLastDay(Byval month_number As Long, Byval is_character As Boolean)

    Dim last_day As Date
    Dim character_last_day As String

    last_day = DateSerial(Year(Date), Month(Date) + 1 + month_number, 1) - 1

    If is_character = True Then
        character_last_day = Replace(last_day, "/", "-")
        thisMonthLastDay = character_last_day
    Else
        thisMonthLastDay = last_day
    End If

End Function

