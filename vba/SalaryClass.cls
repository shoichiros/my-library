VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SalaryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const MAX_REGULAR_TIME As Date = "8:00"


Public Sub timeCalc(work_start_time As Date, work_end_time As Date, _
work_break_time As Date, writing_start_cells As Range)

    Dim time_calc As Date
    time_calc = work_end_time - work_break_time - work_start_time
    writing_start_cells = time_calc

End Sub


Public Sub regularWorkTime(total_work_time As Date, writing_start_cells As Range)
    
    Dim regular_work_time As Date
    
    If total_work_time < MAX_REGULAR_TIME Then
        regular_work_time = total_work_time
    Else
        regular_work_time = MAX_REGULAR_TIME
    End If
    
    writing_start_cells = regular_work_time

End Sub


Public Sub overtimeWork(total_work_time As Date, writing_start_cells As Range)
    
    Dim overtime As Date
    
    If total_work_time > MAX_REGULAR_TIME Then
        overtime = total_work_time - MAX_REGULAR_TIME
    Else
        overtime = 0
    End If
    
    writing_start_cells = overtime
    
End Sub


Public Sub salaryMoneyCalc(writing_start_cells As Range, regular_work_time As Date, _
regular_money_hour As Integer, over_work_time As Date, over_time_ratio As Double)

    Dim regular_money_calc As Long
    Dim over_time_money_calc As Long
    Dim total_money As Long
    Dim time_ajustment As Long
    time_ajustment = 24
    
    regular_money_calc = regular_work_time * time_ajustment * regular_money_hour
    over_time_money_calc = regular_money_hour * time_ajustment * over_time_ratio * over_work_time
    total_money = regular_money_calc + over_time_money_calc
    
    writing_start_cells = total_money
    
End Sub
