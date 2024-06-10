Attribute VB_Name = "Module3"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro

    Range("P1:Y1063").Select
    ActiveWorkbook.Worksheets("GL_Trans").Sort.SortFields.clear
    ActiveWorkbook.Worksheets("GL_Trans").Sort.SortFields.Add2 key:=Range( _
        "T2:T1063"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("GL_Trans").Sort.SortFields.Add2 key:=Range( _
        "Q2:Q1063"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("GL_Trans").Sort.SortFields.Add2 key:=Range( _
        "P2:P1063"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("GL_Trans").Sort
        .SetRange Range("P1:Y1063")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=524
End Sub
