Attribute VB_Name = "Module1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'
    Range("B17:E31").Select
    ActiveWorkbook.Worksheets("�tat des R�sultats").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�tat des R�sultats").Sort.SortFields.Add2 key:= _
        Range("C17:C31"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("�tat des R�sultats").Sort
        .SetRange Range("B17:E31")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
