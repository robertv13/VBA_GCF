Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "1000"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = ">=31/07/2024"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "<01/08/2024"
    Range("O3").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("A1:J529").AdvancedFilter action:=xlFilterCopy, criteriaRange:=Range( _
        "L2:N3"), CopyToRange:=Range("P1:Y1"), Unique:=False
End Sub
