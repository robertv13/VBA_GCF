Attribute VB_Name = "Module3"
Option Explicit

Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ActiveWindow.SmallScroll Down:=-18
    Range("L11:N45,O48:O50,M48:M50").Select
    Range("M48").Activate
    ActiveWindow.SmallScroll Down:=-9
    Range("L11:N45,O48:O50,M48:M50,O9,O3,E3:F3").Select
    Range("E3").Activate
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
