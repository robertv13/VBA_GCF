Attribute VB_Name = "Module2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("M5:S21").Select
    Sheets("Menu").Select
    Sheets("Doc_Formules").Visible = True
    ActiveWindow.SmallScroll Down:=140
    Sheets("Menu").Select
    Sheets("Doc_ConditionalFormatting").Visible = True
    Range("E2").Select
    Application.CutCopyMode = False
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Sheets("GL_BV").Select
    Selection.FormatConditions.add Type:=xlExpression, Formula1:= _
        "=ET($M5<>"""";MOD(LIGNE();3)=1)"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
