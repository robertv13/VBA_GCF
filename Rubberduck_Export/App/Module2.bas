Attribute VB_Name = "Module2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
    Range("M5:S5").Select
    With Selection.Font
        .Name = "Aptos Narrow"
        .Size = 11
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    With Selection.Font
        .Name = "Aptos Narrow"
        .Size = 11
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Bold = True
    Selection.Font.Bold = False
    Selection.Font.Italic = True
    Selection.Font.Italic = False
End Sub
