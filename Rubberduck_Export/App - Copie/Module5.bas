Attribute VB_Name = "Module5"
Option Explicit

Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro8 Macro
'

'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -11511710
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Color = -11511710
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("E26").Select
End Sub
