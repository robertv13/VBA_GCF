Attribute VB_Name = "Module4"
Option Explicit

Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Range("B8").Select
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Range("B16").Select
    Selection.Font.underline = xlUnderlineStyleNone
    Selection.Font.underline = xlUnderlineStyleSingle
    Range("B17").Select
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
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
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
