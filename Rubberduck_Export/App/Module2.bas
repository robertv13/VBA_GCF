Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("btnUpdate")).Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
        .Solid
    End With
End Sub
