Attribute VB_Name = "Module2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("F7:F21").Select
    Selection.Locked = False
    Selection.FormulaHidden = False
End Sub
