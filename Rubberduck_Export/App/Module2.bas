Attribute VB_Name = "Module2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
    columns("A:B").Select
    Range("A2").Activate
    Selection.EntireColumn.Hidden = True
End Sub
