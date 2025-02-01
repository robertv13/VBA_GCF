Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Unprotect
    Range("G2").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
