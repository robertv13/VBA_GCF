Attribute VB_Name = "Module2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("tblTEC_TDB_Data[#All]").AdvancedFilter action:=xlFilterCopy, _
        criteriaRange:=Range("S1:U2"), CopyToRange:=Range("W1:AD1"), Unique:= _
        False
    ActiveWindow.SmallScroll Down:=-11
End Sub
