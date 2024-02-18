Attribute VB_Name = "modHiddenNamedRange"
Option Explicit

Sub unhideAllNames()
'Unhide all names in the currently open Excel file
    For Each tempName In ActiveWorkbook.Names
        tempName.Visible = True
    Next
End Sub

Sub removeAllHiddenNames()
'Remove all hidden names in current workbook, no matter if hidden or not
    For Each tempName In ActiveWorkbook.Names
        If tempName.Visible = False Then
            tempName.Delete
        End If
    Next
End Sub
