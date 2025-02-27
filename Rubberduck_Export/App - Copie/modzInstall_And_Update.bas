Attribute VB_Name = "modzInstall_And_Update"
Option Explicit

Sub Clearcontents()

    Dim ws As Worksheet
    Dim rng As Range
    Dim lastUsedRow As Long, firstDataRow As Long
    
    Set ws = wshBD_Clients
    firstDataRow = 2
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow >= firstDataRow Then
        Set rng = ws.Range("A" & firstDataRow & ":J" & lastUsedRow)
        rng.Select
        MsgBox rng.Address
        rng.Clearcontents
    End If
    
    Set ws = wshzDocLogAppli
    firstDataRow = 2
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow >= firstDataRow Then
        Set rng = ws.Range("A" & firstDataRow & ":C" & lastUsedRow)
        rng.Select
        MsgBox rng.Address
        rng.Clearcontents
    End If
    
End Sub

Sub Generic_Clearcontents(Worksheet As String, Optional headingRows As Integer = 2)

    Dim ws As Worksheet
    Dim rngToClear As Range
    
    MsgBox Worksheet
    Set ws = ThisWorkbook.Worksheets(Worksheet)

End Sub
