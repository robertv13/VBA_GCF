Attribute VB_Name = "modzzzInstall_And_Update"
Option Explicit

Sub zzz_Clearcontents()

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
        rng.ClearContents
    End If
    
    Set ws = wshzDocLogAppli
    firstDataRow = 2
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow >= firstDataRow Then
        Set rng = ws.Range("A" & firstDataRow & ":C" & lastUsedRow)
        rng.Select
        MsgBox rng.Address
        rng.ClearContents
    End If
    
End Sub

