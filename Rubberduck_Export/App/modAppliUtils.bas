Attribute VB_Name = "modAppliUtils"
Option Explicit

Sub Output_Timer_Results(subName As String, t As Double)

    Dim modeOper As Integer
    modeOper = 2 '2024-03-29 @ 11:37
    
    'modeOper = 1 - Dump to immediate Window
    If modeOper = 1 Then
        Dim l As Integer: l = Len(subName)
        Debug.Print vbNewLine & String(40 + l, "*") & vbNewLine & _
        Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & subName & " = " _
        & Format(Timer - t, "##0.0000") & " secondes" & vbNewLine & String(40 + l, "*")
    End If

    'modeOper = 2 - Dump to worksheet
    If modeOper = 2 Then
        With wshzDocLogAppli
            Dim lastUsedRow As Long
            lastUsedRow = .Range("A9999").End(xlUp).row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("B" & lastUsedRow).value = subName
            If t Then
                .Range("C" & lastUsedRow).value = Timer - t
            End If
        End With
    End If

End Sub

