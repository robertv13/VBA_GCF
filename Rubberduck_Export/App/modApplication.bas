Attribute VB_Name = "modApplication"
Option Explicit


Global Const gAppVersion As String = "v2.8.4" '2024-03-05 @ 11:17

Public isTab_Order_Activated As Boolean

Sub BackToMainMenu()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.name <> "Menu" Then ws.Visible = xlSheetHidden
    Next ws
    wshMenu.Activate
    wshMenu.Range("A1").Select

End Sub


Sub RedefineDynamicRange() '2024-02-13 @ 13:30
    
    'Delete existing dynamic named range (assuming it exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptableDescription").delete
    On Error GoTo 0
    
    'Define a new dynamic named range
    Dim newRangeFormula As String
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,1)"
    
    'Create a new dynamic named range
    ThisWorkbook.Names.add name:="dnrPlanComptableDescription", RefersTo:=newRangeFormula
    
End Sub

Sub Hide_All_Worksheet_Except_Menu() '2024-02-20 @ 07:28
    
    Dim timerStart As Double: timerStart = Timer
    
    Dim wsh As Worksheet
    For Each wsh In ThisWorkbook.Worksheets
        If wsh.codeName <> "wshMenu" And _
            InStr(wsh.codeName, "wshzDoc") = 0 Then
                wsh.Visible = xlSheetHidden
        End If
    Next wsh
    
    Call Output_Timer_Results("Hide_All_Worksheet_Except_Menu()", timerStart)
    
End Sub

Sub LoopThroughRows()
    Dim i As Long, lastRow As Long
    Dim pctdone As Single
    lastRow = Range("A" & Rows.count).End(xlUp).row
    lastRow = 30

    '(Step 1) Display your Progress Bar
    ufProgress.LabelProgress.width = 0
    ufProgress.show
    For i = 1 To lastRow
        '(Step 2) Periodically update progress bar
        pctdone = i / lastRow
        With ufProgress
            .Caption = "Étape " & i & " of " & lastRow
            .LabelProgress.width = pctdone * (.FrameProgress.width)
        End With
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        '--------------------------------------
        'the rest of your macro goes below here
        '
        '
        '--------------------------------------
        '(Step 3) Close the progress bar when you're done
        If i = lastRow Then Unload ufProgress
    Next i
End Sub

Sub FractionComplete(pctdone As Single)
    With ufProgress
        .Caption = "Complété à " & pctdone * 100 & "%"
        .LabelProgress.width = pctdone * (.FrameProgress.width)
    End With
    DoEvents
End Sub

Sub Fill_Or_Empty_Range_Background(rng As Range, fill As Boolean, Optional colorIndex As Variant = xlNone)
    If fill Then
        If IsMissing(colorIndex) Or colorIndex = xlNone Then
            rng.Interior.colorIndex = xlColorIndexNone ' Clear the background color
        Else
            rng.Interior.colorIndex = colorIndex ' Fill with specified color
        End If
    Else
        rng.Interior.colorIndex = xlColorIndexNone ' Clear the background color
    End If
End Sub

Sub Tab_Order_Toggle_Mode()

    isTab_Order_Activated = Not isTab_Order_Activated

End Sub

Sub Buttons_Enabled_True_Or_False(clear As Boolean, add As Boolean, _
                                  update As Boolean, delete As Boolean)
    With ufSaisieHeures
        .cmdClear.Enabled = clear
        .cmdAdd.Enabled = add
        .cmdUpdate.Enabled = update
        .cmdDelete.Enabled = delete
    End With

End Sub

Sub Invalid_Date_Message() '2024-03-03 @ 07:45

    MsgBox Prompt:="La valeur saisie ne peut être utilisée comme une date valide", _
        Title:="Validation de la date", _
        Buttons:=vbCritical

End Sub

Sub Erreur_Totaux_DT_CT()

    MsgBox Prompt:="Les totaux (Débit vs. Crédit) sont différents !!!", _
        Title:="Validation des totaux du G/L", _
        Buttons:=vbCritical

End Sub


Sub Pause_Application(s As Double)
    
    If s > 5 Then Stop
    
    Dim endTime As Double
    endTime = Timer + s 'Set end time to 's' seconds from now
    
    Do While Timer < endTime
        'Sleep
    Loop
    
End Sub

