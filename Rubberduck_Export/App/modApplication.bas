Attribute VB_Name = "modApplication"
Option Explicit

Global Const gAppVersion As String = "v2.A" '2024-03-19 @ 08:39

Public isTab_Order_Activated As Boolean

Sub BackToMainMenu()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.name <> "Menu" Then ws.Visible = xlSheetHidden
    Next ws
    wshMenu.Activate
    wshMenu.Range("A1").Select

End Sub

Private Sub auto_open() '2024-03-06 @ 14:36

    Dim timerStart As Double: timerStart = Timer

    Call Output_Timer_Results("auto_open()", timerStart)

End Sub

Private Sub auto_close() '2024-03-06 @ 14:36

    Dim timerStart As Double: timerStart = Timer

    MsgBox "Auto_Close..."
    
    Call Output_Timer_Results("auto_close()", timerStart)
    
End Sub

Sub Dynamic_Range_Redefine_Plan_Comptable() '2024-03-06 @ 13:43
    
    Dim timerStart As Double: timerStart = Timer

    'Delete existing dynamic named range (assuming it exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptableDescription").delete
    On Error GoTo 0
    
    'Define a new dynamic named range for 'Plan Comptable'
    Dim newRangeFormula As String
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,1)"
    
    'Create a new dynamic named range
    ThisWorkbook.Names.add name:="dnrPlanComptableDescription", RefersTo:=newRangeFormula
    
    Call Output_Timer_Results("Dynamic_Range_Redefine_Plan_Comptable()", timerStart)
    
End Sub

Sub Hide_All_Worksheets_Except_Menu() '2024-02-20 @ 07:28
    
    Dim timerStart As Double: timerStart = Timer
    
    Dim wsh As Worksheet
    For Each wsh In ThisWorkbook.Worksheets
        If wsh.codeName <> "wshMenu" And _
            InStr(wsh.codeName, "wshzDoc") = 0 Then
                wsh.Visible = xlSheetHidden
        End If
    Next wsh
    
    Call Output_Timer_Results("Hide_All_Worksheets_Except_Menu()", timerStart)
    
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

Sub Slide_In_All_Menu_Options()

    Dim timerStart As Double: timerStart = Timer
    
    SlideIn_TEC
    SlideIn_Facturation
    SlideIn_Debours
    SlideIn_Comptabilite
    SlideIn_Parametres
    SlideIn_Exit

    Call Output_Timer_Results("Slide_In_All_Menu_Options()", timerStart)

End Sub

'Sub Add_Caption_To_Userform(uf As UserForm, titleText As String)
'
'    uf.Caption = titleText
'
'End Sub
'
'Sub Add_Label_To_Userform(uf As UserForm, labelText As String, leftPos As Single, topPos As Single)
'
'    'Add a label to the userform
'    Dim newLabel As MSForms.Label: Set newLabel = uf.Controls.add("Forms.Label.1")
'
'    With newLabel
'        .Caption = labelText
'        .Left = leftPos
'        .width = 150
'        .Top = topPos
'        'Set other properties as needed
'    End With
'
'End Sub
'

