Attribute VB_Name = "modMENU_FAC"
Option Explicit

'Option # 1
Sub PreparationFacture_Click()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("wshMenuFAC:PreparationFacture_Click()")

    Call SlideIn_PrepFact
    
    Application.ScreenUpdating = False
    
    Call Client_List_Import_All
    Call FAC_Entête_Import_All
    Call FAC_Détails_Import_All
    Call FAC_Comptes_Clients_Import_All
    Call GL_Trans_Import_All

    Application.EnableEvents = True
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Finale.Visible = xlSheetVisible
    
    wshFAC_Brouillon.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("wshMenuFAC:PreparationFacture_Click()", timerStart)
    
End Sub

'Option # 2
Sub SuiviCC_Click()

    Call SlideIn_SuiviCC
    MsgBox "Ajouter la fonction 'Suivi des C/C'"

End Sub

'Option # 3
Sub Encaissement_Click()

    Call SlideIn_Encaissement
    
    Application.ScreenUpdating = False
    
    Call Encaissement_Import_All
    
'    wshFAC.Visible = xlSheetVisible
'    wshENC_Détails.Visible = xlSheetVisible
'    wshENC_Entête.Visible = xlSheetVisible
    With wshENC_Saisie
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
End Sub

'Option # 4
Sub FAC_Historique_Click()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("wshMenuFAC:FAC_Historique_Click()")

    Call SlideIn_FAC_Historique
    
    Application.ScreenUpdating = False
    
    Call FAC_Entête_Import_All
    Call FAC_Détails_Import_All
    Call FAC_Comptes_Clients_Import_All

    Application.EnableEvents = True
    
    wshFAC_Historique.Visible = xlSheetVisible
    wshFAC_Historique.Activate
    
    wshFAC_Historique.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("wshMenuFAC:FAC_Historique_Click()", timerStart)

End Sub


