Attribute VB_Name = "modMENU_FAC"
Option Explicit

'Option # 1
Sub PreparationFacture_Click()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMENU_FAC:PreparationFacture_Click", "", 0)

    Application.ScreenUpdating = False
    
    Application.EnableEvents = True
    
    fromMenu = True '2024-09-03 @ 06:20
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Finale.Visible = xlSheetVisible
    
    wshFAC_Brouillon.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call Log_Record("modMENU_FAC:PreparationFacture_Click", "", startTime)
    
End Sub

'Option # 2
Sub SuiviCC_Click()

    Application.ScreenUpdating = False
    
    wshCAR_Liste_Agée.Application.Calculation = xlCalculationAutomatic
    
    fromMenu = True '2024-09-03 @ 06:20
    
    With wshCAR_Liste_Agée
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

End Sub

'Option # 3
Sub FAC_Historique_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMENU_FAC:FAC_Historique_Click", "", 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterFacEntete
    Call modImport.ImporterFacDetails
    Call modImport.ImporterFacComptesClients

    Application.EnableEvents = True
    
    fromMenu = True '2024-09-03 @ 06:20

    wshFAC_Interrogation.Visible = xlSheetVisible
    wshFAC_Interrogation.Activate
    
    wshFAC_Interrogation.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call Log_Record("modMENU_FAC:FAC_Historique_Click", "", startTime)

End Sub

'Option # 4
Sub FAC_Confirmation_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMENU_FAC:FAC_Confirmation_Click", "", 0)
    
    fromMenu = True '2024-09-03 @ 06:20

    Call modImport.ImporterClients
    Call modImport.ImporterFacComptesClients
    Call modImport.ImporterFacDetails
    Call modImport.ImporterFacEntete
    Call modImport.ImporterFacSommaireTaux
    Call modImport.ImporterTEC
    
    Call Afficher_ufConfirmation
    
    Call Log_Record("modMENU_FAC:FAC_Confirmation_Click", "", startTime)

End Sub

