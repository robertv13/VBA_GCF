Attribute VB_Name = "modMENU_FAC"
Option Explicit

'Option # 1
Sub shpAccederPreparationFacture_Click()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_FAC:PreparationFacture_Click", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Application.EnableEvents = True
    
    gFromMenu = True
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Finale.Visible = xlSheetVisible
    
    wshFAC_Brouillon.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modMENU_FAC:PreparationFacture_Click", vbNullString, startTime)
    
End Sub

'Option # 2
Sub shpAccederListeAgeeCC_Click()

    Application.ScreenUpdating = False
    
    wshCAR_Liste_Agee.Application.Calculation = xlCalculationAutomatic
    
    gFromMenu = True
    
    With wshCAR_Liste_Agee
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

End Sub

'Option # 3
Sub shpAccederInterrogationFacture_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_FAC:FAC_Historique_Click", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Call modImport.ImporterFacEntete
    Call modImport.ImporterFacDetails
    Call modImport.ImporterFacComptesClients

    Application.EnableEvents = True
    
    gFromMenu = True

    wshFAC_Interrogation.Visible = xlSheetVisible
    wshFAC_Interrogation.Activate
    
    wshFAC_Interrogation.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modMENU_FAC:FAC_Historique_Click", vbNullString, startTime)

End Sub

'Option # 4
Sub shpAccederConfirmationFacture_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_FAC:FAC_Confirmation_Click", vbNullString, 0)
    
    gFromMenu = True

    Call modImport.ImporterClients
    Call modImport.ImporterFacComptesClients
    Call modImport.ImporterFacDetails
    Call modImport.ImporterFacEntete
    Call modImport.ImporterFacSommaireTaux
    Call modImport.ImporterGLTransactions
    Call modImport.ImporterTEC
    
    Call AfficherFormulaireConfirmation
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU_FAC:FAC_Confirmation_Click", vbNullString, startTime)

End Sub

Public Sub shpReinitialiserFormesManuellement_Click() '2025-11-02 @ 09:20

    Call ReinitialiserFormesManuellement

End Sub


