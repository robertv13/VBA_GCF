Attribute VB_Name = "modMENU_FAC"
Option Explicit

'Option # 1
Sub PreparationFacture_Click()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("wshMenuFAC:PreparationFacture_Click", 0)

    Application.ScreenUpdating = False
    
    Application.EnableEvents = True
    
    fromMenu = True '2024-09-03 @ 06:20
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Finale.Visible = xlSheetVisible
    
    wshFAC_Brouillon.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call Log_Record("wshMenuFAC:PreparationFacture_Click", startTime)
    
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshMenuFAC:FAC_Historique_Click", 0)

    Application.ScreenUpdating = False
    
    Call FAC_Entête_Import_All
    Call FAC_Détails_Import_All
    Call FAC_Comptes_Clients_Import_All

    Application.EnableEvents = True
    
    fromMenu = True '2024-09-03 @ 06:20

    wshFAC_Historique.Visible = xlSheetVisible
    wshFAC_Historique.Activate
    
    wshFAC_Historique.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call Log_Record("wshMenuFAC:FAC_Historique_Click", startTime)

End Sub

'Option # 4
Sub FAC_Confirmation_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshMenuFAC:FAC_Confirmation_Click", 0)

    Application.ScreenUpdating = False
    
    'Import data files from MASTER
    Call FAC_Comptes_Clients_Import_All
    Call FAC_Entête_Import_All
    Call FAC_Détails_Import_All

    Application.EnableEvents = True
    
    fromMenu = True '2024-09-03 @ 06:20
    
    wshFAC_Confirmation.Visible = xlSheetVisible
    wshFAC_Confirmation.Activate
    
    wshFAC_Confirmation.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

    Call Log_Record("wshMenuFAC:FAC_Confirmation_Click", startTime)

End Sub


