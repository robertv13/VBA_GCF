Attribute VB_Name = "modMENU_DEB"
Option Explicit

'Option # 1
Sub DEB_Saisie_Click()
    
    Call SlideIn_Paiement
    
    Application.ScreenUpdating = False
    
    Call Fournisseur_List_Import_All
    
    Application.EnableEvents = True
    
    With wshDEB_Saisie
        .Visible = xlSheetVisible
        .Activate
    End With
    
    wshDEB_Saisie.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

End Sub

