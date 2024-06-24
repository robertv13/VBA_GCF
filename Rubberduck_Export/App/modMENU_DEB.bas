Attribute VB_Name = "modMENU_DEB"
Option Explicit

'Option # 1
Sub DEB_Saisie_Click()
    
    Call SlideIn_Paiement
    
    wshDEB_Saisie.Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    With wshDEB_Saisie
        .Visible = xlSheetVisible
        .Select
    End With
    Application.ScreenUpdating = True

End Sub

