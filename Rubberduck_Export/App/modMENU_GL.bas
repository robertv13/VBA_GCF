Attribute VB_Name = "modMENU_GL"
Option Explicit

'Option # 1
Sub DEB_Saisie_Click()
    
    Application.ScreenUpdating = False
    
    Call Fournisseur_List_Import_All
    
    Application.EnableEvents = True
    
    fromMenu = True '2024-09-30 @ 09:33
    
    With wshDEB_Saisie
        .Visible = xlSheetVisible
        .Activate
    End With
    
    wshDEB_Saisie.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

End Sub

'Option # 2
Sub Encaissement_Click()

    Application.ScreenUpdating = False
    
    fromMenu = True '2024-09-03 @ 06:20
    
    With wshENC_Saisie
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
End Sub

'Option # 3
Sub EJ_Saisie_Click()
    
    Application.ScreenUpdating = False
    
    wshGL_EJ.Application.Calculation = xlCalculationAutomatic
    
    With wshGL_EJ
        .Visible = xlSheetVisible
        .Activate
        .Select
    End With
    
    Application.ScreenUpdating = True
    
    Application.EnableEvents = True

End Sub

'Option # 4
Sub BV_Click()
    
    Application.ScreenUpdating = False
    
    With wshGL_BV
        .Visible = xlSheetVisible
        .Activate
    End With
    
    Application.ScreenUpdating = True

End Sub

'Option # 5
Sub Rapport_GL_Click()

    Application.ScreenUpdating = False

    With wshGL_Rapport
        .Visible = xlSheetVisible
        .Select
    End With

    Application.ScreenUpdating = True

End Sub

'Option # 6
Sub EF_Click()

    Application.ScreenUpdating = False

    MsgBox "Ajouter la fonction 'États Financiers'"

    Application.ScreenUpdating = True

End Sub


