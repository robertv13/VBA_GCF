Attribute VB_Name = "modMENU_GL"
Option Explicit

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


