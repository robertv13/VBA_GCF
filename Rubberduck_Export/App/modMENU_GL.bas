Attribute VB_Name = "modMENU_GL"
Option Explicit

'Option # 1
Sub EJ_Saisie_Click()
    
    Application.ScreenUpdating = False
    
    Call SlideIn_BV
    Call SlideIn_EJ
    
    wshGL_EJ.Application.Calculation = xlCalculationAutomatic
    With wshGL_EJ
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

End Sub

'Option # 2
Sub BV_Click()
    
    Application.ScreenUpdating = False
    
    Call SlideIn_EJ
    Call SlideIn_BV
    Call SlideIn_GL_Report
    
    With wshGL_BV
        .Visible = xlSheetVisible
        .Activate
    End With
    
    Application.ScreenUpdating = True

End Sub

'Option # 3
Sub Rapport_GL_Click()

    Application.ScreenUpdating = False

    Call SlideIn_BV
    Call SlideIn_GL_Report
    Call SlideIn_EF
    
    With wshGL_Rapport
        .Visible = xlSheetVisible
        .Select
    End With

    Application.ScreenUpdating = True

End Sub

'Option # 4
Sub EF_Click()

    Application.ScreenUpdating = False

    Call SlideIn_GL_Report
    Call SlideIn_EF
    MsgBox "Ajouter la fonction 'États Financiers'"

    Application.ScreenUpdating = True

End Sub


