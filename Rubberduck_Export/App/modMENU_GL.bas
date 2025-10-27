Attribute VB_Name = "modMENU_GL"
Option Explicit

'Option # 1
Sub shpAccederSaisieENC_Click()

    Application.ScreenUpdating = False
    
    gFromMenu = True '2024-09-03 @ 06:20
    
    With wshENC_Saisie
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
End Sub

'Option # 2
Sub shpAccederSaisieDEB_Click()
    
    Application.ScreenUpdating = False
    
    Application.EnableEvents = True
    
    gFromMenu = True '2024-09-30 @ 09:33
    
    With wshDEB_Saisie
        .Visible = xlSheetVisible
        .Activate
    End With
    
    wshDEB_Saisie.Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

End Sub

'Option # 3
Sub shpAccederSaisieEJ_Click()
    
    Application.ScreenUpdating = False
    
    wshGL_EJ.Application.Calculation = xlCalculationAutomatic
    
    gFromMenu = True '2025-10-26 @ 07:59
    
    With wshGL_EJ
        .Visible = xlSheetVisible
        .Activate
        .Select
    End With
    
    Application.ScreenUpdating = True
    
    Application.EnableEvents = True

End Sub

'Option # 4
Sub shpAccederBV_Click()
    
    Application.ScreenUpdating = False
    
    gFromMenu = True '2025-10-26 @ 07:59
    
    With wshGL_BV
        .Visible = xlSheetVisible
        .Activate
    End With
    
    Application.ScreenUpdating = True

End Sub

'Option # 5
Sub shpAccederRapportTransGL_Click()

    gFromMenu = True '2025-10-26 @ 07:59
    
    ufGL_Rapport.show 'vbModal
    
End Sub

'Option # 6
Sub shpAccederEF_Click()

    Application.ScreenUpdating = False

    gFromMenu = True '2025-10-26 @ 07:59
    
    With wshGL_PrepEF
        .Visible = xlSheetVisible
        .Activate
    End With

    Application.ScreenUpdating = True

End Sub

'Option # 7
Sub shpAccederStatsCA_Click()

    Application.ScreenUpdating = False

    gFromMenu = True '2025-10-26 @ 07:59
    
    With wshGL_Stats_CA
        .Visible = xlSheetVisible
        .Activate
    End With

    Application.ScreenUpdating = True

End Sub

