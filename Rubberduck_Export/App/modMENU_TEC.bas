Attribute VB_Name = "modMENU_TEC"
Option Explicit

'Option # 1
Sub SaisieHeures_Click()

    fromMenu = True '2024-09-03 @ 06:20

    Load ufSaisieHeures
    ufSaisieHeures.show vbModeless '2024-08-08 @ 13:56
    
End Sub

'Option # 2
Sub TEC_TdB_Click()

    fromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_TDB
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

End Sub

'Option # 3
Sub TEC_Analyse_Click()

    fromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Analyse
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

End Sub

'Option # 4
Sub TEC_Evaluation_Click()

    fromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Evaluation
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

End Sub





