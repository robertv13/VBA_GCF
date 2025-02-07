Attribute VB_Name = "modMENU_TEC"
Option Explicit

'Option # 1
Sub SaisieHeures_Click()

    startTime = Timer: Call Log_Record("modMENU_TEC:SaisieHeures_Click", "", 0)
    
    fromMenu = True '2024-09-03 @ 06:20

    Load ufSaisieHeures
    ufSaisieHeures.show vbModeless '2024-08-08 @ 13:56
    
    Call Log_Record("modMENU_TEC:SaisieHeures_Click", "", startTime)

End Sub

'Option # 2
Sub TEC_TDB_Click()

    startTime = Timer: Call Log_Record("modMENU_TEC:TEC_TDB_Click", "", 0)
    
    fromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_TDB
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

    Call Log_Record("modMENU_TEC:TEC_TDB_Click", "", startTime)

End Sub

'Option # 3
Sub TEC_Analyse_Click()

    startTime = Timer: Call Log_Record("modMENU_TEC:TEC_Analyse_Click", "", 0)
    
    fromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Analyse
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

    Call Log_Record("modMENU_TEC:TEC_Analyse_Click", "", startTime)

End Sub

'Option # 4
Sub TEC_Evaluation_Click()

    startTime = Timer: Call Log_Record("modMENU_TEC:TEC_Evaluation_Click", "", 0)
    
    fromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Evaluation
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modMENU_TEC:TEC_Evaluation_Click", "", startTime)

End Sub

'Option # 5
Sub TEC_Radiation_Click()

    startTime = Timer: Call Log_Record("modMENU_TEC:TEC_Radiation_Click", "", 0)
    
    fromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Radiation
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modMENU_TEC:TEC_Radiation_Click", "", startTime)

End Sub

Sub shp_Get_Deplacements_From_TEC_Click()

    Call Get_Deplacements_From_TEC
    
End Sub



