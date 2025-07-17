Attribute VB_Name = "modMENU_TEC"
Option Explicit

'Option # 1
Sub SaisieHeures_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:SaisieHeures_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Load ufSaisieHeures
    ufSaisieHeures.show vbModeless '2024-08-08 @ 13:56
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:SaisieHeures_Click", vbNullString, startTime)

End Sub

'Option # 2
Sub TEC_TDB_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:TEC_TDB_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_TDB
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:TEC_TDB_Click", vbNullString, startTime)

End Sub

'Option # 3
Sub TEC_Analyse_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:TEC_Analyse_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Analyse
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:TEC_Analyse_Click", vbNullString, startTime)

End Sub

'Option # 4
Sub TEC_Evaluation_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:TEC_Evaluation_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Evaluation
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:TEC_Evaluation_Click", vbNullString, startTime)

End Sub

'Option # 5
Sub TEC_Radiation_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:TEC_Radiation_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Radiation
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:TEC_Radiation_Click", vbNullString, startTime)

End Sub

Sub shp_Get_Deplacements_From_TEC_Click()

    Call Get_Deplacements_From_TEC
    
End Sub

