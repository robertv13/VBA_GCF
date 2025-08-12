Attribute VB_Name = "modMENU_TEC"
Option Explicit

'Option # 1
Sub shpAccederSaisieHeures_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:shpAccederSaisieHeures_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Load ufSaisieHeures
    ufSaisieHeures.show vbModeless '2024-08-08 @ 13:56
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:shpAccederSaisieHeures_Click", vbNullString, startTime)

End Sub

'Option # 2
Sub shpAccederTECTDB_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:AccederTECTDB_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_TDB
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:AccederTECTDB_Click", vbNullString, startTime)

End Sub

'Option # 3
Sub shpAccederProjetFacture_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:shpAccederProjetFacture_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Analyse
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:shpAccederProjetFacture_Click", vbNullString, startTime)

End Sub

'Option # 4
Sub shpAccederEvaluationTEC_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:shpAccederEvaluationTEC_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Evaluation
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:shpAccederEvaluationTEC_Click", vbNullString, startTime)

End Sub

'Option # 5
Sub shpAccederRadiationTEC_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:shpAccederRadiationTEC_Click", vbNullString, 0)
    
    gFromMenu = True '2024-09-03 @ 06:20

    Application.ScreenUpdating = False
    
    wshTEC_TDB.Application.Calculation = xlCalculationAutomatic
    
    With wshTEC_Radiation
        .Visible = xlSheetVisible
        .Select
    End With
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modMENU_TEC:shpAccederRadiationTEC_Click", vbNullString, startTime)

End Sub

Sub shpListeDesDeplacements_Click()

    Call ObtenirDeplacementsAPartirDesTEC
    
End Sub

