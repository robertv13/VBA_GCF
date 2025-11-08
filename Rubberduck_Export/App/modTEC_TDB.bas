Attribute VB_Name = "modTEC_TDB"
Option Explicit

Sub shpRetournerAuMenu_Click()

    Call RetournerAuMenu
    
End Sub

Sub RetournerAuMenu()

    Call modAppli.QuitterFeuillePourMenu(wshMenuTEC, True)

End Sub

Sub ActualiserTECTableauDeBord()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_TDB:ActualiserTECTableauDeBord", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Call RafraichirTableauDeBordTEC
    Call RafraichirTableauxCroisesTEC
    
    Call AjusterBordurePivotTable
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_TDB:ActualiserTECTableauDeBord", vbNullString, startTime)

End Sub

Sub AjusterBordurePivotTable() '2025-02-01 @ 05:49

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("wshTEC_TDB:AjusterBordurePivotTable", vbNullString, 0)
    
    Dim ws As Worksheet: Set ws = wshTEC_TDB
    
    Dim lastUsedRow As Long: lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastUsedRow <= 10 Then Exit Sub
    
    Dim rng As Range
    Set rng = ws.Range("A10:B" & lastUsedRow - 1) 'Exclure la ligne TOTAL
    
    Dim su As Boolean
    su = Application.ScreenUpdating
    Dim ee As Boolean
    ee = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Call AppliquerBordures(rng)
    
    Set rng = ws.Range("D9")
    If rng.Interior.ThemeColor <> xlThemeColorAccent4 Then
        With rng.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
    End If

    'Libérer la mémoire
    Set rng = Nothing
    Set ws = Nothing

    Application.EnableEvents = ee
    Application.ScreenUpdating = su

    Call modDev_Utils.EnregistrerLogApplication("wshTEC_TDB:AjusterBordurePivotTable", vbNullString, startTime)

End Sub

Public Sub AppliquerBordures(rng As Range) '2025-10-31 @ 08:04

    If rng Is Nothing Then Exit Sub

    Dim b As Variant
    For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With rng.Borders(b)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next b

    For Each b In Array(xlInsideVertical, xlInsideHorizontal)
        With rng.Borders(b)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
    Next b
    
End Sub

