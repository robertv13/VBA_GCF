Attribute VB_Name = "modTEC_TDB"
Option Explicit

Sub shpTEC_TDB_BackToMenu_Click()

    Call TEC_TDB_BackToMenu

End Sub

Sub TEC_TDB_BackToMenu()

    wshTEC_TDB.Visible = xlSheetHidden
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select

End Sub

Sub shpActualiserTEC_TDB_Click()

    Call ActualiserTEC_TDB

End Sub

Sub ActualiserTEC_TDB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_TDB:ActualiserTEC_TDB", "", 0)
    
    Application.ScreenUpdating = False
    
    Call TEC_Update_TDB_From_TEC_Local
    Call TEC_TdB_Refresh_All_Pivot_Tables
    
    Call AjusteBordurePivotTable
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modTEC_TDB:ActualiserTEC_TDB", "", startTime)

End Sub

Sub AjusteBordurePivotTable() '2025-02-01 @ 05:49

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshTEC_TDB:AjusteBordurePivotTable", "", 0)
    
    Dim ws As Worksheet
    Set ws = wshTEC_TDB
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    Dim rng As Range
    Set rng = ws.Range("A10:B" & lastUsedRow - 1) 'Exclure la ligne TOTAL
    
    With rng
        'Bordures extérieures (4)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        'Bordures intérieures (2)
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
    End With
    
    Set rng = ws.Range("D9")
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

    'Libérer la mémoire
    Set rng = Nothing
    Set ws = Nothing

    Call Log_Record("wshTEC_TDB:AjusteBordurePivotTable", "", startTime)

End Sub

