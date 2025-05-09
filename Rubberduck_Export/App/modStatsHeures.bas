Attribute VB_Name = "modStatsHeures"
Option Explicit

Sub Stats_Heures_AF()
    
    'La cellule 'S7' doit contenir le Professionnel
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modStatsHeures:Stats_Heures_AF", "", 0)

    'On utilise la feuille TEC_TDB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Exécuter les 4 filtres AdvancedFilter
    Call ExecuterAdvancedFilter(ws, "S2:U3", "W1", "T10:T14", Array("W2", "Y2", "Z2"), "W2:AD")
    Call ExecuterAdvancedFilter(ws, "AF2:AH3", "AJ1", "AG10:AG14", Array("AJ2", "AL2", "AM2"), "AJ2:AQ")
    Call ExecuterAdvancedFilter(ws, "AS2:AU3", "AW1", "AT10:AT14", Array("AW2", "AY2", "AZ2"), "AW2:BD")
    Call ExecuterAdvancedFilter(ws, "BF2:BH3", "BJ1", "BG10:BG14", Array("BJ2", "BL2", "BM2"), "BJ2:BQ")
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modStatsHeures:Stats_Heures_AF", "", startTime)

End Sub

Sub ExecuterAdvancedFilter(ws As Worksheet, criteriaRange As String, resultStartCell As String, logRange As String, sortKeys As Variant, sortRange As String)

    Dim rngData As Range, rngCriteria As Range, rngResult As Range
    Dim lastResultRow As Long
    
    'Journaliser le temps de traitement
    ws.Range(logRange).ClearContents
    ws.Range(logRange).Cells(1, 1).value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range source des données
    Set rngData = ws.Range("l_tbl_TEC_TDB_data[#All]")
    ws.Range(logRange).Cells(2, 1).value = rngData.Address
    
    'Définir les critères
    Set rngCriteria = ws.Range(criteriaRange)
    ws.Range(logRange).Cells(3, 1).value = rngCriteria.Address
    
    'Effacer les résultats précédents
    Set rngResult = ws.Range(resultStartCell).CurrentRegion
    If rngResult.Rows.count > 1 Then
        rngResult.offset(1, 0).Clear
    End If
    Set rngResult = ws.Range(resultStartCell).CurrentRegion
    ws.Range(logRange).Cells(4, 1).value = rngResult.Address
    
    'Appliquer AdvancedFilter
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
    
    'Compter les lignes
    lastResultRow = ws.Cells(ws.Rows.count, rngResult.Cells(1, 1).Column).End(xlUp).row
    ws.Range(logRange).Cells(5, 1).value = lastResultRow - 1 & " lignes"
    
    'Trier les résultats
    Dim i As Long
    If lastResultRow > 2 Then
        With ws.Sort
            .SortFields.Clear
            For i = LBound(sortKeys) To UBound(sortKeys)
                .SortFields.Add key:=ws.Range(sortKeys(i)), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
            Next i
            .SetRange ws.Range(sortRange & lastResultRow)
            .Header = xlYes
            .Apply
        End With
    End If

End Sub

Sub shp_Back_To_ufSaisieHeures_Click()

    Call Back_To_ufSaisieHeures
    
End Sub

Sub Back_To_ufSaisieHeures()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modStatsHeures:Back_To_ufSaisieHeures", "", 0)
   
    wshStatsHeuresPivotTables.Visible = xlSheetHidden
    
    ufSaisieHeures.show vbModeless

    Call Log_Record("modStatsHeures:Back_To_ufSaisieHeures", "", startTime)

End Sub

