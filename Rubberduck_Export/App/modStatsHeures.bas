Attribute VB_Name = "modStatsHeures"
Option Explicit

Sub StatsHeures_AdvancedFilters()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modStatsHeures:StatsHeures_AdvancedFilters", 0)

    'Voir la feuille TEC_TDB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    Dim lastResultRow As Long
    Dim rngResult As Range
    Dim rngCriteria As Range
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'AdvancedFilter # 1 - Semaine
    
    'Définir le range des résultats et effacer avant le traitement
    Set rngResult = ws.Range("W1").CurrentRegion
    rngResult.Offset(1, 0).Clear
    Set rngResult = ws.Range("W1").CurrentRegion
    
    'Définir le range des critères
    Set rngCriteria = ws.Range("S1:U2")
    
    ws.Range("tblTEC_TDB_Data[#All]").AdvancedFilter _
        action:=xlFilterCopy, _
        criteriaRange:=rngCriteria, _
        CopyToRange:=rngResult, _
        Unique:=False
        
    'Tri des informations
    lastResultRow = ws.Cells(ws.rows.count, "W").End(xlUp).Row
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On ProfID
            .SortFields.Add key:=ws.Range("W2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.Add key:=ws.Range("Y2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.Add key:=ws.Range("Z2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("W2:AD" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If

    'AdvancedFilter # 2 - Mois
    
    'Définir le range des résultats et effacer avant le traitement
    Set rngResult = ws.Range("AJ1").CurrentRegion
    rngResult.Offset(1, 0).Clear
    Set rngResult = ws.Range("AJ1").CurrentRegion
    
    'Définir le range des critères
    Set rngCriteria = ws.Range("AF1:AH2")
    
    ws.Range("tblTEC_TDB_Data[#All]").AdvancedFilter _
        action:=xlFilterCopy, _
        criteriaRange:=rngCriteria, _
        CopyToRange:=rngResult, _
        Unique:=False
        
    'Tri des informations
    lastResultRow = ws.Cells(ws.rows.count, "AJ").End(xlUp).Row
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On ProfID
            .SortFields.Add key:=ws.Range("AJ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.Add key:=ws.Range("AL2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.Add key:=ws.Range("AM2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("AJ2:AQ" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If

    'AdvancedFilter # 3 - Trimestre
    
    'Définir le range des résultats et effacer avant le traitement
    Set rngResult = ws.Range("AW1").CurrentRegion
    rngResult.Offset(1, 0).Clear
    Set rngResult = ws.Range("AW1").CurrentRegion
    
    'Définir le range des critères
    Set rngCriteria = ws.Range("AS1:AU2")
    
    ws.Range("tblTEC_TDB_Data[#All]").AdvancedFilter _
        action:=xlFilterCopy, _
        criteriaRange:=rngCriteria, _
        CopyToRange:=rngResult, _
        Unique:=False
        
    'Tri des informations
    lastResultRow = ws.Cells(ws.rows.count, "AW").End(xlUp).Row
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On ProfID
            .SortFields.Add key:=ws.Range("AW2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.Add key:=ws.Range("AY2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.Add key:=ws.Range("AZ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("AW2:BD" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If
    
    'AdvancedFilter # 4 - Année financière
    
    'Définir le range des résultats et effacer avant le traitement
    Set rngResult = ws.Range("BJ1").CurrentRegion
    rngResult.Offset(1, 0).Clear
    Set rngResult = ws.Range("BJ1").CurrentRegion
    
    'Définir le range des critères
    Set rngCriteria = ws.Range("BF1:BH2")
    
    ws.Range("tblTEC_TDB_Data[#All]").AdvancedFilter _
        action:=xlFilterCopy, _
        criteriaRange:=rngCriteria, _
        CopyToRange:=rngResult, _
        Unique:=False
        
    'Tri des informations
    lastResultRow = ws.Cells(ws.rows.count, "BJ").End(xlUp).Row
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On ProfID
            .SortFields.Add key:=ws.Range("BJ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.Add key:=ws.Range("BL2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.Add key:=ws.Range("BM2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("BJ2:BQ" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Call Log_Record("modStatsHeures:StatsHeures_AdvancedFilters", startTime)

End Sub

Sub Stats_Back_To_ufSaisieHeures()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modStatsHeures:Stats_Back_To_ufSaisieHeures", 0)
   
    wshStatsHeuresPivotTables.Visible = xlSheetHidden
    
    ufSaisieHeures.show vbModeless

    Call Log_Record("modStatsHeures:Stats_Back_To_ufSaisieHeures()", startTime)

End Sub


