Attribute VB_Name = "modStatsHeures"
Option Explicit

Sub CreatePivotWithSpecificSorting()

    Dim ws As Worksheet
    Dim wsPivot As Worksheet
    Dim filterRange As Range
    Dim resultRange As Range
    Dim lastRow As Long
    Dim pivotCache As pivotCache
    Dim pivotTable As pivotTable
    
    Set ws = wshTEC_TDB_Data ' Feuille contenant les données filtrées
    Set wsPivot = ThisWorkbook.Sheets("PivotSheet") ' Feuille où le tableau croisé sera créé
    
    ' Appliquer l'AdvancedFilter
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Set filterRange = ws.Range("A1:Q" & lastRow) ' Plage de données avec W comme colonne de tri
    
    ' Result Range
    Set resultRange = ws.Range("W1").CurrentRegion
    resultRange.Offset(1, 0).Clear
    Set resultRange = ws.Range("W1").CurrentRegion
    
    ' Utiliser AdvancedFilter ici
    filterRange.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=ws.Range("S1:U2"), _
                CopyToRange:=ws.Range("W1:AD1"), _
                Unique:=False
    
    ' Définir la plage des résultats filtrés en excluant la colonne W
    lastRow = ws.Cells(ws.rows.count, "W").End(xlUp).Row ' Supposons que les résultats sont à partir de AD
    Set resultRange = ws.Range("W1:AD" & lastRow)
    
    'Supprimer tout ancien PivotTable
    RemoveExistingPivotTable wsPivot, "FilteredPivot"
    
    ' Créer un cache PivotTable en utilisant les colonnes excluant W
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=resultRange.Address)
    
    ' Créer le tableau croisé dynamique à partir des résultats filtrés
    Set pivotTable = wsPivot.PivotTables.Add( _
        pivotCache:=pivotCache, TableDestination:=wsPivot.Range("A3"), TableName:="FilteredPivot")
    
    ' Configurer les champs du tableau croisé dynamique
    With pivotTable
        .PivotFields("Prof").Orientation = xlRowField
        .PivotFields("Date").Orientation = xlRowField
        
        ' Configurer le champ de valeurs
        With .PivotFields("H_N_D")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00" ' Appliquer format nombre avec 2 décimales
            .Position = 1 ' Facultatif : définir la position du champ
        End With
        
        ' Changer le libellé de l'en-tête après avoir ajouté le champ
        .PivotFields("H_N_D").Caption = "Hres/Nettes" ' Nouveau libellé pour l'en-tête
        
        ' Désactiver le tri automatique pour respecter l'ordre filtré
        .PivotFields("Prof").AutoSort xlManual, .PivotFields("Prof").SourceName
    End With
    
    ' Actualiser le tableau croisé dynamique
    pivotTable.RefreshTable
    
    Set filterRange = Nothing
    Set pivotCache = Nothing
    Set pivotTable = Nothing
    Set resultRange = Nothing
    Set ws = Nothing
    Set wsPivot = Nothing

End Sub

Sub RemoveExistingPivotTable(wsPivot As Worksheet, pivotTableName As String)
    Dim pt As pivotTable
    On Error Resume Next
    Set pt = wsPivot.PivotTables(pivotTableName)
    On Error GoTo 0

    If Not pt Is Nothing Then
        pt.TableRange2.Clear ' Cela supprime les données du PivotTable
        pt.RefreshTable ' Cela actualise le tableau croisé dynamique
    End If
End Sub

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

