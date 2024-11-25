Attribute VB_Name = "modStatsHeures"
Option Explicit

Sub Stats_Heures_AF()
    
    'La cellule 'S7' doit contenir le Professionnel
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modStatsHeures:Stats_Heures_AF", 0)

    'On utilise la feuille TEC_TDB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    
    'Les même objects seront utilisés avec les 4 AdvancedFilters
    Dim rngData As Range
    Dim rngCriteria As Range
    Dim rngResult As Range
    Dim lastResultRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'AdvancedFilter # 1 - Semaine
    
    'Effacer les données de la dernière utilisation
    ws.Range("T16:T20").ClearContents
    ws.Range("T16").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
    ws.Range("T17").value = rngData.Address
    
    'Définir le range des critères
    Set rngCriteria = ws.Range("S2:U3")
    ws.Range("T18").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Set rngResult = ws.Range("W1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("W1").CurrentRegion
    ws.Range("T19").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Tri des informations
    lastResultRow = ws.Cells(ws.Rows.count, "W").End(xlUp).row
    ws.Range("T20").value = lastResultRow - 1 & " lignes"
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On ProfID
            .SortFields.add key:=ws.Range("W2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.add key:=ws.Range("Y2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.add key:=ws.Range("Z2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("W2:AD" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If

    'AdvancedFilter # 2 - Mois
    
    'Effacer les données de la dernière utilisation
    ws.Range("AG16:AG20").ClearContents
    ws.Range("AG16").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
    ws.Range("AG17").value = rngData.Address
    
    'Définir le range des critères
    Set rngCriteria = ws.Range("AF2:AH3")
    ws.Range("AG18").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Set rngResult = ws.Range("AJ1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("AJ1").CurrentRegion
    ws.Range("AG19").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Tri des informations
    lastResultRow = ws.Cells(ws.Rows.count, "AJ").End(xlUp).row
    ws.Range("AG20").value = lastResultRow - 1 & " lignes"
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On ProfID
            .SortFields.add key:=ws.Range("AJ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.add key:=ws.Range("AL2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.add key:=ws.Range("AM2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("AJ2:AQ" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If

    'AdvancedFilter # 3 - Trimestre
    
    'Effacer les données de la dernière utilisation
    ws.Range("AT16:AT20").ClearContents
    ws.Range("AT16").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
    ws.Range("AT17").value = rngData.Address
    
    'Définir le range des critères
    Set rngCriteria = ws.Range("AS2:AU3")
    ws.Range("AT18").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Set rngResult = ws.Range("AW1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("AW1").CurrentRegion
    ws.Range("AT19").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Tri des informations
    lastResultRow = ws.Cells(ws.Rows.count, "AW").End(xlUp).row
    ws.Range("AT20").value = lastResultRow - 1 & " lignes"
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On ProfID
            .SortFields.add key:=ws.Range("AW2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.add key:=ws.Range("AY2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.add key:=ws.Range("AZ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("AW2:BD" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If
    
    'AdvancedFilter # 4 - Année financière
    
    'Effacer les données de la dernière utilisation
    ws.Range("BG16:BG20").ClearContents
    ws.Range("BG16").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
    ws.Range("BG17").value = rngData.Address
    
    'Définir le range des critères
    Set rngCriteria = ws.Range("BF2:BH3")
    ws.Range("BG18").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Set rngResult = ws.Range("BJ1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("BJ1").CurrentRegion
    ws.Range("BG19").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Tri des informations
    lastResultRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).row
    ws.Range("BG20").value = lastResultRow - 1 & " lignes"
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On ProfID
            .SortFields.add key:=ws.Range("BJ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.add key:=ws.Range("BL2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.add key:=ws.Range("BM2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("BJ2:BQ" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modStatsHeures:Stats_Heures_AF", startTime)

End Sub

Sub Stats_Back_To_ufSaisieHeures()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modStatsHeures:Stats_Back_To_ufSaisieHeures", 0)
   
    wshStatsHeuresPivotTables.Visible = xlSheetHidden
    
    ufSaisieHeures.show vbModeless

    Call Log_Record("modStatsHeures:Stats_Back_To_ufSaisieHeures", startTime)

End Sub


