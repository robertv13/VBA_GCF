Attribute VB_Name = "modStatsHeures"
Option Explicit

Sub Stats_Heures_AF()
    
    'La cellule 'S7' doit contenir le Professionnel
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modStatsHeures:Stats_Heures_AF", "", 0)

    'On utilise la feuille TEC_TDB_Data
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Ex�cuter les 4 filtres AdvancedFilter
    Call ExecuterAdvancedFilter(ws, "S2:U3", "W1", "T10:T14", Array("W2", "Y2", "Z2"), "W2:AD")
    Call ExecuterAdvancedFilter(ws, "AF2:AH3", "AJ1", "AG10:AG14", Array("AJ2", "AL2", "AM2"), "AJ2:AQ")
    Call ExecuterAdvancedFilter(ws, "AS2:AU3", "AW1", "AT10:AT14", Array("AW2", "AY2", "AZ2"), "AW2:BD")
    Call ExecuterAdvancedFilter(ws, "BF2:BH3", "BJ1", "BG10:BG14", Array("BJ2", "BL2", "BM2"), "BJ2:BQ")
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
'    'Les m�me objects seront utilis�s avec les 4 AdvancedFilters
'    Dim rngData As Range
'    Dim rngCriteria As Range
'    Dim rngResult As Range
'    Dim lastResultRow As Long
'
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'
'    'AdvancedFilter # 1 - Semaine
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("T10:T14").ClearContents
'    ws.Range("T10").value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
'    ws.Range("T11").value = rngData.Address
'
'    'D�finir le range des crit�res
'    Set rngCriteria = ws.Range("S2:U3")
'    ws.Range("T12").value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Set rngResult = ws.Range("W1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("W1").CurrentRegion
'    ws.Range("T13").value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Tri des informations
'    lastResultRow = ws.Cells(ws.Rows.count, "W").End(xlUp).row
'    ws.Range("T14").value = lastResultRow - 1 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 2 Then
'        With ws.Sort 'Sort - ID, Date, TecID
'            .SortFields.Clear
'            'First sort On ProfID
'            .SortFields.Add key:=ws.Range("W2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On Date
'            .SortFields.Add key:=ws.Range("Y2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("Z2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange wshTEC_Local.Range("W2:AD" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
'
'    'AdvancedFilter # 2 - Mois
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("AG10:AG14").ClearContents
'    ws.Range("AG10").value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
'    ws.Range("AG11").value = rngData.Address
'
'    'D�finir le range des crit�res
'    Set rngCriteria = ws.Range("AF2:AH3")
'    ws.Range("AG12").value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Set rngResult = ws.Range("AJ1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("AJ1").CurrentRegion
'    ws.Range("AG13").value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Tri des informations
'    lastResultRow = ws.Cells(ws.Rows.count, "AJ").End(xlUp).row
'    ws.Range("AG14").value = lastResultRow - 1 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 2 Then
'        With ws.Sort 'Sort - ID, Date, TecID
'            .SortFields.Clear
'            'First sort On ProfID
'            .SortFields.Add key:=ws.Range("AJ2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On Date
'            .SortFields.Add key:=ws.Range("AL2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("AM2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange wshTEC_Local.Range("AJ2:AQ" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
'
'    'AdvancedFilter # 3 - Trimestre
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("AT10:AT16").ClearContents
'    ws.Range("AT10").value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
'    ws.Range("AT11").value = rngData.Address
'
'    'D�finir le range des crit�res
'    Set rngCriteria = ws.Range("AS2:AU3")
'    ws.Range("AT12").value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Set rngResult = ws.Range("AW1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("AW1").CurrentRegion
'    ws.Range("AT13").value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Tri des informations
'    lastResultRow = ws.Cells(ws.Rows.count, "AW").End(xlUp).row
'    ws.Range("AT14").value = lastResultRow - 1 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 2 Then
'        With ws.Sort 'Sort - ID, Date, TecID
'            .SortFields.Clear
'            'First sort On ProfID
'            .SortFields.Add key:=ws.Range("AW2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On Date
'            .SortFields.Add key:=ws.Range("AY2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("AZ2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange wshTEC_Local.Range("AW2:BD" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
'
'    'AdvancedFilter # 4 - Ann�e financi�re
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("BG10:BG14").ClearContents
'    ws.Range("BG10").value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
'    ws.Range("BG11").value = rngData.Address
'
'    'D�finir le range des crit�res
'    Set rngCriteria = ws.Range("BF2:BH3")
'    ws.Range("BG12").value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Set rngResult = ws.Range("BJ1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("BJ1").CurrentRegion
'    ws.Range("BG13").value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Tri des informations
'    lastResultRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).row
'    ws.Range("BG14").value = lastResultRow - 1 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 2 Then
'        With ws.Sort 'Sort - ID, Date, TecID
'            .SortFields.Clear
'            'First sort On ProfID
'            .SortFields.Add key:=ws.Range("BJ2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On Date
'            .SortFields.Add key:=ws.Range("BL2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("BM2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange wshTEC_Local.Range("BJ2:BQ" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
'    Set rngCriteria = Nothing
'    Set rngData = Nothing
'    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modStatsHeures:Stats_Heures_AF", "", startTime)

End Sub

Sub ExecuterAdvancedFilter(ws As Worksheet, criteriaRange As String, resultStartCell As String, logRange As String, sortKeys As Variant, sortRange As String)

    Dim rngData As Range, rngCriteria As Range, rngResult As Range
    Dim lastResultRow As Long
    
    'Journaliser le temps de traitement
    ws.Range(logRange).ClearContents
    ws.Range(logRange).Cells(1, 1).Value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'D�finir le range source des donn�es
    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
    ws.Range(logRange).Cells(2, 1).Value = rngData.Address
    
    'D�finir les crit�res
    Set rngCriteria = ws.Range(criteriaRange)
    ws.Range(logRange).Cells(3, 1).Value = rngCriteria.Address
    
    'Effacer les r�sultats pr�c�dents
    Set rngResult = ws.Range(resultStartCell).CurrentRegion
    If rngResult.Rows.count > 1 Then
        rngResult.offset(1, 0).Clear
    End If
    Set rngResult = ws.Range(resultStartCell).CurrentRegion
    ws.Range(logRange).Cells(4, 1).Value = rngResult.Address
    
    'Appliquer AdvancedFilter
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
    
    'Compter les lignes
    lastResultRow = ws.Cells(ws.Rows.count, rngResult.Cells(1, 1).Column).End(xlUp).row
    ws.Range(logRange).Cells(5, 1).Value = lastResultRow - 1 & " lignes"
    
    'Trier les r�sultats
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

'Sub Stats_Heures_AF_OK()
'
'    'La cellule 'S7' doit contenir le Professionnel
'
'    Dim startTime as Double: startTime = Timer: Call Log_Record("modStatsHeures:Stats_Heures_AF", "", 0)
'
'    'On utilise la feuille TEC_TDB_Data
'    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
'
'    'Les m�me objects seront utilis�s avec les 4 AdvancedFilters
'    Dim rngData As Range
'    Dim rngCriteria As Range
'    Dim rngResult As Range
'    Dim lastResultRow As Long
'
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'
'    'AdvancedFilter # 1 - Semaine
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("T10:T14").ClearContents
'    ws.Range("T10").Value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
'    ws.Range("T11").Value = rngData.Address
'
'    'D�finir le range des crit�res
'    Set rngCriteria = ws.Range("S2:U3")
'    ws.Range("T12").Value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Set rngResult = ws.Range("W1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("W1").CurrentRegion
'    ws.Range("T13").Value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Tri des informations
'    lastResultRow = ws.Cells(ws.Rows.count, "W").End(xlUp).row
'    ws.Range("T14").Value = lastResultRow - 1 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 2 Then
'        With ws.Sort 'Sort - ID, Date, TecID
'            .SortFields.Clear
'            'First sort On ProfID
'            .SortFields.Add key:=ws.Range("W2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On Date
'            .SortFields.Add key:=ws.Range("Y2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("Z2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange wshTEC_Local.Range("W2:AD" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
'
'    'AdvancedFilter # 2 - Mois
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("AG10:AG14").ClearContents
'    ws.Range("AG10").Value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
'    ws.Range("AG11").Value = rngData.Address
'
'    'D�finir le range des crit�res
'    Set rngCriteria = ws.Range("AF2:AH3")
'    ws.Range("AG12").Value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Set rngResult = ws.Range("AJ1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("AJ1").CurrentRegion
'    ws.Range("AG13").Value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Tri des informations
'    lastResultRow = ws.Cells(ws.Rows.count, "AJ").End(xlUp).row
'    ws.Range("AG14").Value = lastResultRow - 1 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 2 Then
'        With ws.Sort 'Sort - ID, Date, TecID
'            .SortFields.Clear
'            'First sort On ProfID
'            .SortFields.Add key:=ws.Range("AJ2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On Date
'            .SortFields.Add key:=ws.Range("AL2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("AM2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange wshTEC_Local.Range("AJ2:AQ" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
'
'    'AdvancedFilter # 3 - Trimestre
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("AT10:AT16").ClearContents
'    ws.Range("AT10").Value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
'    ws.Range("AT11").Value = rngData.Address
'
'    'D�finir le range des crit�res
'    Set rngCriteria = ws.Range("AS2:AU3")
'    ws.Range("AT12").Value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Set rngResult = ws.Range("AW1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("AW1").CurrentRegion
'    ws.Range("AT13").Value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Tri des informations
'    lastResultRow = ws.Cells(ws.Rows.count, "AW").End(xlUp).row
'    ws.Range("AT14").Value = lastResultRow - 1 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 2 Then
'        With ws.Sort 'Sort - ID, Date, TecID
'            .SortFields.Clear
'            'First sort On ProfID
'            .SortFields.Add key:=ws.Range("AW2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On Date
'            .SortFields.Add key:=ws.Range("AY2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("AZ2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange wshTEC_Local.Range("AW2:BD" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
'
'    'AdvancedFilter # 4 - Ann�e financi�re
'
'    'Effacer les donn�es de la derni�re utilisation
'    ws.Range("BG10:BG14").ClearContents
'    ws.Range("BG10").Value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'D�finir le range pour la source des donn�es en utilisant un tableau
'    Set rngData = ws.Range("tblTEC_TDB_data[#All]")
'    ws.Range("BG11").Value = rngData.Address
'
'    'D�finir le range des crit�res
'    Set rngCriteria = ws.Range("BF2:BH3")
'    ws.Range("BG12").Value = rngCriteria.Address
'
'    'D�finir le range des r�sultats et effacer avant le traitement
'    Set rngResult = ws.Range("BJ1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("BJ1").CurrentRegion
'    ws.Range("BG13").Value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Tri des informations
'    lastResultRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).row
'    ws.Range("BG14").Value = lastResultRow - 1 & " lignes"
'
'    'Est-il n�cessaire de trier les r�sultats ?
'    If lastResultRow > 2 Then
'        With ws.Sort 'Sort - ID, Date, TecID
'            .SortFields.Clear
'            'First sort On ProfID
'            .SortFields.Add key:=ws.Range("BJ2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Second, sort On Date
'            .SortFields.Add key:=ws.Range("BL2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            'Third, sort On TecID
'            .SortFields.Add key:=ws.Range("BM2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal
'            .SetRange wshTEC_Local.Range("BJ2:BQ" & lastResultRow)
'            .Apply 'Apply Sort
'         End With
'    End If
'
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'
'    'Lib�rer la m�moire
'    Set rngCriteria = Nothing
'    Set rngData = Nothing
'    Set rngResult = Nothing
'    Set ws = Nothing
'
'    Call Log_Record("modStatsHeures:Stats_Heures_AF", "", startTime)
'
'End Sub
'
Sub shp_Back_To_ufSaisieHeures_Click()

    Call Back_To_ufSaisieHeures
    
End Sub

Sub Back_To_ufSaisieHeures()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modStatsHeures:Back_To_ufSaisieHeures", "", 0)
   
    wshStatsHeuresPivotTables.Visible = xlSheetHidden
    
    ufSaisieHeures.show vbModeless

    Call Log_Record("modStatsHeures:Back_To_ufSaisieHeures", "", startTime)

End Sub


