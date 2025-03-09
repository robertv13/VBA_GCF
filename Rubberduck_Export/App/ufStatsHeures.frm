VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufStatsHeures 
   Caption         =   "Statistiques d'heures"
   ClientHeight    =   8250.001
   ClientLeft      =   180
   ClientTop       =   690
   ClientWidth     =   15075
   OleObjectBlob   =   "ufStatsHeures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufStatsHeures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:UserForm_Initialize", "", 0)

    Call ChargerListBoxAvec52DernieresSemaines
    
    Call AddColonnesSemaine
    Call AddColonnesMois
    Call AddColonnesTrimestre
    Call AddColonnesAnneeFinanciere
    
    Call Log_Record("ufStatsHeures:UserForm_Initialize", "", startTime)
    
End Sub

Private Sub lbxDatesSemaines_Click() '2024-12-04 @ 07:36

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:lbxDatesSemaines_Click", lbxDatesSemaines.value, 0)
    
    Call lbxDatesSemaines_Click_or_DblClick(lbxDatesSemaines.value)
    
    Call Log_Record("ufStatsHeures:lbxDatesSemaines_Click", "", startTime)

End Sub

Private Sub lbxDatesSemaines_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:lbxDatesSemaines_DblClick", lbxDatesSemaines.value, 0)
    
    Call lbxDatesSemaines_Click_or_DblClick(lbxDatesSemaines.value)

    Call Log_Record("ufStatsHeures:lbxDatesSemaines_DblClick", "", startTime)

End Sub

Private Sub lbxDatesSemaines_Click_or_DblClick(ByVal valeur As Variant) '2024-12-04 @ 07:36
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:lbxDatesSemaines_Click_or_DblClick", lbxDatesSemaines.value, 0)
    
    Dim selectedWeek As String
    
    'Vérifier qu'un élément est bien sélectionné
    If lbxDatesSemaines.ListIndex <> -1 Then
        'Récupérer la semaine sélectionnée
        selectedWeek = lbxDatesSemaines.List(lbxDatesSemaines.ListIndex)
        Dim dateLundi As Date, dateDimanche As Date
        dateLundi = Left$(selectedWeek, InStr(selectedWeek, " au ") - 1)
        dateDimanche = Right$(selectedWeek, InStr(selectedWeek, " au ") - 1)
        
        'Il doit y avoir un écart de 6 entre les deux dates (semaine)
        If dateDimanche - dateLundi <> 6 Then
            MsgBox "Il semble y avoir un problème de format de date", vbCritical, _
                   "Dates de semaine NON VALIDES (" & dateLundi & " au " & dateDimanche & ")"
        End If
        
        'Initialisation du listBox et des totaux
        ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.RowSource = ""
        ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.Clear
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNettes.value = Format$(0, "##0.00") 'Formatage du total en deux décimales
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresFact.value = Format$(0, "##0.00") 'Formatage du total en deux décimales
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNF.value = Format$(0, "##0.00") 'Formatage du total en deux décimales

        'Envoie les deux dates à wshTEC_TDB_Data pour les AdvancedFilters
        Dim rngCriteriaDate1 As Range
        Dim formule1 As String
        Set rngCriteriaDate1 = wshTEC_TDB_Data.Range("T7")
        formule1 = rngCriteriaDate1.formula
        'Pour le premier changement de date, on ne veut pas passer par WorkSheet_Change
        Application.EnableEvents = False
        rngCriteriaDate1.value = dateValue(dateLundi)
        Application.EnableEvents = True
        
        Dim rngCriteriaDate2 As Range
        Dim formule2 As String
        Set rngCriteriaDate2 = wshTEC_TDB_Data.Range("U7")
        formule2 = rngCriteriaDate2.formula
        rngCriteriaDate2.value = dateValue(dateDimanche)
        
        If wshTEC_TDB_Data.Range("W2").value <> "" Then
            'Force une mise à jour du listBox en changeant le RowSource
            ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.RowSource = ""
            Dim lastUsedRow As Long
            lastUsedRow = wshTEC_TDB_Data.Cells(wshTEC_TDB_Data.Rows.count, "W").End(xlUp).row
            ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.RowSource = wshTEC_TDB_Data.Name & "!" & "StatsHeuresSemaine_uf"
'            ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.RowSource = wshTEC_TDB_Data.Range("W2:AD" & lastUsedRow).Address(external:=True)
'            Debug.Print wshTEC_TDB_Data.Name & "!" & "StatsHeuresSemaine_uf"
            
            DoEvents
        Else
            MsgBox "Il n'y a aucune heure d'enregistrée pour cette semaine", vbInformation
        End If
        
        Call AddColonnesSemaine
       
        'Rétablir les formules d'origine
        Application.EnableEvents = False
        rngCriteriaDate1.formula = "=DateDebutSemaine"
        rngCriteriaDate2.formula = "=DateFinSemaine"
        Application.EnableEvents = True
    Else
        MsgBox "Aucun élément sélectionné."
    End If
    
    'Libérer la mémoire
    Set rngCriteriaDate1 = Nothing
    Set rngCriteriaDate2 = Nothing
    
    Call Log_Record("ufStatsHeures:lbxDatesSemaines_Click_or_DblClick", "", startTime)

End Sub

Sub AddColonnesSemaine()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:AddColonnesSemaine", "", 0)
    
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    
    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim lastUsedResult As Long
    lastUsedResult = ws.Cells(ws.Rows.count, "W").End(xlUp).row
    Dim rngResult As Range
    Set rngResult = ws.Range("W2:AD" & lastUsedResult)
    
    t1 = Application.WorksheetFunction.Sum(rngResult.Columns(6))
    t2 = Application.WorksheetFunction.Sum(rngResult.Columns(7))
    t3 = Application.WorksheetFunction.Sum(rngResult.Columns(8))
    
    ufStatsHeures.lblTotaux = "* Totaux de la semaine (" & _
        Format$(wshTEC_TDB_Data.Range("T7").value, wshAdmin.Range("B1").value) & " au " & _
        Format$(wshTEC_TDB_Data.Range("U7").value, wshAdmin.Range("B1").value) & ") *"
    
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNettes.value = Format$(t1, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresFact.value = Format$(t2, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNF.value = Format$(t3, "#,##0.00") 'Formatage du total en deux décimales

    'Libérer la mémoire
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("ufStatsHeures:AddColonnesSemaine", "", startTime)

End Sub

Sub AddColonnesMois()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:AddColonnesMois", "", 0)
    
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    
    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim lastUsedResult As Long
    lastUsedResult = ws.Cells(ws.Rows.count, "AJ").End(xlUp).row
    Dim rngResult As Range
    Set rngResult = ws.Range("AJ2:AQ" & lastUsedResult)
    
    t1 = Application.WorksheetFunction.Sum(rngResult.Columns(6))
    t2 = Application.WorksheetFunction.Sum(rngResult.Columns(7))
    t3 = Application.WorksheetFunction.Sum(rngResult.Columns(8))
    
    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresNettes.value = Format$(t1, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresFact.value = Format$(t2, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresNF.value = Format$(t3, "#,##0.00") 'Formatage du total en deux décimales

    'Libérer la mémoire
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("ufStatsHeures:AddColonnesMois", "", startTime)

End Sub

Sub AddColonnesTrimestre()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:AddColonnesTrimestre", "", 0)
    
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    
    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim lastUsedResult As Long
    lastUsedResult = ws.Cells(ws.Rows.count, "AW").End(xlUp).row
    Dim rngResult As Range
    Set rngResult = ws.Range("AW2:BD" & lastUsedResult)
    
    t1 = Application.WorksheetFunction.Sum(rngResult.Columns(6))
    t2 = Application.WorksheetFunction.Sum(rngResult.Columns(7))
    t3 = Application.WorksheetFunction.Sum(rngResult.Columns(8))
    
    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresNettes.value = Format$(t1, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresFact.value = Format$(t2, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresNF.value = Format$(t3, "#,##0.00") 'Formatage du total en deux décimales

    'Libérer la mémoire
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("ufStatsHeures:AddColonnesTrimestre", "", startTime)

End Sub

Sub AddColonnesAnneeFinanciere()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:AddColonnesAnneeFinanciere", "", 0)
    
    Dim ws As Worksheet: Set ws = wshTEC_TDB_Data
    
    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim lastUsedResult As Long
    lastUsedResult = ws.Cells(ws.Rows.count, "BJ").End(xlUp).row
    Dim rngResult As Range
    Set rngResult = ws.Range("BJ2:BQ" & lastUsedResult)
    
    t1 = Application.WorksheetFunction.Sum(rngResult.Columns(6))
    t2 = Application.WorksheetFunction.Sum(rngResult.Columns(7))
    t3 = Application.WorksheetFunction.Sum(rngResult.Columns(8))
    
    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresNettes.value = Format$(t1, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresFact.value = Format$(t2, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresNF.value = Format$(t3, "#,##0.00") 'Formatage du total en deux décimales

    'Libérer la mémoire
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("ufStatsHeures:AddColonnesAnneeFinanciere", "", startTime)

End Sub

Sub ChargerListBoxAvec52DernieresSemaines()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:ChargerListBoxAvec52DernieresSemaines", "", 0)
    
    Dim i As Integer
    Dim dtLundi As Date
    Dim dtDimanche As Date
    
    'Référence à la ListBox
    Dim lstSemaines As Control
    Set lstSemaines = ufStatsHeures.MultiPage1.Pages("pSemaine").Controls("lbxDatesSemaines")
    
    'Nettoyer la ListBox avant d'ajouter des éléments
    lstSemaines.Clear
    
    'Calculer la date du lundi de la semaine actuelle
    dtLundi = Date - Weekday(Date, vbMonday) + 1
    
    'Boucle pour charger les 52 dernières semaines
    Dim semaines(1 To 53) As String
    For i = 53 To 1 Step -1
        'Calculer la date du dimanche de la semaine correspondante
        dtDimanche = dtLundi + 6
        
        'Ajouter l'intervalle dans la ListBox
        semaines(i) = Format$(CLng(dtLundi), wshAdmin.Range("B1").value) & " au " & Format$(CLng(dtDimanche), wshAdmin.Range("B1").value)
        
        'Passer à la semaine précédente
        dtLundi = dtLundi - 7
    Next i
    
    'Charger les éléments dans la ListBox (les plus anciens en premier)
    For i = 1 To 53
        lstSemaines.AddItem semaines(i)
    Next i
    
    'On se positionne à la fin de la liste (évite de monter/descendre)
    lstSemaines.TopIndex = lstSemaines.ListCount - 1
    lstSemaines.ListIndex = lstSemaines.ListCount - 1 '2024-12-04 @ 07:49

    'Libérer la mémoire
'    Set lstSemaines = Nothing
    
    Call Log_Record("ufStatsHeures:ChargerListBoxAvec52DernieresSemaines", "", startTime)
    
End Sub

