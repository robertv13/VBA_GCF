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

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:UserForm_Initialize", 0)

    Call ChargerListBoxAvec52DernieresSemaines
    
    Call AddColonnesSemaine
    Call AddColonnesMois
    Call AddColonnesTrimestre
    Call AddColonnesAnneeFinanciere
    
    Call Log_Record("ufStatsHeures:UserForm_Initialize", startTime)
    
End Sub

Private Sub lbxDatesSemaines_Click() '2024-12-04 @ 07:36

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:lbxDatesSemaines_Click(" & lbxDatesSemaines.value & ")", 0)
    
    Call lbxDatesSemaines_Click_or_DblClick(lbxDatesSemaines.value)
    
    Call Log_Record("ufStatsHeures:lbxDatesSemaines_Click", startTime)

End Sub

Private Sub lbxDatesSemaines_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:lbxDatesSemaines_DblClick(" & lbxDatesSemaines.value & ")", 0)
    
    Call lbxDatesSemaines_Click_or_DblClick(lbxDatesSemaines.value)

    Call Log_Record("ufStatsHeures:lbxDatesSemaines_DblClick", startTime)

End Sub

Private Sub lbxDatesSemaines_Click_or_DblClick(ByVal Valeur As Variant) '2024-12-04 @ 07:36
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:lbxDatesSemaines_Click_or_DblClick(" & lbxDatesSemaines.value & ")", 0)
    
    Dim selectedWeek As String
    
    'Vérifier qu'un élément est bien sélectionné
    If lbxDatesSemaines.ListIndex <> -1 Then
        'Récupérer l'élément sélectionné
        selectedWeek = lbxDatesSemaines.List(lbxDatesSemaines.ListIndex)
        Dim dateLundi As Date, dateDimanche As Date
        dateLundi = Left(selectedWeek, 10)
        dateDimanche = Right(selectedWeek, 10)
        
        'Initialisation du listBox et des totaux
        ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.RowSource = ""
        ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.Clear
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNettes.value = Format(0, "##0.00") 'Formatage du total en deux décimales
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresFact.value = Format(0, "##0.00") 'Formatage du total en deux décimales
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNF.value = Format(0, "##0.00") 'Formatage du total en deux décimales

        'Envoie les deux dates à wshTEC_TDB_Data pour les AdvancedFilters
        Dim criteriaDate1 As Range
        Dim formule1 As String
        Set criteriaDate1 = wshTEC_TDB_Data.Range("T7")
        formule1 = criteriaDate1.formula
        'Pour le premier changement de date, on ne veut pas passer par WorkSheet_Change
        Application.EnableEvents = False
        criteriaDate1 = dateValue(dateLundi)
        Application.EnableEvents = True
        
        Dim criteriaDate2 As Range
        Dim formule2 As String
        Set criteriaDate2 = wshTEC_TDB_Data.Range("U7")
        formule2 = criteriaDate2.formula
        criteriaDate2 = dateValue(dateDimanche)
        
        If wshTEC_TDB_Data.Range("W2").value <> "" Then
            'Force une mise à jour du listBox en changeant le RowSource
            ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.RowSource = "StatsHeuresSemaine_uf"
            DoEvents
        End If
        
        ufStatsHeures.lblTotaux = "Totaux de la semaine (" & _
                    Format$(dateLundi, wshAdmin.Range("B1").value) & " au " & _
                    Format$(dateDimanche, wshAdmin.Range("B1").value) & ")"
        Call AddColonnesSemaine
       
        'Rétablir les formules d'origine
        Application.EnableEvents = False
        criteriaDate1.formula = "=DateDebutSemaine"
        criteriaDate2.formula = "=DateFinSemaine"
        Application.EnableEvents = True
    Else
        MsgBox "Aucun élément sélectionné."
    End If
    
    'Libérer la mémoire
    Set criteriaDate1 = Nothing
    Set criteriaDate2 = Nothing
    
    Call Log_Record("ufStatsHeures:lbxDatesSemaines_Click_or_DblClick", startTime)

End Sub

Sub AddColonnesSemaine()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:AddColonnesSemaine", 0)
    
    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.ListCount - 1
        t1 = t1 + CCur(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 4))
        t2 = t2 + CCur(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 5))
        t3 = t3 + CCur(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 6))
    Next i

'    Dim selectedWeek As String
'    selectedWeek = ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.ListCount - 1
'    Dim dateLundi As Date, dateDimanche As Date
'    dateLundi = Left(selectedWeek, 10)
'    dateDimanche = Right(selectedWeek, 10)
'
'    ufStatsHeures.lblTotaux = "Totaux de la semaine (" & _
'        Format$(dateLundi, wshAdmin.Range("B1").value) & " au " & _
'        Format$(dateDimanche, wshAdmin.Range("B1").value) & ")"

    'Affiche le total dans la TextBox
'    ufStatsHeures.lblTotaux = "Totaux de la semaine (" & dernSemaine & ")"

    ufStatsHeures.lblTotaux = "Totaux de la semaine (" & _
        Format$(wshTEC_TDB_Data.Range("T7").value, wshAdmin.Range("B1").value) & " au " & _
        Format$(wshTEC_TDB_Data.Range("U7").value, wshAdmin.Range("B1").value) & ")"
    
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNettes.value = Format(t1, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresFact.value = Format(t2, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNF.value = Format(t3, "#,##0.00") 'Formatage du total en deux décimales

'    Call ChargerListBoxAvec52DernieresSemaines
    
    Call Log_Record("ufStatsHeures:AddColonnesSemaine", startTime)

End Sub

Sub AddColonnesMois()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:AddColonnesMois", 0)
    
    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.ListCount - 1
        t1 = t1 + CCur(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 4))
        t2 = t2 + CCur(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 5))
        t3 = t3 + CCur(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresNettes.value = Format(t1, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresFact.value = Format(t2, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresNF.value = Format(t3, "#,##0.00") 'Formatage du total en deux décimales

    Call Log_Record("ufStatsHeures:AddColonnesMois", startTime)

End Sub

Sub AddColonnesTrimestre()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:AddColonnesTrimestre", 0)
    
    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.ListCount - 1
        t1 = t1 + CCur(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 4))
        t2 = t2 + CCur(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 5))
        t3 = t3 + CCur(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresNettes.value = Format(t1, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresFact.value = Format(t2, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresNF.value = Format(t3, "#,##0.00") 'Formatage du total en deux décimales

    Call Log_Record("ufStatsHeures:AddColonnesTrimestre", startTime)

End Sub

Sub AddColonnesAnneeFinanciere()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:AddColonnesAnneeFinanciere", 0)
    
    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.ListCount - 1
        t1 = t1 + CCur(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 4))
        t2 = t2 + CCur(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 5))
        t3 = t3 + CCur(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresNettes.value = Format(t1, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresFact.value = Format(t2, "#,##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresNF.value = Format(t3, "#,##0.00") 'Formatage du total en deux décimales

    Call Log_Record("ufStatsHeures:AddColonnesAnneeFinanciere", startTime)

End Sub

Sub ChargerListBoxAvec52DernieresSemaines()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufStatsHeures:ChargerListBoxAvec52DernieresSemaines", 0)
    
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
        semaines(i) = Format$(dtLundi, wshAdmin.Range("B1").value) & " au " & Format$(dtDimanche, wshAdmin.Range("B1").value)
        
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
    
    Call Log_Record("ufStatsHeures:ChargerListBoxAvec52DernieresSemaines", startTime)
    
End Sub

