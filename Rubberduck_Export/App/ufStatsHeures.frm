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

Private Sub cmdAide_Click()

End Sub

Private Sub lbxDatesSemaines_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim selectedWeek As String
    
    'V�rifier qu'un �l�ment est bien s�lectionn�
    If lbxDatesSemaines.ListIndex <> -1 Then
        'R�cup�rer l'�l�ment s�lectionn�
        selectedWeek = lbxDatesSemaines.List(lbxDatesSemaines.ListIndex)
        Dim dateLundi As Date, dateDimanche As Date
        dateLundi = Left(selectedWeek, 10)
        dateDimanche = Right(selectedWeek, 10)
        
        'Initialisation du listBox et des totaux
        ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.RowSource = ""
        ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.Clear
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNettes.value = Format(0, "##0.00") 'Formatage du total en deux d�cimales
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresFact.value = Format(0, "##0.00") 'Formatage du total en deux d�cimales
        ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNF.value = Format(0, "##0.00") 'Formatage du total en deux d�cimales

        'Envoie les deux dates � wshTEC_TDB_Data pour les AdvancedFilters
        Dim criteriaDate1 As Range
        Dim formule1 As String
        Set criteriaDate1 = wshTEC_TDB_Data.Range("T7")
        formule1 = criteriaDate1.formula
        criteriaDate1 = dateValue(dateLundi)
        
        Dim criteriaDate2 As Range
        Dim formule2 As String
        Set criteriaDate2 = wshTEC_TDB_Data.Range("U7")
        formule2 = criteriaDate2.formula
        criteriaDate2 = dateValue(dateDimanche)
        
        If wshTEC_TDB_Data.Range("W2").value <> "" Then
            'Force une mise � jour du listBox en changeant le RowSource
            ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.RowSource = "StatsHeuresSemaine_uf"
            DoEvents
            ufStatsHeures.lblTotaux = "Totaux de la semaine (" & dateLundi & " au " & dateDimanche & ")"
            Call AddColonnesSemaine
        End If
        
        'R�tablir les formules d'origine
        Application.EnableEvents = False
        criteriaDate1.formula = "=DateDebutSemaine"
        criteriaDate2.formula = "=DateFinSemaine"
        Application.EnableEvents = True
    Else
        MsgBox "Aucun �l�ment s�lectionn�."
    End If
    
    'Lib�rer la m�moire
    Set criteriaDate1 = Nothing
    Set criteriaDate2 = Nothing
    
End Sub

Private Sub UserForm_Initialize()

    Call AddColonnesSemaine
    Call AddColonnesMois
    Call AddColonnesTrimestre
    Call AddColonnesAnneeFinanciere
    
End Sub

Sub AddColonnesSemaine()

    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.ListCount - 1
        t1 = t1 + CCur(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 4))
        t2 = t2 + CCur(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 5))
        t3 = t3 + CCur(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNettes.value = Format(t1, "#,##0.00") 'Formatage du total en deux d�cimales
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresFact.value = Format(t2, "#,##0.00") 'Formatage du total en deux d�cimales
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNF.value = Format(t3, "#,##0.00") 'Formatage du total en deux d�cimales

    Call ChargerListBoxAvec52DernieresSemaines
    
End Sub

Sub AddColonnesMois()

    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.ListCount - 1
        t1 = t1 + CCur(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 4))
        t2 = t2 + CCur(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 5))
        t3 = t3 + CCur(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresNettes.value = Format(t1, "#,##0.00") 'Formatage du total en deux d�cimales
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresFact.value = Format(t2, "#,##0.00") 'Formatage du total en deux d�cimales
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresNF.value = Format(t3, "#,##0.00") 'Formatage du total en deux d�cimales

End Sub

Sub AddColonnesTrimestre()

    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.ListCount - 1
        t1 = t1 + CCur(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 4))
        t2 = t2 + CCur(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 5))
        t3 = t3 + CCur(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresNettes.value = Format(t1, "#,##0.00") 'Formatage du total en deux d�cimales
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresFact.value = Format(t2, "#,##0.00") 'Formatage du total en deux d�cimales
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresNF.value = Format(t3, "#,##0.00") 'Formatage du total en deux d�cimales

End Sub

Sub AddColonnesAnneeFinanciere()

    Dim t1 As Currency, t2 As Currency, t3 As Currency
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.ListCount - 1
        t1 = t1 + CCur(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 4))
        t2 = t2 + CCur(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 5))
        t3 = t3 + CCur(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresNettes.value = Format(t1, "#,##0.00") 'Formatage du total en deux d�cimales
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresFact.value = Format(t2, "#,##0.00") 'Formatage du total en deux d�cimales
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresNF.value = Format(t3, "#,##0.00") 'Formatage du total en deux d�cimales

End Sub

Sub ChargerListBoxAvec52DernieresSemaines()
    
    Dim i As Integer
    Dim dtLundi As Date
    Dim dtDimanche As Date
    
    'R�f�rence � la ListBox
    Dim lstSemaines As Control
    Set lstSemaines = ufStatsHeures.MultiPage1.Pages("pSemaine").Controls("lbxDatesSemaines")
    
    'Nettoyer la ListBox avant d'ajouter des �l�ments
    lstSemaines.Clear
    
    'Calculer la date du lundi de la semaine actuelle
    dtLundi = Date - Weekday(Date, vbMonday) + 1
    
    'Boucle pour charger les 52 derni�res semaines
    Dim semaines(1 To 53) As String
    For i = 53 To 1 Step -1
        'Calculer la date du dimanche de la semaine correspondante
        dtDimanche = dtLundi + 6
        
        'Ajouter l'intervalle dans la ListBox
        semaines(i) = Format(dtLundi, "dd/mm/yyyy") & " au " & Format(dtDimanche, "dd/mm/yyyy")
        
        'Passer � la semaine pr�c�dente
        dtLundi = dtLundi - 7
    Next i
    
    'Charger les �l�ments dans la ListBox (les plus anciens en premier)
    For i = 1 To 53
        lstSemaines.AddItem semaines(i)
    Next i
    
    'On se positionne � la fin de la liste (�vite de monter/descendre)
    lstSemaines.TopIndex = lstSemaines.ListCount - 1

    'Lib�rer la m�moire
    Set lstSemaines = Nothing
    
End Sub

