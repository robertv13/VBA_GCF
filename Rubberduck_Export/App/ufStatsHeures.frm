VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufStatsHeures 
   Caption         =   "Statistiques d'heures"
   ClientHeight    =   8250.001
   ClientLeft      =   120
   ClientTop       =   465
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

    Call AddColonnesSemaine
    Call AddColonnesMois
    Call AddColonnesTrimestre
    Call AddColonnesAnneeFinanciere
    
End Sub

Sub AddColonnesSemaine()

    Dim t1 As Double, t2 As Double, t3 As Double
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.ListCount - 1
        t1 = t1 + CDbl(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 4))
        t2 = t2 + CDbl(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 5))
        t3 = t3 + CDbl(ufStatsHeures.MultiPage1.Pages("pSemaine").lbxSemaine.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNettes.value = Format(t1, "##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresFact.value = Format(t2, "##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pSemaine").txtSemaineHresNF.value = Format(t3, "##0.00") 'Formatage du total en deux décimales

End Sub

Sub AddColonnesMois()

    Dim t1 As Double, t2 As Double, t3 As Double
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.ListCount - 1
        t1 = t1 + CDbl(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 4))
        t2 = t2 + CDbl(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 5))
        t3 = t3 + CDbl(ufStatsHeures.MultiPage1.Pages("pMois").lbxMois.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresNettes.value = Format(t1, "##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresFact.value = Format(t2, "##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pMois").txtMoisHresNF.value = Format(t3, "##0.00") 'Formatage du total en deux décimales

End Sub

Sub AddColonnesTrimestre()

    Dim t1 As Double, t2 As Double, t3 As Double
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.ListCount - 1
        t1 = t1 + CDbl(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 4))
        t2 = t2 + CDbl(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 5))
        t3 = t3 + CDbl(ufStatsHeures.MultiPage1.Pages("pTrimestre").lbxTrimestre.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresNettes.value = Format(t1, "##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresFact.value = Format(t2, "##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pTrimestre").txtTrimHresNF.value = Format(t3, "##0.00") 'Formatage du total en deux décimales

End Sub

Sub AddColonnesAnneeFinanciere()

    Dim t1 As Double, t2 As Double, t3 As Double
    
    Dim i As Long
    For i = 0 To ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.ListCount - 1
        t1 = t1 + CDbl(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 4))
        t2 = t2 + CDbl(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 5))
        t3 = t3 + CDbl(ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").lbxAnneeFinanciere.List(i, 6))
    Next i

    'Affiche le total dans la TextBox
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresNettes.value = Format(t1, "##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresFact.value = Format(t2, "##0.00") 'Formatage du total en deux décimales
    ufStatsHeures.MultiPage1.Pages("pAnneeFinanciere").txtAnneeFinanciereHresNF.value = Format(t3, "##0.00") 'Formatage du total en deux décimales

End Sub

