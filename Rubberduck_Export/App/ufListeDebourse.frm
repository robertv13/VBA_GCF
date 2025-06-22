VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeDebourse 
   Caption         =   "Liste des déboursés"
   ClientHeight    =   5580
   ClientLeft      =   96
   ClientTop       =   384
   ClientWidth     =   19008
   OleObjectBlob   =   "ufListeDebourse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListeDebourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dataArray() As Variant     'Tous les renregistrements DEB_Trans
Private recentArray() As Variant   'Enregistrements récents (< 75 jours)
Private filteredArray() As Variant 'Enregistrements filtrés (si filtre)
'

Private Sub UserForm_Initialize()
    
    Call ChargerDebDonnees
    
End Sub

Private Sub ChargerDebDonnees()

    'Définir la feuille source et la plage des données
    Dim ws As Worksheet
    Set ws = wsdDEB_Trans
    dataArray = ws.Range("A2:S" & ws.Cells(ws.Rows.count, 1).End(xlUp).row).Value
    
    'Définir la date limite (75 jours avant aujourd'hui)
    Dim dateLimite As Date
    dateLimite = Date - 75
    
    'Définir les indices des colonnes à afficher
    Dim columnsToShow As Variant
    columnsToShow = Array(fDebTDate, fDebTBeneficiaire, fDebTDescription, fDebTCodeTaxe, fDebTTotal, _
                            fDebTCréditTPS, fDebTCréditTVQ, fDebTDépense, fDebTCompte, fDebTType, 0)
    
    'Déterminer le nombre de colonnes à afficher
    Dim nbColonnesAffichees As Long
    nbColonnesAffichees = UBound(columnsToShow) - LBound(columnsToShow) + 1

    'Déterminer le nombre de lignes requises dans recentArray
    On Error GoTo 0
    Dim rowCount As Long
    Dim i As Long
    For i = 1 To UBound(dataArray, 1)
        If dataArray(i, fDebTDate) >= dateLimite And _
            InStr(dataArray(i, fDebTDescription), " (RENVERSÉ par ") = 0 And _
            InStr(dataArray(i, fDebTDescription), " (RENVERSEMENT de ") = 0 Then
            rowCount = rowCount + 1
        End If
    Next i
    
    'Tableau des données filtrées
    If rowCount > 0 Then
        ReDim recentArray(1 To rowCount, 1 To nbColonnesAffichees)
        'Définir la largeur des colonnes du listBox
        Me.lsbListeDebourse.ColumnWidths = "60;160;190;55;72;72;72;72;190;160;20"

        'Filtrer les enregistrements de moins de 75 jours
        Dim j As Long
        rowCount = 0
            For i = LBound(dataArray, 1) To UBound(dataArray, 1)
            If IsDate(dataArray(i, 2)) Then
                If dataArray(i, fDebTDate) >= dateLimite And _
                InStr(dataArray(i, fDebTDescription), " (RENVERSÉ par ") = 0 And _
                InStr(dataArray(i, fDebTDescription), " (RENVERSEMENT de ") = 0 Then
                    rowCount = rowCount + 1
                    'Copier uniquement les colonnes sélectionnées
                    For j = LBound(columnsToShow) To UBound(columnsToShow)
                        If j < 10 Then
                            recentArray(rowCount, j + 1) = dataArray(i, columnsToShow(j))
                        Else
                            'Emmagasine le numéro de déboursé pour retrouver les informations
                            recentArray(rowCount, j + 1) = CLng(dataArray(i, 1))
                        End If
                    Next j
                End If
            End If
        Next i
        
        'Charger dans le ListBox après avoir effectué un tri sur la date et formater les colonnes
        Call Array_2D_Bubble_Sort(recentArray)
        
        Call FormatArrayBeforeAddingToDebListBox(recentArray)
        
        Me.lsbListeDebourse.List = recentArray
        
        'Positionne à la dernière entrée
        If Me.lsbListeDebourse.ListCount > 0 Then
            Me.lsbListeDebourse.ListIndex = Me.lsbListeDebourse.ListCount - 1
        End If
    Else
        Me.lsbListeDebourse.Clear
    End If
    
    numeroDebourseARenverser = -1
    wshDEB_Saisie.Range("B7").Value = False
    
End Sub

Private Sub txtFiltre_Change()

    'Filtrer les données à chaque changement dans le TextBox
    Call UpdateFilteredArray(Me.txtFiltre.Text)
    
End Sub

Private Sub UpdateFilteredArray(filtre As String)

    'Récupérer le texte du TextBox pour filtrer
    Dim filterText As String
    filterText = Me.txtFiltre.Text

    'Initialiser rowCount pour les résultats filtrés
    Dim rowCount As Long
    rowCount = 0
    
    'Exploration pour voir si on a des enregistrements en fonction du filtre
    Dim i As Long
    For i = 1 To UBound(recentArray, 1)
        If InStr(1, recentArray(i, 2), filterText, vbTextCompare) > 0 Or _
            InStr(1, recentArray(i, 3), filterText, vbTextCompare) > 0 Or _
            InStr(1, recentArray(i, 9), filterText, vbTextCompare) > 0 Or _
            InStr(1, recentArray(i, 10), filterText, vbTextCompare) > 0 Then
            rowCount = rowCount + 1
        End If
    Next i
'    Debug.Print "'" & filterText & "' - " & rowCount

    'Si des lignes valides sont trouvées, créer filteredArray
    Dim j As Long, k As Long
    If rowCount > 0 Then
        ReDim filteredArray(1 To rowCount, 1 To UBound(recentArray, 2))
        'Copier les données filtrées de recentArray vers filteredArray
        j = 0
        For i = 1 To UBound(recentArray, 1)
            If InStr(1, recentArray(i, 2), filterText, vbTextCompare) > 0 Or _
                InStr(1, recentArray(i, 3), filterText, vbTextCompare) > 0 Or _
                InStr(1, recentArray(i, 9), filterText, vbTextCompare) > 0 Or _
                InStr(1, recentArray(i, 10), filterText, vbTextCompare) > 0 Then
                    j = j + 1
                    'Copier les données dans filteredArray
                    For k = 1 To UBound(recentArray, 2)
                        filteredArray(j, k) = recentArray(i, k)
                    Next k
            End If
        Next i
        'Charger filteredArray dans le ListBox
'        Call FormatArrayBeforeAddingToDebListBox(filteredArray)
        Me.lsbListeDebourse.List = filteredArray
    Else
        Me.lsbListeDebourse.Clear  ' Si aucun enregistrement, vider la ListBox
    End If
    
    numeroDebourseARenverser = -1
    wshDEB_Saisie.Range("B7").Value = False
    
End Sub

Private Sub lsbListeDebourse_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim selectedRow As Long
    
    'Vérifier si une ligne a été sélectionnée
    If lsbListeDebourse.ListIndex <> -1 Then
        'Récupérer le numéro de déboursé à renverser
        selectedRow = lsbListeDebourse.ListIndex
        numeroDebourseARenverser = lsbListeDebourse.List(selectedRow, 10)
        wshDEB_Saisie.Range("B7").Value = True
    Else
        numeroDebourseARenverser = -1
        wshDEB_Saisie.Range("B7").Value = False
    End If
        
    Unload Me
    
End Sub

Sub FormatArrayBeforeAddingToDebListBox(ByRef arrData As Variant)
    
    'Supposons que la première colonne (1) est une date
    Dim i As Long, j As Long
    For i = 1 To UBound(arrData, 1)
        'Formater la première colonne comme date
        arrData(i, 1) = Format$(arrData(i, 1), wsdADMIN.Range("B1").Value)
        'Formater les colonnes contenant des montants, alignées à droite avec espaces
        For j = 5 To 8
            arrData(i, j) = Format$(arrData(i, j), "#,##0.00;-#,##0.00;-")
            arrData(i, j) = Space(11 - Len(arrData(i, j))) & arrData(i, j)
        Next j
    Next i
    
End Sub

Private Sub cmdFermer_Click()

    'Pas de rowNumber pour renverser
    numeroDebourseARenverser = -1
    wshDEB_Saisie.Range("B7").Value = False
    
    Unload Me
    
End Sub


