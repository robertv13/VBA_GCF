VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeProjetsFacture 
   Caption         =   "Facturation des projets de facture"
   ClientHeight    =   7665
   ClientLeft      =   180
   ClientTop       =   690
   ClientWidth     =   11520
   OleObjectBlob   =   "ufListeProjetsFacture.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListeProjetsFacture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize() '2025-06-01 @ 06:54

    Dim ws As Worksheet
    Set ws = wsdFAC_Projets_Entete

    Dim lo As ListObject
    Set lo = ws.ListObjects("l_tbl_FAC_Projets_Entête")

    'Vérifier qu’il y a des données réelles dans le tableau (ignore les lignes vides)
    If Not TableauContientDesDonnees(lo) Then
        Unload Me
        Exit Sub
    End If

    Dim arr() As Variant
    Dim i As Long, nbRows As Long
    Dim ligne As Range
    Dim estDetruite As Variant

    ReDim arr(1 To lo.ListRows.count, 1 To 4)

    For i = 1 To lo.ListRows.count
        Set ligne = lo.ListRows(i).Range
        'Ignorer si toute la ligne est vide (ligne fantôme)
        If Application.WorksheetFunction.CountA(ligne) = 0 Then GoTo ProchaineLigne

        estDetruite = ligne.Columns(lo.ListColumns("estDetruite").index).Value

        If UCase$(estDetruite) <> "VRAI" Then
            nbRows = nbRows + 1
            arr(nbRows, 1) = ligne.Columns(lo.ListColumns("nomClient").index).Value
            arr(nbRows, 2) = ligne.Columns(lo.ListColumns("date").index).Value
            arr(nbRows, 3) = Fn_Pad_A_String(Format$(ligne.Columns(lo.ListColumns("HonoTotal").index).Value, "#,##0.00$"), " ", 11, "L")
            arr(nbRows, 4) = ligne.Columns(lo.ListColumns("ProjetID").index).Value
        End If

ProchaineLigne:
    Next i
    
    If nbRows = 0 Then
        Unload Me
        Exit Sub
    End If

    'Redimensionner proprement
    Call Array_2D_Resizer(arr, nbRows, 4)

    'Trier les données (si souhaité)
    Call Array_2D_Bubble_Sort(arr)

    ' Préparer la ListBox
    With Me.lsbProjetsFacture
        .Clear
        .ColumnHeads = True
        .ColumnCount = 4
        .ColumnWidths = "350; 68; 85; 20"
        .List = arr
    End With

'    ' Charger les données
'    Dim j As Long
'    For i = LBound(arr, 1) To UBound(arr, 1)
'        Me.lsbProjetsFacture.AddItem
'        For j = LBound(arr, 2) To UBound(arr, 2)
'            Me.lsbProjetsFacture.List(i - 1, j - 1) = arr(i, j)
'        Next j
'    Next i
'
    ' Nettoyage
    Set ligne = Nothing
    Set lo = Nothing
    Set ws = Nothing
    
End Sub

'CommentOut - 2025-06-01 @ 06:44
'Private Sub UserForm_Initialize()
'
'    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Entete
'
'    Dim lastUsedRow As Long
'    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
'    If lastUsedRow <= 2 Then Exit Sub 'Empty List
'
'    Dim arr() As Variant
'    ReDim arr(1 To lastUsedRow - 1, 1 To 4)
'
'    'Populate the array with non-contiguous columns
'    Dim i As Long, nbRows As Long
'    For i = 2 To lastUsedRow
'        If UCase$(ws.Cells(i, 26).value) <> "VRAI" Then 'Exclude those projects with isDetruite set to True
'            nbRows = nbRows + 1
'            arr(nbRows, 1) = ws.Cells(i, 2).value 'nomClient
'            arr(nbRows, 2) = ws.Cells(i, 4).value 'date
'            arr(nbRows, 3) = Fn_Pad_A_String(Format$(ws.Cells(i, 5).value, "#,##0.00$"), " ", 11, "L") 'Honoraires
'            arr(nbRows, 4) = ws.Cells(i, 1).value 'ProjetID
'        End If
'    Next i
'
'    If nbRows > 0 Then
'
'        Call Array_2D_Resizer(arr, nbRows, UBound(arr, 2))
'
'        'Sort the list
'        Call Array_2D_Bubble_Sort(arr)
'
'        'Clear the ListBox
'        Me.lsbProjetsFacture.Clear
'
'        With lsbProjetsFacture
'            .ColumnHeads = True
'            .ColumnCount = 4
'            .ColumnWidths = "350; 68; 85; 20"
'        End With
'
'        'Populate the ListBox with the array
'        Dim j As Long
'        For i = LBound(arr, 1) To UBound(arr, 1)
'            Me.lsbProjetsFacture.AddItem
'            For j = LBound(arr, 2) To UBound(arr, 2)
'                Me.lsbProjetsFacture.List(i - 1, j - 1) = arr(i, j)
'            Next j
'        Next i
'    Else
'        Unload Me
'    End If
'
'    'Libérer la mémoire
'    Set ws = Nothing
'    Exit Sub
'
'End Sub

Private Sub lsbProjetsFacture_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '2024-07-21 @ 16:38

    Dim rowSelected As Long
    Dim nomClient As String, dte As Date
    Dim honorairesTotal As Double
    Dim projetID As Long
    
    rowSelected = lsbProjetsFacture.ListIndex
    nomClient = lsbProjetsFacture.List(rowSelected, 0)
    dte = CDate(lsbProjetsFacture.List(rowSelected, 1))
    honorairesTotal = lsbProjetsFacture.List(rowSelected, 2)
    projetID = lsbProjetsFacture.List(rowSelected, 3)
    
    Application.EnableEvents = False
    
    wshFAC_Brouillon.Range("B51").Value = nomClient
    wshFAC_Brouillon.Range("B52").Value = projetID
    wshFAC_Brouillon.Range("B53").Value = dte
    wshFAC_Brouillon.Range("B54").Value = honorairesTotal
    
    Application.EnableEvents = True
    
    Unload ufListeProjetsFacture

End Sub

Private Sub UserForm_Terminate()
    
    Unload ufListeProjetsFacture
    
End Sub

