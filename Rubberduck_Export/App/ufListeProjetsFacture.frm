VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeProjetsFacture 
   Caption         =   "Facturation des projets de facture"
   ClientLeft      =   -15
   ClientTop       =   -30
   OleObjectBlob   =   "ufListeProjetsFacture.frx":0000
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

        If UCase$(estDetruite) <> "VRAI" And estDetruite <> -1 Then
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
    Call RedimensionnerTableau2D(arr, nbRows, 4)

    'Trier les données (si souhaité)
    Call TrierTableau2DBubble(arr)

    ' Préparer la ListBox
    With Me.lsbProjetsFacture
        .Clear
        .ColumnHeads = True
        .ColumnCount = 4
        .ColumnWidths = "350; 68; 85; 20"
        .List = arr
    End With

    'Nettoyage
    Set ligne = Nothing
    Set lo = Nothing
    Set ws = Nothing
    
End Sub

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

