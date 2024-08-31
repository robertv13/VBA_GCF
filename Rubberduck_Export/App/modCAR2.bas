Attribute VB_Name = "modCAR2"
Option Explicit

Sub Generer_Liste_Agée_CAR()

    ' Initialiser les feuilles
    Dim wsFactures As Worksheet
    Set wsFactures = ThisWorkbook.Sheets("FAC_Comptes_Clients")
    Dim wsPaiements As Worksheet
    Set wsPaiements = ThisWorkbook.Sheets("ENC_Détails")
    
    'Créer une nouvelle feuille pour les résultats
    On Error Resume Next
    Dim wsResultat As Worksheet
    Application.DisplayAlerts = False
    Set wsResultat = ThisWorkbook.Sheets("ListeAgee")
    If Not wsResultat Is Nothing Then wsResultat.delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsResultat = ThisWorkbook.Sheets.add
    wsResultat.name = "ListeAgee"
    Dim rngResultat As Range
    Set rngResultat = wsResultat.Range("A1")
    
    'Demander à l'utilisateur le niveau de détail
    Dim niveauDetail As String
    niveauDetail = InputBox("Choisissez le niveau de détail : Client, Facture, Transaction")
    
    'Entêtes de colonnes en fonction du niveau de détail
    If LCase(niveauDetail) = "client" Then
        wsResultat.Range("A1:F1").value = Array("Client", "Total", "0-30 jours", "31-60 jours", "61-90 jours", "90+ jours")
    End If

    Call Make_It_As_Header(wsResultat.Range("A1:F1"))

    'Initialiser le dictionnaire pour les résultats (Nom du client, Solde)
    Dim dictClients As Object ' Utilisez un dictionnaire pour stocker les résultats
    Set dictClients = CreateObject("Scripting.Dictionary")
    
    'Date actuelle pour le calcul de l'âge des factures
    Dim dateAujourdhui As Date
    dateAujourdhui = Date
    
    'Boucle sur les factures
    Dim derniereLigne As Long
    derniereLigne = wsFactures.Cells(wsFactures.rows.count, "A").End(xlUp).Row
    Dim rngFactures As Range
    Set rngFactures = wsFactures.Range("A3:A" & derniereLigne) '2 lignes d'entête
    
    Dim client As String
    Dim numFacture As String
    Dim dateFacture As Date
    Dim dateDue As Date
    Dim montantFacture As Currency
    Dim montantPaye As Currency
    Dim montantRestant As Currency
    Dim ageFacture As Long
    Dim tranche As String
    Dim i As Long
    
    For i = 1 To rngFactures.rows.count
        'Récupérer les données de la facture directement du Range
        client = rngFactures.Cells(i, 4).value
        'Obtenir le nom du client pour trier par nom de client plutôt que par code de client
        client = Fn_Get_Client_Name(client)
        numFacture = rngFactures.Cells(i, 1).value
        dateFacture = rngFactures.Cells(i, 2).value
        dateDue = rngFactures.Cells(i, 7).value
        montantFacture = CCur(rngFactures.Cells(i, 8).value)
        
        'Obtenir les paiemnets pour cette facture
        montantPaye = CCur(Application.WorksheetFunction.SumIf(wsPaiements.Range("B:B"), numFacture, wsPaiements.Range("E:E")))
        montantRestant = montantFacture - montantPaye
        
        'Calcul de l'âge de la facture
        ageFacture = WorksheetFunction.Max(dateAujourdhui - dateDue, 0)
        
        'Détermine la tranche d'âge
        Select Case ageFacture
            Case 0 To 30
                tranche = "0-30 jours"
            Case 31 To 60
                tranche = "31-60 jours"
            Case 61 To 90
                tranche = "61-90 jours"
            Case Is > 90
                tranche = "90+ jours"
            Case Else
                tranche = "Non défini"
        End Select
        
        Dim ngPaiements As Range
        Dim rowOffset As Long
        Dim tableau As Variant
        'Ajouter les données au dictionnaire en fonction du niveau de détail
        Select Case LCase(niveauDetail)
            Case "client"
                If Not dictClients.Exists(client) Then
                    dictClients.add client, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
                End If
                tableau = dictClients(client) 'Obtenir le tableau
                
                'Ajouter le montant le montant de la facture au total (0)
                tableau(0) = tableau(0) + montantRestant
                
                'Ajouter le montant restant à la tranche correspondante (1 @ 4)
                Select Case tranche
                    Case "0-30 jours"
                        tableau(1) = tableau(1) + montantRestant
                    Case "31-60 jours"
                        tableau(2) = tableau(2) + montantRestant
                    Case "61-90 jours"
                        tableau(3) = tableau(3) + montantRestant
                    Case "90+ jours"
                        tableau(4) = tableau(4) + montantRestant
                End Select
                dictClients(client) = tableau ' Remplacer le tableau dans le dictionnaire
            Case "facture"
                'Ajouter chaque facture avec son montant restant dû
                wsResultat.Cells(i, 1).value = client
                wsResultat.Cells(i, 2).value = numFacture
                wsResultat.Cells(i, 3).value = dateFacture
                wsResultat.Cells(i, 4).value = montantFacture
                wsResultat.Cells(i, 5).value = montantRestant
                wsResultat.Cells(i, 6).value = tranche
                
            Case "transaction"
                'Type Facture
                wsResultat.Cells(rowOffset, 1).value = client
                wsResultat.Cells(rowOffset, 2).value = "Facture"
                wsResultat.Cells(rowOffset, 3).value = numFacture
                wsResultat.Cells(rowOffset, 4).value = montantFacture
                wsResultat.Cells(rowOffset, 5).value = tranche
                wsResultat.Cells(rowOffset, 6).value = dateFacture
                wsResultat.Cells(rowOffset, 7).value = ""
                rowOffset = rowOffset + 1
                
                'Type paiements
                Dim rngPaiementsAssoc As Range
                Set rngPaiementsAssoc = wsPaiements.Range("B:B").Find(numFacture, , xlValues, xlWhole)
                Do Until rngPaiementsAssoc Is Nothing
                    wsResultat.Cells(rowOffset, 1).value = client
                    wsResultat.Cells(rowOffset, 2).value = "Paiement"
                    wsResultat.Cells(rowOffset, 3).value = rngPaiementsAssoc.value
                    wsResultat.Cells(rowOffset, 4).value = rngPaiementsAssoc.Offset(0, 1).value ' Montant du paiement
                    wsResultat.Cells(rowOffset, 5).value = tranche
                    wsResultat.Cells(rowOffset, 6).value = ""
                    wsResultat.Cells(rowOffset, 7).value = rngPaiementsAssoc.Offset(0, 2).value ' Date du paiement
                    
                    Set rngPaiementsAssoc = wsPaiements.Range("B:B").FindNext(rngPaiementsAssoc)
                    rowOffset = rowOffset + 1
                Loop
        End Select
    Next i
    
    'Si niveau de détail est "Client", ajouter les résultats au tableau final
    If LCase(niveauDetail) = "client" Then
        i = 2
        Dim cle As Variant
        For Each cle In dictClients.keys
            If dictClients(cle)(0) <> 0 Then
                wsResultat.Cells(i, 1).value = cle ' Nom du client
                wsResultat.Cells(i, 2).value = dictClients(cle)(0) ' Total
                wsResultat.Cells(i, 3).value = dictClients(cle)(1) ' 0-30 jours
                wsResultat.Cells(i, 4).value = dictClients(cle)(2) ' 31-60 jours
                wsResultat.Cells(i, 5).value = dictClients(cle)(3) ' 61-90 jours
                wsResultat.Cells(i, 6).value = dictClients(cle)(4) ' 90+ jours
                i = i + 1
            End If
        Next cle
    End If
    
    'Tri alphabétique par nom de client
    derniereLigne = wsResultat.Cells(wsResultat.rows.count, "A").End(xlUp).Row
    Set rngResultat = wsResultat.Range("A1:F" & derniereLigne)
    With wsResultat.Sort
        .SortFields.clear
        .SortFields.add key:=wsResultat.Range("A2"), Order:=xlAscending 'Trier par la colonne A
        .SetRange rngResultat
        .Header = xlYes
        .Apply
    End With
    
    With wsResultat
        .columns("A:A").ColumnWidth = 55
        .columns("B:F").ColumnWidth = 14
        .Range("B" & derniereLigne + 2).formula = "=Sum(B2:B" & derniereLigne & ")"
        .Range("C" & derniereLigne + 2).formula = "=Sum(C2:C" & derniereLigne & ")"
        .Range("D" & derniereLigne + 2).formula = "=Sum(D2:D" & derniereLigne & ")"
        .Range("E" & derniereLigne + 2).formula = "=Sum(E2:E" & derniereLigne & ")"
        .Range("F" & derniereLigne + 2).formula = "=Sum(F2:F" & derniereLigne & ")"
        .Range("B" & derniereLigne + 2 & ":F" & derniereLigne + 2).Font.Bold = True
    End With
    
    MsgBox "Liste âgée générée avec succès."
    
End Sub

