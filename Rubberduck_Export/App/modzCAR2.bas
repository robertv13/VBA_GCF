Attribute VB_Name = "modzCAR2"
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
    Set wsResultat = ThisWorkbook.Sheets("X_Liste_Âgée_CAR")
    If Not wsResultat Is Nothing Then wsResultat.delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsResultat = ThisWorkbook.Sheets.add
    wsResultat.name = "X_Liste_Âgée_CAR"
    Dim rngResultat As Range
    Set rngResultat = wsResultat.Range("A1")
    
    'Demander à l'utilisateur le niveau de détail
    Dim niveauDetail As String
    niveauDetail = InputBox("Choisissez le niveau de détail : Client, Facture, Transaction")
    
    'Entêtes de colonnes en fonction du niveau de détail
    If LCase(niveauDetail) = "client" Then
        wsResultat.Range("A1:F1").value = Array("Client", "Solde", "0-30 jours", "31-60 jours", "61-90 jours", "90+ jours")
        Call Make_It_As_Header(wsResultat.Range("A1:F1"))
    End If

    'Entêtes de colonnes en fonction du niveau de détail (Facture)
    If LCase(niveauDetail) = "facture" Then
        wsResultat.Range("A1:H1").value = Array("Client", "No. Facture", "Date Facture", "Solde", "0-30 jours", "31-60 jours", "61-90 jours", "90+ jours")
        Call Make_It_As_Header(wsResultat.Range("A1:H1"))
    End If

    'Entêtes de colonnes en fonction du niveau de détail (Transaction)
    If LCase(niveauDetail) = "transaction" Then
        wsResultat.Range("A1:H1").value = Array("Client", "No. Facture", "Date", "Montant", "0-30 jours", "31-60 jours", "61-90 jours", "90+ jours")
        Call Make_It_As_Header(wsResultat.Range("A1:H1"))
    End If

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
    Dim i As Long, r As Long
    
    r = 1
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
        
        'Exclus les soldes de facture à 0,00 $
        If montantRestant = 0 Then
            GoTo Next_Invoice
        End If
        
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
                dictClients(client) = tableau ' Replacer le tableau dans le dictionnaire
            Case "facture"
                'Ajouter chaque facture avec son montant restant dû
                r = r + 1
                wsResultat.Cells(r, 1).value = client
                wsResultat.Cells(r, 2).value = numFacture
                wsResultat.Cells(r, 3).value = dateFacture
                wsResultat.Cells(r, 4).value = montantRestant
                Select Case tranche
                    Case "0-30 jours"
                        wsResultat.Cells(r, 5).value = montantRestant
                    Case "31-60 jours"
                        wsResultat.Cells(r, 6).value = montantRestant
                    Case "61-90 jours"
                        wsResultat.Cells(r, 7).value = montantRestant
                    Case "90+ jours"
                        wsResultat.Cells(r, 8).value = montantRestant
                End Select
                
            Case "transaction"
                'Type Facture
                r = r + 1
                wsResultat.Cells(r, 1).value = client
                wsResultat.Cells(r, 2).value = numFacture
                wsResultat.Cells(r, 3).value = dateFacture
                wsResultat.Cells(r, 4).value = montantFacture
                Select Case tranche
                    Case "0-30 jours"
                        wsResultat.Cells(r, 5).value = montantRestant
                    Case "31-60 jours"
                        wsResultat.Cells(r, 6).value = montantRestant
                    Case "61-90 jours"
                        wsResultat.Cells(r, 7).value = montantRestant
                    Case "90+ jours"
                        wsResultat.Cells(r, 8).value = montantRestant
                End Select
                
                'Transactions de type paiements
                Dim rngPaiementsAssoc As Range
                Dim firstAddress As String
                Set rngPaiementsAssoc = wsPaiements.Range("B:B").Find(numFacture, LookIn:=xlValues, LookAt:=xlWhole)
                If Not rngPaiementsAssoc Is Nothing Then
                    firstAddress = rngPaiementsAssoc.Address
                        Do
                        r = r + 1
                        wsResultat.Cells(r, 1).value = client
                        wsResultat.Cells(r, 2).value = numFacture
                        wsResultat.Cells(r, 3).value = rngPaiementsAssoc.Offset(0, 2).value
                        wsResultat.Cells(r, 4).value = -rngPaiementsAssoc.Offset(0, 3).value ' Montant du paiement
                        Set rngPaiementsAssoc = wsPaiements.columns("B:B").FindNext(rngPaiementsAssoc)
                    Loop While Not rngPaiementsAssoc Is Nothing And rngPaiementsAssoc.Address <> firstAddress
                End If
        End Select

Next_Invoice:
    Next i
    
    'Si niveau de détail est "Client", ajouter les résultats au tableau final
    If LCase(niveauDetail) = "client" Then
        i = 2
        Dim cle As Variant
        For Each cle In dictClients.keys
            wsResultat.Cells(i, 1).value = cle ' Nom du client
            wsResultat.Cells(i, 2).value = dictClients(cle)(0) ' Total
            wsResultat.Cells(i, 3).value = dictClients(cle)(1) ' 0-30 jours
            wsResultat.Cells(i, 4).value = dictClients(cle)(2) ' 31-60 jours
            wsResultat.Cells(i, 5).value = dictClients(cle)(3) ' 61-90 jours
            wsResultat.Cells(i, 6).value = dictClients(cle)(4) ' 90+ jours
            i = i + 1
        Next cle
    End If
    
    'Tri alphabétique par nom de client
    derniereLigne = wsResultat.Cells(wsResultat.rows.count, "A").End(xlUp).Row
    Set rngResultat = wsResultat.Range("A1:H" & derniereLigne)
    With wsResultat.Sort
        .SortFields.clear
        .SortFields.add key:=wsResultat.Range("A2"), Order:=xlAscending 'Trier par la colonne A
        .SetRange rngResultat
        .Header = xlYes
        .Apply
    End With
    
    derniereLigne = derniereLigne + 2
    With wsResultat
        .columns("A:A").ColumnWidth = 55
        .columns("B:H").ColumnWidth = 12
        Select Case LCase(niveauDetail)
            Case "client"
                .Range("C2:F" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("B" & derniereLigne).formula = "=Sum(B2:B" & derniereLigne - 2 & ")"
                .Range("C" & derniereLigne).formula = "=Sum(C2:C" & derniereLigne - 2 & ")"
                .Range("D" & derniereLigne).formula = "=Sum(D2:D" & derniereLigne - 2 & ")"
                .Range("E" & derniereLigne).formula = "=Sum(E2:E" & derniereLigne - 2 & ")"
                .Range("F" & derniereLigne).formula = "=Sum(F2:F" & derniereLigne - 2 & ")"
                .Range("B" & derniereLigne & ":F" & derniereLigne).Font.Bold = True
            Case "facture"
                .columns("B:C").HorizontalAlignment = xlCenter
                .Range("D2:H" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("D" & derniereLigne).formula = "=Sum(D2:D" & derniereLigne - 2 & ")"
                .Range("E" & derniereLigne).formula = "=Sum(E2:E" & derniereLigne - 2 & ")"
                .Range("F" & derniereLigne).formula = "=Sum(F2:F" & derniereLigne - 2 & ")"
                .Range("G" & derniereLigne).formula = "=Sum(G2:G" & derniereLigne - 2 & ")"
                .Range("H" & derniereLigne).formula = "=Sum(H2:H" & derniereLigne - 2 & ")"
                .Range("D" & derniereLigne & ":H" & derniereLigne).Font.Bold = True
            Case "transaction"
                .columns("B:C").HorizontalAlignment = xlCenter
                .Range("D2:H" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("D" & derniereLigne).formula = "=Sum(D2:D" & derniereLigne - 2 & ")"
                .Range("E" & derniereLigne).formula = "=Sum(E2:E" & derniereLigne - 2 & ")"
                .Range("F" & derniereLigne).formula = "=Sum(F2:F" & derniereLigne - 2 & ")"
                .Range("G" & derniereLigne).formula = "=Sum(G2:G" & derniereLigne - 2 & ")"
                .Range("H" & derniereLigne).formula = "=Sum(H2:H" & derniereLigne - 2 & ")"
                .Range("D" & derniereLigne & ":H" & derniereLigne).Font.Bold = True
        End Select
    End With
    
    'Result print setup - 2024-08-31 @ 12:19
    Dim lastUsedRow As Long
    lastUsedRow = derniereLigne
    
    Dim rngToPrint As Range:
    Select Case LCase(niveauDetail)
        Case "client"
            Set rngToPrint = wsResultat.Range("A2:F" & lastUsedRow)
        Case "facture"
            Set rngToPrint = wsResultat.Range("A2:H" & lastUsedRow)
        Case "transaction"
            Set rngToPrint = wsResultat.Range("A2:H" & lastUsedRow)
    End Select
    Call Apply_Conditional_Formatting_Alternate(rngToPrint, 1, False)
    With rngToPrint.Font
        .name = "Segoe UI"
        .size = 9
    End With
    Dim header1 As String: header1 = "Liste âgée des comptes clients"
    Dim header2 As String
    If LCase(niveauDetail) = "client" Then
        header2 = "1 ligne par client"
    ElseIf LCase(niveauDetail) = "facture" Then
        header2 = "1 ligne par Facture"
    Else
        header2 = "1 ligne par transaction"
    End If
    header2 = "Par ordre alphabétique - " & header2
    
    Call Simple_Print_Setup(wsResultat, rngToPrint, header1, header2, "$1:$1", "L")
    
    MsgBox "La préparation est terminé" & vbNewLine & vbNewLine & "Voir la feuille 'X_Liste_Âgée_CAR'", vbInformation
    
    ThisWorkbook.Worksheets("X_Liste_Âgée_CAR").Activate
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set wsResultat = Nothing

End Sub

