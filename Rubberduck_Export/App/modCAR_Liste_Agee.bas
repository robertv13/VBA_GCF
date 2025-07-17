Attribute VB_Name = "modCAR_Liste_Agee"
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder("Rapport_ListeAgéeCC")

Option Explicit

Sub CC_PreparerListeAgee_Click()

    Dim ws As Worksheet
    Set ws = wshCAR_Liste_Agee
    
    Call EffacerResultatAnterieur(ws)
    
    Call CreerListeAgee

End Sub

Sub CreerListeAgee() '2024-09-08 @ 15:55

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("CreerListeAgee:CreerListeAgee", vbNullString, 0)
   
    Application.ScreenUpdating = False
    
    'Débloque la feuille
    ActiveSheet.Unprotect
    
    'Initialiser les feuilles nécessaires
    Dim wsFactures As Worksheet: Set wsFactures = wsdFAC_Comptes_Clients
    Dim wsPaiements As Worksheet: Set wsPaiements = wsdENC_Details
    Dim wsRégularisations As Worksheet: Set wsRégularisations = wsdCC_Regularisations
    
    'Cache les 2 formes de navigation (shpVersBas & shpVersHaut)
    Call GererBoutonsNavigation(False)
    
    Application.ScreenUpdating = False

    'Niveau de détail
    Dim niveauDetail As String
    niveauDetail = wshCAR_Liste_Agee.Range("B4").Value
    
    Application.EnableEvents = False
    
    'Entêtes de colonnes en fonction du niveau de détail
    If LCase$(niveauDetail) = "client" Then
        wshCAR_Liste_Agee.Range("B8:G8").Value = Array("Client", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Agee.Range("B8:G8"), RGB(84, 130, 53))
    End If

    'Entêtes de colonnes en fonction du niveau de détail (Facture)
    If LCase$(niveauDetail) = "facture" Then
        wshCAR_Liste_Agee.Range("B8:I8").Value = Array("Client", "No. Facture", "Date Facture", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Agee.Range("B8:I8"), RGB(84, 130, 53))
    End If

    'Entêtes de colonnes en fonction du niveau de détail (Transaction)
    If LCase$(niveauDetail) = "transaction" Then
        wshCAR_Liste_Agee.Range("B8:J8").Value = Array("Client", "No. Facture", "Type", "Date", "Montant", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Agee.Range("B8:J8"), RGB(84, 130, 53))
    End If

    Application.EnableEvents = True

    'Initialiser le dictionnaire pour les résultats (Nom du client, Solde)
    Dim dictClients As Object 'Utilisez un dictionnaire pour stocker les résultats
    Set dictClients = CreateObject("Scripting.Dictionary")
    
    'Date actuelle pour le calcul de l'âge des factures
    Dim dateAujourdhui As Date
    dateAujourdhui = Date
    
    'Boucle sur les factures
    Dim DerniereLigne As Long
    DerniereLigne = wsFactures.Cells(wsFactures.Rows.count, 1).End(xlUp).Row
    Dim rngFactures As Range
    Set rngFactures = wsFactures.Range("A3:A" & DerniereLigne) '2 lignes d'entête
    
    Dim client As String, numFacture As String
    Dim dateFacture As Date, dateDue As Date
    Dim montantFacture As Currency, montantPaye As Currency, montantRegul As Currency, montantRestant As Currency
    Dim trancheAge As String
    Dim ageFacture As Long, i As Long, r As Long
    
    Application.EnableEvents = False

    r = 8
    For i = 1 To rngFactures.Rows.count
        'Récupérer les données de la facture directement du Range
        numFacture = CStr(rngFactures.Cells(i, fFacCCInvNo).Value)
        'Do not process non Confirmed invoice
        If Fn_Get_Invoice_Type(numFacture) <> "C" Then
            GoTo Next_Invoice
        End If
        
        'Est-ce que la facture est à l'intérieur de la date limite ?
        dateFacture = rngFactures.Cells(i, fFacCCInvoiceDate).Value
        If rngFactures.Cells(i, fFacCCInvoiceDate).Value > CDate(wshCAR_Liste_Agee.Range("H4").Value) Then
            Debug.Print "#022 - Comparaison de date - " & rngFactures.Cells(i, fFacCCInvoiceDate).Value & " .vs. " & wshCAR_Liste_Agee.Range("H4").Value
            GoTo Next_Invoice
        End If
        
        client = rngFactures.Cells(i, fFacCCCodeClient).Value
        'Obtenir le nom du client (MF) pour trier par nom de client plutôt que par code de client
        client = Fn_Get_Client_Name(client)
        dateDue = rngFactures.Cells(i, fFacCCDueDate).Value
        montantFacture = CCur(rngFactures.Cells(i, fFacCCTotal).Value)
        
        'Obtenir les paiements et régularisations pour cette facture
        montantPaye = Fn_Obtenir_Paiements_Facture(numFacture, wshCAR_Liste_Agee.Range("H4").Value)
        montantRegul = Fn_Obtenir_Régularisations_Facture(numFacture, wshCAR_Liste_Agee.Range("H4").Value)
        
        montantRestant = montantFacture - montantPaye + montantRegul
        
        'Exclus les soldes de facture à 0,00 $ SI ET SEULMENT SI F4 = "NON"
        If UCase$(wshCAR_Liste_Agee.Range("F4").Value) = "NON" And montantRestant = 0 Then
            GoTo Next_Invoice
        End If
        
        'Calcul de l'âge de la facture
        ageFacture = WorksheetFunction.Max(dateAujourdhui - dateDue, 0)
        
        'Détermine la trancheAge d'âge
        Select Case ageFacture
            Case 0 To 30
                trancheAge = "- de 30 jours"
            Case 31 To 60
                trancheAge = "31 @ 60 jours"
            Case 61 To 90
                trancheAge = "61 @ 90 jours"
            Case Is > 90
                trancheAge = "+ de 90 jours"
            Case Else
                trancheAge = "Non défini"
        End Select
        
        Dim rngPaiements As Range
        Dim RowOffset As Long
        Dim tableau As Variant
        'Ajouter les données au dictionnaire en fonction du niveau de détail
        Select Case LCase$(niveauDetail)
            Case "client"
                If Not dictClients.Exists(client) Then
                    dictClients.Add client, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
                End If
                tableau = dictClients(client) 'Obtenir le tableau a partir du dictionary
                
                'Ajouter le solde de la facture au total (0)
                tableau(0) = tableau(0) + montantRestant
                
                'Ajouter le montant restant à la trancheAge correspondante (1 @ 4)
                Select Case trancheAge
                    Case "- de 30 jours"
                        tableau(1) = tableau(1) + montantRestant
                    Case "31 @ 60 jours"
                        tableau(2) = tableau(2) + montantRestant
                    Case "61 @ 90 jours"
                        tableau(3) = tableau(3) + montantRestant
                    Case "+ de 90 jours"
                        tableau(4) = tableau(4) + montantRestant
                End Select
                dictClients(client) = tableau ' Replacer le tableau dans le dictionnaire
            
            Case "facture"
                'Ajouter chaque facture avec son montant restant dû
                r = r + 1
                wshCAR_Liste_Agee.Cells(r, 2).Value = client
                wshCAR_Liste_Agee.Cells(r, 3).Value = numFacture
                wshCAR_Liste_Agee.Cells(r, 4).Value = dateFacture
                wshCAR_Liste_Agee.Cells(r, 4).NumberFormat = wsdADMIN.Range("B1").Value
                wshCAR_Liste_Agee.Cells(r, 5).Value = montantRestant
                Select Case trancheAge
                    Case "- de 30 jours"
                        wshCAR_Liste_Agee.Cells(r, 6).Value = montantRestant
                    Case "31 @ 60 jours"
                        wshCAR_Liste_Agee.Cells(r, 7).Value = montantRestant
                    Case "61 @ 90 jours"
                        wshCAR_Liste_Agee.Cells(r, 8).Value = montantRestant
                    Case "+ de 90 jours"
                        wshCAR_Liste_Agee.Cells(r, 9).Value = montantRestant
                End Select
                
            Case "transaction"
                'La facture en premier...
                r = r + 1
                wshCAR_Liste_Agee.Cells(r, 2).Value = client
                wshCAR_Liste_Agee.Cells(r, 3).Value = numFacture
                wshCAR_Liste_Agee.Cells(r, 4).Value = "Facture"
                wshCAR_Liste_Agee.Cells(r, 5).Value = dateFacture
                wshCAR_Liste_Agee.Cells(r, 5).NumberFormat = wsdADMIN.Range("B1").Value
                wshCAR_Liste_Agee.Cells(r, 6).Value = montantFacture
                Select Case trancheAge
                    Case "- de 30 jours"
                        wshCAR_Liste_Agee.Cells(r, 7).Value = montantRestant
                    Case "31 @ 60 jours"
                        wshCAR_Liste_Agee.Cells(r, 8).Value = montantRestant
                    Case "61 @ 90 jours"
                        wshCAR_Liste_Agee.Cells(r, 9).Value = montantRestant
                    Case "+ de 90 jours"
                        wshCAR_Liste_Agee.Cells(r, 10).Value = montantRestant
                End Select
                
                'Transactions de paiements par la suite
                Dim rngPaiementsAssoc As Range
                Dim pmtFirstAddress As String
                'Obtenir tous les paiements pour la facture
                Set rngPaiementsAssoc = wsPaiements.Range("B:B").Find(numFacture, LookIn:=xlValues, LookAt:=xlWhole)
                If Not rngPaiementsAssoc Is Nothing Then
                    pmtFirstAddress = rngPaiementsAssoc.Address
                    Do
                        If rngPaiementsAssoc.offset(0, 2).Value <= CDate(wshCAR_Liste_Agee.Range("H4").Value) Then
                            r = r + 1
                            wshCAR_Liste_Agee.Cells(r, 2).Value = client
                            wshCAR_Liste_Agee.Cells(r, 3).Value = numFacture
                            wshCAR_Liste_Agee.Cells(r, 4).Value = "Paiement"
                            wshCAR_Liste_Agee.Cells(r, 5).Value = rngPaiementsAssoc.offset(0, 2).Value
                            wshCAR_Liste_Agee.Cells(r, 6).Value = -rngPaiementsAssoc.offset(0, 3).Value 'Montant du paiement
                        End If
                        Set rngPaiementsAssoc = wsPaiements.Columns("B:B").FindNext(rngPaiementsAssoc)
                    Loop While Not rngPaiementsAssoc Is Nothing And rngPaiementsAssoc.Address <> pmtFirstAddress
                End If
                'Transactions de régularisations par la suite
                Dim rngRégularisationAssoc As Range
                Dim regulFirstAddress As String
                'Obtenir toutes les régularisations pour la facture
                Set rngRégularisationAssoc = wsRégularisations.Range("B:B").Find(numFacture, LookIn:=xlValues, LookAt:=xlWhole)
                If Not rngRégularisationAssoc Is Nothing Then
                    regulFirstAddress = rngRégularisationAssoc.Address
                    Do
                        If rngRégularisationAssoc.offset(0, 1).Value <= CDate(wshCAR_Liste_Agee.Range("H4").Value) Then
                            r = r + 1
                            wshCAR_Liste_Agee.Cells(r, 2).Value = client
                            wshCAR_Liste_Agee.Cells(r, 3).Value = numFacture
                            wshCAR_Liste_Agee.Cells(r, 4).Value = "Régularisation"
                            wshCAR_Liste_Agee.Cells(r, 5).Value = rngRégularisationAssoc.offset(0, 1).Value
                            wshCAR_Liste_Agee.Cells(r, 6).Value = rngRégularisationAssoc.offset(0, 4).Value + _
                                rngRégularisationAssoc.offset(0, 5).Value + _
                                rngRégularisationAssoc.offset(0, 6).Value + _
                                rngRégularisationAssoc.offset(0, 7).Value
                        End If
                        Set rngRégularisationAssoc = wsRégularisations.Columns("B:B").FindNext(rngRégularisationAssoc)
                    Loop While Not rngRégularisationAssoc Is Nothing And rngRégularisationAssoc.Address <> regulFirstAddress
                End If
        End Select

Next_Invoice:
    Next i
    
    Application.EnableEvents = True
    
    'Si niveau de détail est "client", ajouter les soldes du client (dictionary) au tableau final
    If LCase$(niveauDetail) = "client" Then
        r = 8
        Dim cle As Variant
        
        Application.EnableEvents = False
        
        For Each cle In dictClients.keys
            r = r + 1
            wshCAR_Liste_Agee.Cells(r, 2).Value = cle ' Nom du client
            wshCAR_Liste_Agee.Cells(r, 3).Value = dictClients(cle)(0) ' Total
            wshCAR_Liste_Agee.Cells(r, 4).Value = dictClients(cle)(1) ' - de 30 jours
            wshCAR_Liste_Agee.Cells(r, 5).Value = dictClients(cle)(2) ' 31 @ 60 jours
            wshCAR_Liste_Agee.Cells(r, 6).Value = dictClients(cle)(3) ' 61 @ 90 jours
            wshCAR_Liste_Agee.Cells(r, 7).Value = dictClients(cle)(4) ' + de 90 jours
        Next cle
        
        Application.EnableEvents = True

    End If
    
    'Tri alphabétique par nom de client
    Dim rngResultat As Range
    Set rngResultat = wshCAR_Liste_Agee.Range("B8:J" & DerniereLigne)
    DerniereLigne = wshCAR_Liste_Agee.Cells(wshCAR_Liste_Agee.Rows.count, "B").End(xlUp).Row
    
    Application.EnableEvents = False
    
    Dim ordreTri As String
    If DerniereLigne > 9 Then 'Le tri n'est peut-être pas nécessaire
        With wshCAR_Liste_Agee.Sort
            .SortFields.Clear
            If wshCAR_Liste_Agee.Range("D4").Value = "Nom de client" Then
                .SortFields.Add _
                    key:=wshCAR_Liste_Agee.Range("B8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par nom de client
                ordreTri = "Ordre de nom de client"
            Else
                .SortFields.Add _
                    key:=wshCAR_Liste_Agee.Range("C8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par numéro de facture
                .SortFields.Add _
                    key:=wshCAR_Liste_Agee.Range("D8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier date de transaction
                ordreTri = "Ordre de numéro de facture"
            End If
            .SetRange rngResultat
            .Header = xlYes
            .Apply
        End With
    End If
    
    DerniereLigne = DerniereLigne + 2
    
    Dim t(0 To 4) As Currency
    
    With wshCAR_Liste_Agee
        .Columns("B:B").ColumnWidth = 50
        .Columns("C:J").ColumnWidth = 13
        Select Case LCase$(niveauDetail)
            Case "client"
                .Range("C9:G" & DerniereLigne).NumberFormat = "#,##0.00 $"
                .Range("C9:G" & DerniereLigne).HorizontalAlignment = xlRight
                .Range("C" & DerniereLigne).formula = "=Sum(C9:C" & DerniereLigne - 2 & ")"
                t(0) = .Range("C" & DerniereLigne).Value
                .Range("D" & DerniereLigne).formula = "=Sum(D9:D" & DerniereLigne - 2 & ")"
                t(1) = .Range("D" & DerniereLigne).Value
                .Range("E" & DerniereLigne).formula = "=Sum(E9:E" & DerniereLigne - 2 & ")"
                t(2) = .Range("E" & DerniereLigne).Value
                .Range("F" & DerniereLigne).formula = "=Sum(F9:F" & DerniereLigne - 2 & ")"
                t(3) = .Range("F" & DerniereLigne).Value
                .Range("G" & DerniereLigne).formula = "=Sum(G9:G" & DerniereLigne - 2 & ")"
                t(4) = .Range("G" & DerniereLigne).Value
                .Range("C" & DerniereLigne & ":G" & DerniereLigne).Font.Bold = True
            Case "facture"
                .Range("C9:C" & DerniereLigne).HorizontalAlignment = xlCenter
                .Range("D9:D" & DerniereLigne).HorizontalAlignment = xlCenter
                .Range("E9:I" & DerniereLigne).NumberFormat = "#,##0.00 $"
                .Range("E9:I" & DerniereLigne).HorizontalAlignment = xlRight
                .Range("E" & DerniereLigne).formula = "=Sum(E9:E" & DerniereLigne - 2 & ")"
                t(0) = .Range("E" & DerniereLigne).Value
                .Range("F" & DerniereLigne).formula = "=Sum(F9:F" & DerniereLigne - 2 & ")"
                t(1) = .Range("F" & DerniereLigne).Value
                .Range("G" & DerniereLigne).formula = "=Sum(G9:G" & DerniereLigne - 2 & ")"
                t(2) = .Range("G" & DerniereLigne).Value
                .Range("H" & DerniereLigne).formula = "=Sum(H9:H" & DerniereLigne - 2 & ")"
                t(3) = .Range("H" & DerniereLigne).Value
                .Range("I" & DerniereLigne).formula = "=Sum(I9:I" & DerniereLigne - 2 & ")"
                t(4) = .Range("I" & DerniereLigne).Value
                .Range("E" & DerniereLigne & ":I" & DerniereLigne).Font.Bold = True
            Case "transaction"
                .Columns("C:E").HorizontalAlignment = xlCenter
                .Columns("D").HorizontalAlignment = xlLeft
                .Range("F9:J" & DerniereLigne).NumberFormat = "#,##0.00 $"
                .Range("F9:J" & DerniereLigne).HorizontalAlignment = xlRight
                .Range("F" & DerniereLigne).formula = "=Sum(F9:F" & DerniereLigne - 2 & ")"
                t(0) = .Range("F" & DerniereLigne).Value
                .Range("G" & DerniereLigne).formula = "=Sum(G9:G" & DerniereLigne - 2 & ")"
                t(1) = .Range("G" & DerniereLigne).Value
                .Range("H" & DerniereLigne).formula = "=Sum(H9:H" & DerniereLigne - 2 & ")"
                t(2) = .Range("H" & DerniereLigne).Value
                .Range("I" & DerniereLigne).formula = "=Sum(I9:I" & DerniereLigne - 2 & ")"
                t(3) = .Range("I" & DerniereLigne).Value
                .Range("J" & DerniereLigne).formula = "=Sum(J9:J" & DerniereLigne - 2 & ")"
                t(4) = .Range("J" & DerniereLigne).Value
                .Range("F" & DerniereLigne & ":J" & DerniereLigne).Font.Bold = True
        End Select
        .Range("B" & DerniereLigne).Value = "Totaux de la liste"
        .Range("B" & DerniereLigne).Font.Bold = True
        DerniereLigne = DerniereLigne + 1
        
        'Ligne de pourcentages
        .Range("B" & DerniereLigne).Value = "Pourcentages"
        .Range("B" & DerniereLigne & ":J" & DerniereLigne).Font.Bold = True
        .Range("C" & DerniereLigne & ":J" & DerniereLigne).NumberFormat = "##0.00"
        .Range("C" & DerniereLigne & ":J" & DerniereLigne).HorizontalAlignment = xlRight
        Dim totalListe As Currency
        totalListe = t(0)
        If totalListe <> 0 Then
            Select Case LCase$(niveauDetail)
                Case "client"
                    .Range("C" & DerniereLigne).Value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("D" & DerniereLigne).Value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("E" & DerniereLigne).Value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("F" & DerniereLigne).Value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("G" & DerniereLigne).Value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
                Case "facture"
                    .Range("E" & DerniereLigne).Value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("F" & DerniereLigne).Value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("G" & DerniereLigne).Value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("H" & DerniereLigne).Value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("I" & DerniereLigne).Value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
                Case "transaction"
                    .Range("F" & DerniereLigne).Value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("G" & DerniereLigne).Value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("H" & DerniereLigne).Value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("I" & DerniereLigne).Value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("J" & DerniereLigne).Value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
            End Select
        End If
    End With
    
    Call GererBoutonsNavigation(True)
    
    Dim ligneVolet As Long
    ligneVolet = ActiveWindow.SplitRow
    Dim saveDerniereLigne As Long
    saveDerniereLigne = DerniereLigne
    
    If DerniereLigne > (20 + ligneVolet) Then
        DerniereLigne = DerniereLigne - 20
        Application.GoTo wshCAR_Liste_Agee.Cells(DerniereLigne, 1), Scroll:=True '2025-07-08 @ 13:39
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    DoEvents
   
    'Result print setup - 2024-08-31 @ 12:19
    Dim lastUsedRow As Long
    lastUsedRow = saveDerniereLigne
    
    Dim rngToPrint As Range:
    Select Case LCase$(niveauDetail)
        Case "client"
            Set rngToPrint = wshCAR_Liste_Agee.Range("B9:G" & lastUsedRow)
        Case "facture"
            Set rngToPrint = wshCAR_Liste_Agee.Range("B9:I" & lastUsedRow)
        Case "transaction"
            Set rngToPrint = wshCAR_Liste_Agee.Range("B9:J" & lastUsedRow)
    End Select
    
    Application.EnableEvents = False

    Call modAppli_Utils.AppliquerConditionalFormating(rngToPrint, 0, RGB(198, 224, 180))
    
    'Caractères pour le rapport
    With rngToPrint.Font
        .Name = "Aptos Narrow"
        .size = 10
    End With
    
    Application.EnableEvents = True
    
    DoEvents

    Dim header1 As String: header1 = "Liste âgée des comptes clients au " & wshCAR_Liste_Agee.Range("H4").Value
    Dim header2 As String
    If LCase$(niveauDetail) = "client" Then
        header2 = "1 ligne par client"
    ElseIf LCase$(niveauDetail) = "facture" Then
        header2 = "1 ligne par Facture"
    Else
        header2 = "1 ligne par transaction"
    End If
    header2 = ordreTri & " - " & header2
    
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wshCAR_Liste_Agee, rngToPrint, header1, header2, "$8:$8", "P")
    
    Application.ScreenUpdating = True
    
    MsgBox "La préparation de la liste âgée est terminée", vbInformation
    
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set cle = Nothing
    Set dictClients = Nothing
    Set rngFactures = Nothing
    Set rngPaiementsAssoc = Nothing
    Set rngResultat = Nothing
    Set rngToPrint = Nothing
    Set wsFactures = Nothing
    Set wsPaiements = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("CreerListeAgee:CreerListeAgee", vbNullString, startTime)
    
End Sub

Sub AfficherMenuContextuel(ByVal Target As Range) '2025-02-21 @ 19:10

    Dim menu As CommandBar
    Dim menuItem As CommandBarButton

    'Supprimer le menu contextuel personnalisé s'il existe déjà
    On Error Resume Next
    Application.CommandBars("FactureMenu").Delete
    On Error GoTo 0
    
    'Détermine les coordonnées de la colonne qui a été cliquée
    Dim numeroLigne As Long, numeroColonne As Long
    Call ExtraireLigneColonneCellule(Target.Address, numeroLigne, numeroColonne)
    
    Dim numeroFacture As String
    numeroFacture = ActiveSheet.Cells(numeroLigne, "C").Value
    If Trim$(numeroFacture) = vbNullString Then
        Exit Sub
    End If
    
    'Créer un nouveau menu contextuel
    Set menu = Application.CommandBars.Add(Name:="FactureMenu", position:=msoBarPopup, Temporary:=True)

    'Option # 1 - Visualiser la facture
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.Caption = "Visualiser la facture (format PDF)"
        menuItem.OnAction = "'modFAC_Interrogation.VisualiserFacturePDF """ & numeroFacture & """'"

    'Option # 2 - Visualiser la facture
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.Caption = "Envoi d'un rappel par courriel"
        menuItem.OnAction = "'EnvoyerRappelParCourriel """ & numeroFacture & """'"

    'Afficher le menu contextuel
    menu.ShowPopup

End Sub

Sub EnvoyerRappelParCourriel(noFact As String)

    Dim montantDu As String

    'Retrouver le code du client
    Dim codeClient As String
    Dim dateFact As Date
    Dim allCols As Variant
    allCols = Fn_Get_A_Row_From_A_Worksheet("FAC_Entête", noFact, fFacEInvNo)
    'Vérifier les résultats
    If IsArray(allCols) Then
        codeClient = allCols(fFacECustID)
        dateFact = allCols(fFacEDateFacture)
    Else
        MsgBox "Enregistrement '" & noFact & "' non trouvée !!!", vbCritical
        Exit Sub
    End If
    
    If codeClient = vbNullString Then
        MsgBox "Le code client pour cette facture est INVALIDE", vbCritical, "Information erronée / manquante"
        Exit Sub
    End If
    
    'Retrouver le nom du client, le nom du contact et son adresse courriel
    Dim clientNom As String
    Dim clientContactFact As String
    Dim clientCourriel As String
    allCols = Fn_Get_A_Row_From_A_Worksheet("BD_Clients", codeClient, fClntFMClientID)
    'Vérifier les résultats
    If IsArray(allCols) Then
        clientNom = allCols(fClntFMClientNom)
        'Élimine le ou les nom(s) de contact
        clientNom = Fn_Strip_Contact_From_Client_Name(clientNom)
        clientContactFact = Trim$(allCols(fClntFMContactFacturation)) + " "
        '0 à 2 adresses courriel
        clientCourriel = allCols(fClntFMCourrielFacturation)
    Else
        MsgBox "Enregistrement '" & codeClient & "' non trouvée !!!", vbCritical
        Exit Sub
    End If
    
    'Retrouver le solde de la facture, le total des paiements & le total des régularisations
    Dim factSolde As Currency
    Dim factSommePmts As Currency
    Dim factSommeRegul As Currency
    allCols = Fn_Get_A_Row_From_A_Worksheet("FAC_Comptes_Clients", noFact, fFacCCInvNo)
    'Vérifier les résultats
    If IsArray(allCols) Then
        factSolde = allCols(fFacCCBalance)
        factSommePmts = allCols(fFacCCTotalPaid)
        factSommeRegul = allCols(fFacCCTotalRegul)
    Else
        MsgBox "Enregistrement '" & noFact & "' non trouvée !!!", vbCritical
        Exit Sub
    End If
    
    'Vérification pour éviter d'envoyer un rappel pour une facture à 0 $ ou créditeur
    If factSolde <= 0 Then
        MsgBox "Il n'y a pas lieu d'envoyer un rappel" & vbNewLine & vbNewLine & _
                "pour cette facture. Solde à " & Format$(factSolde, "#,##0.00 $"), vbCritical, "Solde à 0,00 $ ou créditeur"
        Exit Sub
    End If
    
    'Vérifier si l'email est valide
    If Fn_ValiderCourriel(clientCourriel) = False Then
        MsgBox "L'adresse courriel est vide OU invalide" & vbNewLine & vbNewLine & _
                "pour ce client.", vbExclamation, "Impossible d'envoyer un rappel (Adresse courriel invalide)"
        Exit Sub
    End If

    'Ajouter la copie de la facture (format PDF)
    Dim attachmentFullPathName As String
    attachmentFullPathName = wsdADMIN.Range("F5").Value & gFACT_PDF_PATH & Application.PathSeparator & _
                     noFact & ".pdf"
    
    'Vérification de l'existence de la pièce jointe
    Dim fileExists As Boolean
    fileExists = Dir(attachmentFullPathName) <> vbNullString
    If Not fileExists Then
        MsgBox "La pièce jointe (Facture en format PDF) n'existe pas à" & vbNewLine & _
                    "l'emplacement spécifié, soit " & attachmentFullPathName, vbCritical
        GoTo Exit_Sub
    End If
    
    'Chemin du template (.oft) de courriel
    Dim templateFullPathName As String
    templateFullPathName = Environ$("appdata") & "\Microsoft\Templates\GCF_Rappel.oft"

    'Vérification de l'existence du template
    fileExists = Dir(templateFullPathName) <> vbNullString
    If Not fileExists Then
        MsgBox "Le gabarit 'GCF_Rappel.oft' est introuvable " & _
                    "à l'emplacement spécifié, soit " & Environ$("appdata") & "\Microsoft\Templates", _
                    vbCritical
        GoTo Exit_Sub
    End If
    
    'Initialiser Outlook
    On Error Resume Next
    Dim OutlookApp As Object
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    
    'Vérifier si Outlook est bien ouvert
    If OutlookApp Is Nothing Then
        MsgBox "Impossible d'ouvrir Outlook. Vérifiez votre installation.", vbCritical
        Exit Sub
    End If

    Dim mailItem As Object
    Set mailItem = OutlookApp.CreateItemFromTemplate(templateFullPathName)
    On Error GoTo 0

    mailItem.Attachments.Add attachmentFullPathName

    'Ajuster les derniers paramètres du courriel
    Dim adresseEmail  As Variant
    adresseEmail = Split(clientCourriel, "; ") '2025-03-02 @ 16:36
    Dim nbAdresseCourriel As Integer
    nbAdresseCourriel = UBound(adresseEmail)
    
    With mailItem
        Select Case nbAdresseCourriel
            Case 0
                .To = adresseEmail(0)
            Case Is > 0
                .To = adresseEmail(0)
                .cc = adresseEmail(1)
            Case Else
        End Select
        
        .Subject = wsdADMIN.Range("NomEntreprise") & " - Rappel pour facture impayée - " & clientNom & " - Facture # " & noFact
        .Display  'Pour afficher le mail avant envoi (remplacez par .Send pour envoyer directement)
    End With

Exit_Sub:

    ' Nettoyer les objets
    Set mailItem = Nothing
    Set OutlookApp = Nothing
    
End Sub

Sub EffacerResultatAnterieur(ws As Worksheet)

    ws.Unprotect
    
    'Efface les résultats antérieurs
    Dim lastUsedRow As Long
    With ws
        lastUsedRow = .Cells(.Rows.count, "B").End(xlUp).Row
        If lastUsedRow > 9 Then
            Application.EnableEvents = False
            'Effacement du contenu + marges (B8:J + 3 lignes de totaux)
            .Range("B8:J" & lastUsedRow + 3).Clear
            Application.EnableEvents = True
        End If
    End With
    
    With ws
        .Shapes("shpVersBas").Visible = False
        .Shapes("shpVersHaut").Visible = False
    End With
    
    ws.Activate
    DoEvents
    Application.ScreenUpdating = True
    
    ws.Protect userInterfaceOnly:=True

End Sub

Sub shpVersBas_Click() '2025-06-30 @ 10:59

    Call AllerAuCentreDesResultats

End Sub

Sub AllerAuCentreDesResultats() '2025-06-30@ 11:02

    Dim derLigne As Long
    Dim nbLignesVisibles As Long
    Dim ligneCible As Long

    With ActiveWindow.VisibleRange
        nbLignesVisibles = .Rows.count
    End With

    With ActiveSheet
        'Trouve la dernière ligne avec données sur colonne B
        derLigne = .Cells(.Rows.count, 2).End(xlUp).Row
        
        'Centre la ligne dans la fenêtre visible si possible
        ligneCible = Application.Max(1, derLigne - Int(nbLignesVisibles / 2))
        
        Application.GoTo Reference:=.Cells(ligneCible, 1), Scroll:=True
    End With
    
End Sub

Sub shpVersHaut_Click() '2025-06-30 @ 10:59

    Call RetournerEnHaut

End Sub

Sub RetournerEnHaut() '2025-06-30 @ 11:08

    Application.GoTo Reference:=ActiveSheet.Cells(1, 1), Scroll:=True
    
End Sub

Sub GererBoutonsNavigation(totauxImprimes As Boolean) '2025-06-30 @ 11:40

    Dim f As Worksheet: Set f = ActiveSheet
    Dim colDerniere As Long, colBoutons As Long
    Dim ligneTotaux As Long
    Dim shpVersHaut As Shape, shpVersBas As Shape

    'Déterminer la dernière colonne utilisée, selon le niveau de détail
    colDerniere = f.Range("B8").End(xlToRight).Column

    'Position des boutons (deuxième colonne après la dernière colonne utilisée)
    colBoutons = colDerniere + 1

    'Déterminer la ligne des totaux
    ligneTotaux = f.Cells(f.Rows.count, "B").End(xlUp).Row

    'Récupérer les formes déjà présentes
    On Error Resume Next
    Set shpVersHaut = f.Shapes("shpVersHaut")
    Set shpVersBas = f.Shapes("shpVersBas")
    On Error GoTo 0

    If Not shpVersBas Is Nothing Then
        If totauxImprimes Then
            With f.Cells(9, colBoutons)
                shpVersBas.Top = .Top
                shpVersBas.Left = f.Cells(9, colBoutons).Left + 15
            End With
        End If
        shpVersBas.Visible = totauxImprimes
    End If

    If Not shpVersHaut Is Nothing Then
        If totauxImprimes Then
            With f.Cells(ligneTotaux - 1, colBoutons)
                shpVersHaut.Top = .Top
                shpVersHaut.Left = f.Cells(ligneTotaux - 1, colBoutons).Left + 15
            End With
        End If
        shpVersHaut.Visible = totauxImprimes
    End If
    
    'Libérer la mémoire
    Set shpVersBas = Nothing
    Set shpVersHaut = Nothing
    
End Sub

'Sub AjouterBoutonRetourHaut() '2025-06-30 @ 11:06
'
'    Dim f As Worksheet: Set f = ActiveSheet
'    Dim forme As Shape
'    Dim derLigne As Long
'    Dim topPosition As Double
'    Dim leftPosition As Long
'    Dim existe As Boolean
'
'    'Recherche de la forme existante
'    On Error Resume Next
'    Set forme = f.Shapes("shpHaut")
'    existe = Not forme Is Nothing
'    On Error GoTo 0
'
'    'Détermination de la dernière ligne utilisée
'    derLigne = f.Cells(f.Rows.count, 2).End(xlUp).Row
'    topPosition = f.Rows(derLigne).Top
'    leftPosition = f.Cells(derLigne - 2, 11).Left
'
'    'Crée ou déplace le bouton
'    If Not existe Then
'        Set forme = f.Shapes.AddShape(msoShapeRoundedRectangle, 10, topPosition, 100, 25)
'        With forme
'            .Name = "shpHaut"
'            .TextFrame.Characters.Text = "Retour en haut"
'            .OnAction = "RetourAuDébut"
'            .Fill.ForeColor.RGB = RGB(200, 200, 255)
'        End With
'    Else
'        forme.Top = topPosition
'        forme.Left = leftPosition
'    End If
'
'End Sub

Sub shpRetournerMenuFacturation_Click()

    Call EffacerResultatAnterieur(wshCAR_Liste_Agee)
    
    Call RetournerMenuFacturation

End Sub

Sub RetournerMenuFacturation()

    wshCAR_Liste_Agee.Visible = xlSheetHidden
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select

End Sub


