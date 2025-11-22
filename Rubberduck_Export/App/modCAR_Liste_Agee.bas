Attribute VB_Name = "modCAR_Liste_Agee"
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder("Rapport_ListeAgéeCC")

Option Explicit

Sub PreparerListeAgeeCC_Click()

    Dim ws As Worksheet
    Set ws = wshCAR_Liste_Agee
    
    Call EffacerResultatAnterieur(ws)
    
    Call CreerListeAgeeCC

End Sub

Sub CreerListeAgeeCC() '2025-10-21 @ 09:36

    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modCAR_Liste_Agee:CreerListeAgeeCC", vbNullString, 0)
   
    Application.ScreenUpdating = False
    
    'Débloque la feuille
    ActiveSheet.Unprotect
    
    'Charger dans un dictionary les codes & noms de clients
    Dim wsBDClients As Worksheet: Set wsBDClients = wsdBD_Clients
    Dim lastRowBDClients As Long
    lastRowBDClients = wsBDClients.Cells(wsBDClients.Rows.count, 1).End(xlUp).Row
    
    Dim dictNomsClients As Object
    Set dictNomsClients = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, code As String, nom As String
    For i = 2 To lastRowBDClients
        code = Trim(wsBDClients.Cells(i, fClntFMClientID).Value)
        nom = Trim(wsBDClients.Cells(i, fClntFMClientNom).Value)
        If Len(code) > 0 And Len(nom) > 0 Then
            If Not dictNomsClients.Exists(code) Then
                dictNomsClients.Add code, nom
            End If
        End If
    Next i

    'Charger en mémoire (tableauFactures) les comptes-clients
    Dim wsFacturesCC As Worksheet
    Set wsFacturesCC = wsdFAC_Comptes_Clients
    Dim DerniereLigne As Long
    DerniereLigne = wsFacturesCC.Cells(wsFacturesCC.Rows.count, 1).End(xlUp).Row
    Dim tableauFactures As Variant
    tableauFactures = wsFacturesCC.Range("A3:M" & DerniereLigne).Value
    
    'Charger dans un dictionary les factures confirmées
    Dim wsEntete As Worksheet: Set wsEntete = wsdFAC_Entete
    Dim lastRow As Long
    lastRow = wsEntete.Cells(wsEntete.Rows.count, 1).End(xlUp).Row
    
    Dim dictTypesFactures As Object
    Set dictTypesFactures = CreateObject("Scripting.Dictionary")
    
    Dim numFac As String, typeFac As String
    For i = 2 To lastRow
        numFac = Trim(wsEntete.Cells(i, fFacEInvNo).Value)
        typeFac = UCase(Trim(wsEntete.Cells(i, fFacEACouC).Value))
        If Len(numFac) > 0 Then
            dictTypesFactures(numFac) = typeFac
        End If
    Next i
    
    'Charger dans un dictionary les transactions de paiement
    Dim wsPmt As Worksheet: Set wsPmt = wsdENC_Details
    Dim lastRowPmt As Long
    lastRowPmt = wsPmt.Cells(wsPmt.Rows.count, 2).End(xlUp).Row
    
    Dim dictPaiements As Object
    Set dictPaiements = CreateObject("Scripting.Dictionary")
    
    Dim datePmt As Date, montant As Currency
    For i = 2 To lastRowPmt
        numFac = Trim(wsPmt.Cells(i, fEncDInvNo).Value)
        If Len(numFac) > 0 And IsDate(wsPmt.Cells(i, fEncDPayDate).Value) Then
            datePmt = CDate(wsPmt.Cells(i, fEncDPayDate).Value)
            montant = wsPmt.Cells(i, fEncDPayAmount).Value
            If Not dictPaiements.Exists(numFac) Then dictPaiements.Add numFac, New Collection
            dictPaiements(numFac).Add Array(datePmt, montant)
        End If
    Next i
    
    'Charger dans un dictionary les transactions de régularisations
    Dim wsReg As Worksheet: Set wsReg = wsdCC_Regularisations
    Dim lastRowReg As Long
    lastRowReg = wsReg.Cells(wsReg.Rows.count, fREGULInvNo).End(xlUp).Row
    
    Dim dictRegul As Object
    Set dictRegul = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRowReg
        numFac = Trim(wsReg.Cells(i, fREGULInvNo).Value)
        If Len(numFac) > 0 And IsDate(wsReg.Cells(i, fREGULDate).Value) Then
            datePmt = CDate(wsReg.Cells(i, fREGULDate).Value)
            montant = wsReg.Cells(i, fREGULHono).Value + wsReg.Cells(i, fREGULFrais).Value + _
                      wsReg.Cells(i, fREGULTPS).Value + wsReg.Cells(i, fREGULTVQ).Value
            If Not dictRegul.Exists(numFac) Then dictRegul.Add numFac, New Collection
            dictRegul(numFac).Add Array(datePmt, montant)
        End If
    Next i
    
    'Cache les 2 formes de navigation (shpVersBas & shpVersHaut)
    Call GererBoutonsNavigation(False)
    
    'Niveau de détail
    Dim niveauDetail As String
    niveauDetail = wshCAR_Liste_Agee.Range("B4").Value
    
    Application.EnableEvents = False
    
    Dim nbColonnes As Long
    Select Case LCase$(niveauDetail)
        Case "client"
            nbColonnes = 6
            wshCAR_Liste_Agee.Range("B8:G8").Value = Array("Client", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
            Call CreerEnteteDeFeuille(wshCAR_Liste_Agee.Range("B8:G8"), RGB(84, 130, 53))
        Case "facture"
            nbColonnes = 8
            wshCAR_Liste_Agee.Range("B8:I8").Value = Array("Client", "No. Facture", "Date Facture", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
            Call CreerEnteteDeFeuille(wshCAR_Liste_Agee.Range("B8:I8"), RGB(84, 130, 53))
        Case "transaction"
            nbColonnes = 9
            wshCAR_Liste_Agee.Range("B8:J8").Value = Array("Client", "No. Facture", "Type", "Date", "Montant", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
            Call CreerEnteteDeFeuille(wshCAR_Liste_Agee.Range("B8:J8"), RGB(84, 130, 53))
    End Select

    Application.EnableEvents = True

    'Initialiser le dictionary pour les soldes par clients
    Dim dictSoldesClients As Object
    Set dictSoldesClients = CreateObject("Scripting.Dictionary")
    
    'Déterminer la taille Maximale de buffer
    Dim maxNombreLignes As Long
    maxNombreLignes = DerniereLigne + lastRowPmt + lastRowReg
    Dim buffer() As Variant
    Dim nbLignesMax As Long: nbLignesMax = maxNombreLignes
    ReDim buffer(1 To nbLignesMax + 20, 1 To nbColonnes) '20 lignes d'entête / totaux

    Dim client As String, numFacture As String
    Dim dateFacture As Date, dateDue As Date, dateLimite As Date
    Dim montantFacture As Currency, montantPaye As Currency, montantRegul As Currency, montantRestant As Currency
    Dim trancheAge As String
    Dim formatDate As String
    Dim ageFacture As Long, r As Long
    Dim inclureSoldesNuls As Boolean
    
    dateLimite = CDate(wshCAR_Liste_Agee.Range("H4").Value) '2025-10-21 @08:08
    inclureSoldesNuls = (UCase$(wshCAR_Liste_Agee.Range("F4").Value) <> "NON")
    formatDate = wsdADMIN.Range("USER_DATE_FORMAT").Value
    
    Application.EnableEvents = False

    r = 0
    For i = 1 To UBound(tableauFactures, 1)
        'Récupérer les données de la facture directement du Range
        numFacture = CStr(tableauFactures(i, fFacCCInvNo))
        'On traite seulement les transactions pour les factures confirmées
        If Not dictTypesFactures.Exists(numFacture) Or dictTypesFactures(numFacture) <> "C" Then
            GoTo Next_Invoice
        End If
        
        'Est-ce que la facture est à l'intérieur de la date limite ?
        dateFacture = tableauFactures(i, fFacCCInvoiceDate)
        If tableauFactures(i, fFacCCInvoiceDate) > CDate(dateLimite) Then
            Debug.Print "#022 - Comparaison de date - " & tableauFactures(i, fFacCCInvoiceDate) & " .vs. " & dateLimite
            GoTo Next_Invoice
        End If
        
        client = tableauFactures(i, fFacCCCodeClient)
        'Obtenir le nom du client (MF) pour trier par nom de client plutôt que par code de client
        If dictNomsClients.Exists(client) Then
            client = dictNomsClients(client)
        Else
            client = "Client inconnu"
        End If
        dateDue = CDate(tableauFactures(i, fFacCCDueDate))
        montantFacture = CCur(tableauFactures(i, fFacCCTotal))
        
        'Obtenir les paiements et régularisations pour cette facture
        montantPaye = 0
        Dim Pmt As Variant
        If dictPaiements.Exists(numFacture) Then
            For Each Pmt In dictPaiements(numFacture)
                If Pmt(0) <= dateLimite Then montantPaye = montantPaye + Pmt(1)
            Next
        End If
        montantRegul = 0
        Dim Reg As Variant
        If dictRegul.Exists(numFacture) Then
            For Each Reg In dictRegul(numFacture)
                If Reg(0) <= dateLimite Then montantRegul = montantRegul + Reg(1)
            Next
        End If
        
        montantRestant = montantFacture - montantPaye + montantRegul
        
        'Exclus les soldes de facture à 0,00 $ SI ET SEULMENT SI F4 = "NON"
        If UCase$(inclureSoldesNuls) = False And montantRestant = 0 Then
            GoTo Next_Invoice
        End If
        
        'Calcul de l'âge de la facture
        ageFacture = WorksheetFunction.Max(dateLimite - dateDue, 0)
        
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
        
        Dim tableau As Variant
        'Ajouter les données au dictionnaire en fonction du niveau de détail
        Select Case LCase$(niveauDetail)
            Case "client"
                If Not dictSoldesClients.Exists(client) Then
                    dictSoldesClients.Add client, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
                End If
                tableau = dictSoldesClients(client) 'Obtenir le tableau a partir du dictionary
                
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
                dictSoldesClients(client) = tableau ' Replacer le tableau dans le dictionnaire
            
            Case "facture"
                'Ajouter chaque facture avec son montant restant dû
                r = r + 1
                buffer(r, 1) = client
                buffer(r, 2) = numFacture
                buffer(r, 3) = dateFacture
                buffer(r, 4) = montantRestant
                Select Case trancheAge
                    Case "- de 30 jours"
                        buffer(r, 5) = montantRestant
                    Case "31 @ 60 jours"
                        buffer(r, 6) = montantRestant
                    Case "61 @ 90 jours"
                        buffer(r, 7) = montantRestant
                    Case "+ de 90 jours"
                        buffer(r, 8) = montantRestant
                End Select
                
            Case "transaction"
                'La facture en premier...
                r = r + 1
                buffer(r, 1) = client
                buffer(r, 2) = numFacture
                buffer(r, 3) = "Facture"
                buffer(r, 4) = dateFacture
                buffer(r, 5) = montantFacture
                Select Case trancheAge
                    Case "- de 30 jours"
                        buffer(r, 6) = montantRestant
                    Case "31 @ 60 jours"
                        buffer(r, 7) = montantRestant
                    Case "61 @ 90 jours"
                        buffer(r, 8) = montantRestant
                    Case "+ de 90 jours"
                        buffer(r, 9) = montantRestant
                End Select
                
                'Transactions de paiements par la suite
                If dictPaiements.Exists(numFacture) Then
                    For Each Pmt In dictPaiements(numFacture)
                        If Pmt(0) <= dateLimite Then
                            r = r + 1
                            buffer(r, 1) = client
                            buffer(r, 2) = numFacture
                            buffer(r, 3) = "Paiement"
                            buffer(r, 4) = Pmt(0)
                            buffer(r, 5) = -Pmt(1)
                        End If
                    Next
                End If
                'Transactions de régularisations par la suite
                If dictRegul.Exists(numFacture) Then
                    For Each Reg In dictRegul(numFacture)
                        If Reg(0) <= dateLimite Then
                            r = r + 1
                            buffer(r, 1) = client
                            buffer(r, 2) = numFacture
                            buffer(r, 3) = "Régularisation"
                            buffer(r, 4) = Reg(0)
                            buffer(r, 5) = Reg(1)
                        End If
                    Next
                End If
        End Select

Next_Invoice:
    Next i
    
    Application.EnableEvents = True
    
    'Si niveau de détail est "client", ajouter les soldes du client (dictionary) au tableau final
    If LCase$(niveauDetail) = "client" Then
        r = 0
        Dim cle As Variant
        
        Application.EnableEvents = False
        
        For Each cle In dictSoldesClients.keys
            r = r + 1
            buffer(r, 1) = cle 'Nom du client
            buffer(r, 2) = dictSoldesClients(cle)(0) 'Total
            buffer(r, 3) = dictSoldesClients(cle)(1) '- de 30 jours
            buffer(r, 4) = dictSoldesClients(cle)(2) '31 @ 60 jours
            buffer(r, 5) = dictSoldesClients(cle)(3) '61 @ 90 jours
            buffer(r, 6) = dictSoldesClients(cle)(4) '+ de 90 jours
        Next cle
        
        Application.EnableEvents = True

    End If
    
    'Tri alphabétique par nom de client
    wshCAR_Liste_Agee.Range("B9").Resize(UBound(buffer, 1), UBound(buffer, 2)).Value = buffer
    
    DerniereLigne = wshCAR_Liste_Agee.Cells(wshCAR_Liste_Agee.Rows.count, "B").End(xlUp).Row
    Dim rngResultat As Range
    Set rngResultat = wshCAR_Liste_Agee.Range("B8:J" & DerniereLigne)

    Application.EnableEvents = False
    
    Dim ordreTri As String
    If Not rngResultat Is Nothing Then
        If rngResultat.Rows.count > 1 Then 'Au moins 1 ligne de données
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
    
'    wshCAR_Liste_Agee.Range("B9").Resize(r - 8, nbColonnes).Value = buffer
    
    Call GererBoutonsNavigation(True)
    
    Dim ligneVolet As Long
    ligneVolet = ActiveWindow.SplitRow
    Dim saveDerniereLigne As Long
    saveDerniereLigne = DerniereLigne
    
    If DerniereLigne > (20 + ligneVolet) Then
        DerniereLigne = DerniereLigne - 20
        Application.Goto wshCAR_Liste_Agee.Cells(DerniereLigne, 1), Scroll:=True '2025-07-08 @ 13:39
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

    Dim header1 As String: header1 = "Liste âgée des comptes clients au " & dateLimite
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
    Set dictNomsClients = Nothing
    Set dictPaiements = Nothing
    Set dictRegul = Nothing
    Set dictSoldesClients = Nothing
    Set rngResultat = Nothing
    Set rngToPrint = Nothing
    Set wsFacturesCC = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modCAR_Liste_Agee:CreerListeAgeeCC", vbNullString, startTime)
    
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
        menuItem.caption = "Visualiser la facture (format PDF)"
        menuItem.OnAction = "'modFAC_Interrogation.VisualiserFacturePDF """ & numeroFacture & """'"

    'Option # 2 - Visualiser la facture
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.caption = "Envoi d'un rappel par courriel"
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
    allCols = modFunctions.Fn_ObtenirLigneDeFeuille("FAC_Entete", noFact, fFacEInvNo)
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
    allCols = modFunctions.Fn_ObtenirLigneDeFeuille("BD_Clients", codeClient, fClntFMClientID)
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
    allCols = modFunctions.Fn_ObtenirLigneDeFeuille("FAC_Comptes_Clients", noFact, fFacCCInvNo)
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
    attachmentFullPathName = wsdADMIN.Range("PATH_DATA_FILES").Value & gFACT_PDF_PATH & Application.PathSeparator & _
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
    
    ws.Protect UserInterfaceOnly:=True

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
        
        Application.Goto Reference:=.Cells(ligneCible, 1), Scroll:=True
    End With
    
End Sub

Sub shpVersHaut_Click() '2025-06-30 @ 10:59

    Call RetournerEnHaut

End Sub

Sub RetournerEnHaut() '2025-06-30 @ 11:08

    Application.Goto Reference:=ActiveSheet.Cells(1, 1), Scroll:=True
    
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

Sub shpRetournerAuMenu_Click()

    Call EffacerResultatAnterieur(wshCAR_Liste_Agee)
    
    Call RetournerMenuFacturation

End Sub

Sub RetournerMenuFacturation()

    Call modAppli.QuitterFeuillePourMenu(wshMenuFAC, True) '2025-08-20 @ 07:17

End Sub


