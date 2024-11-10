Attribute VB_Name = "modCAR_Liste_Agée"
Option Explicit

Sub CAR_Creer_Liste_Agee() '2024-09-08 @ 15:55

    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Liste_Agée:CAR_Creer_Liste_Agee", 0)
   
    'Initialiser les feuilles nécessaires
    Dim wsFactures As Worksheet: Set wsFactures = ThisWorkbook.Sheets("FAC_Comptes_Clients")
    Dim wsPaiements As Worksheet: Set wsPaiements = ThisWorkbook.Sheets("ENC_Détails")
    
    'Utilisation de la même feuille
    Dim rngResultat As Range
    Set rngResultat = wshCAR_Liste_Agée.Range("B8")
    Dim lastUsedRow As Long
    lastUsedRow = wshCAR_Liste_Agée.Cells(wshCAR_Liste_Agée.rows.count, "B").End(xlUp).row
    If lastUsedRow > 7 Then
        Application.EnableEvents = False
        wshCAR_Liste_Agée.Range("B8:J" & lastUsedRow + 5).Clear
        Application.EnableEvents = True
    End If
    
    'Niveau de détail
    Dim niveauDetail As String
    niveauDetail = wshCAR_Liste_Agée.Range("B4").value
    
    Application.EnableEvents = False
    
    'Entêtes de colonnes en fonction du niveau de détail
    If LCase(niveauDetail) = "client" Then
        wshCAR_Liste_Agée.Range("B8:G8").value = Array("Client", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Agée.Range("B8:G8"))
    End If

    'Entêtes de colonnes en fonction du niveau de détail (Facture)
    If LCase(niveauDetail) = "facture" Then
        wshCAR_Liste_Agée.Range("B8:I8").value = Array("Client", "No. Facture", "Date Facture", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Agée.Range("B8:I8"))
    End If

    'Entêtes de colonnes en fonction du niveau de détail (Transaction)
    If LCase(niveauDetail) = "transaction" Then
        wshCAR_Liste_Agée.Range("B8:J8").value = Array("Client", "No. Facture", "Type", "Date", "Montant", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Agée.Range("B8:J8"))
    End If

    Application.EnableEvents = True
    
    'Initialiser le dictionnaire pour les résultats (Nom du client, Solde)
    Dim dictClients As Object ' Utilisez un dictionnaire pour stocker les résultats
    Set dictClients = CreateObject("Scripting.Dictionary")
    
    'Date actuelle pour le calcul de l'âge des factures
    Dim dateAujourdhui As Date
    dateAujourdhui = Date
    
    'Boucle sur les factures
    Dim derniereLigne As Long
    derniereLigne = wsFactures.Cells(wsFactures.rows.count, "A").End(xlUp).row
    Dim rngFactures As Range
    Set rngFactures = wsFactures.Range("A3:A" & derniereLigne) '2 lignes d'entête
    
    Dim client As String, numFacture As String
    Dim dateFacture As Date, dateDue As Date
    Dim montantFacture As Currency, montantPaye As Currency, montantRestant As Currency
    Dim trancheAge As String
    Dim ageFacture As Long, i As Long, r As Long
    
    Application.EnableEvents = False
    
    r = 8
    For i = 1 To rngFactures.rows.count
        'Récupérer les données de la facture directement du Range
        numFacture = CStr(rngFactures.Cells(i, 1).value)
        'Do not process non Confirmed invoice
        If Fn_Get_Invoice_Type(numFacture) <> "C" Then
            GoTo Next_Invoice
        End If
        'Est-ce que la facture est à l'intérieur e la date limite ?
        dateFacture = CDate(rngFactures.Cells(i, 2).value)
        If rngFactures.Cells(i, 2).value > CDate(wshCAR_Liste_Agée.Range("H4").value) Then
            Debug.Print "#0081 - Comparaison de date - " & rngFactures.Cells(i, 2).value & " .vs. " & wshCAR_Liste_Agée.Range("H4").value
            GoTo Next_Invoice
        End If
        
        client = rngFactures.Cells(i, 4).value
        'Obtenir le nom du client (MF) pour trier par nom de client plutôt que par code de client
        client = Fn_Get_Client_Name(client)
        dateDue = CDate(rngFactures.Cells(i, 7).value)
        montantFacture = CCur(rngFactures.Cells(i, 8).value)
        
        'Obtenir les paiemnets pour cette facture
        montantPaye = CCur(Application.WorksheetFunction.SumIf(wsPaiements.Range("B:B"), numFacture, wsPaiements.Range("E:E")))
        montantRestant = montantFacture - montantPaye
        
        'Exclus les soldes de facture à 0,00 $ SI ET SEULMENT SI F4 = "NON"
        If UCase(wshCAR_Liste_Agée.Range("F4").value) = "NON" And montantRestant = 0 Then
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
        
        Dim ngPaiements As Range
        Dim rowOffset As Long
        Dim tableau As Variant
        'Ajouter les données au dictionnaire en fonction du niveau de détail
        Select Case LCase(niveauDetail)
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
                wshCAR_Liste_Agée.Cells(r, 2).value = client
                wshCAR_Liste_Agée.Cells(r, 3).value = numFacture
                wshCAR_Liste_Agée.Cells(r, 4).value = Format$(dateFacture, wshAdmin.Range("B1").value)
                wshCAR_Liste_Agée.Cells(r, 5).value = montantRestant
                Select Case trancheAge
                    Case "- de 30 jours"
                        wshCAR_Liste_Agée.Cells(r, 6).value = montantRestant
                    Case "31 @ 60 jours"
                        wshCAR_Liste_Agée.Cells(r, 7).value = montantRestant
                    Case "61 @ 90 jours"
                        wshCAR_Liste_Agée.Cells(r, 8).value = montantRestant
                    Case "+ de 90 jours"
                        wshCAR_Liste_Agée.Cells(r, 9).value = montantRestant
                End Select
                
            Case "transaction"
                'La facture en premier...
                r = r + 1
                wshCAR_Liste_Agée.Cells(r, 2).value = client
                wshCAR_Liste_Agée.Cells(r, 3).value = numFacture
                wshCAR_Liste_Agée.Cells(r, 4).value = "Facture"
                wshCAR_Liste_Agée.Cells(r, 5).value = Format$(dateFacture, wshAdmin.Range("B1").value)
                wshCAR_Liste_Agée.Cells(r, 6).value = montantFacture
                Select Case trancheAge
                    Case "- de 30 jours"
                        wshCAR_Liste_Agée.Cells(r, 7).value = montantRestant
                    Case "31 @ 60 jours"
                        wshCAR_Liste_Agée.Cells(r, 8).value = montantRestant
                    Case "61 @ 90 jours"
                        wshCAR_Liste_Agée.Cells(r, 9).value = montantRestant
                    Case "+ de 90 jours"
                        wshCAR_Liste_Agée.Cells(r, 10).value = montantRestant
                End Select
                
                'Transactions de paiements par la suite
                Dim rngPaiementsAssoc As Range
                Dim firstAddress As String
                'Obtenir tous les paiements pour la facture
                Set rngPaiementsAssoc = wsPaiements.Range("B:B").Find(numFacture, LookIn:=xlValues, LookAt:=xlWhole)
                If Not rngPaiementsAssoc Is Nothing Then
                    firstAddress = rngPaiementsAssoc.Address
                        Do
                        r = r + 1
                        wshCAR_Liste_Agée.Cells(r, 2).value = client
                        wshCAR_Liste_Agée.Cells(r, 3).value = numFacture
                        wshCAR_Liste_Agée.Cells(r, 4).value = "Paiement"
                        wshCAR_Liste_Agée.Cells(r, 5).value = rngPaiementsAssoc.Offset(0, 2).value
                        wshCAR_Liste_Agée.Cells(r, 6).value = -rngPaiementsAssoc.Offset(0, 3).value ' Montant du paiement
                        Set rngPaiementsAssoc = wsPaiements.columns("B:B").FindNext(rngPaiementsAssoc)
                    Loop While Not rngPaiementsAssoc Is Nothing And rngPaiementsAssoc.Address <> firstAddress
                End If
        End Select

Next_Invoice:
    Next i
    
    Application.EnableEvents = True
    
    'Si niveau de détail est "client", ajouter les soldes du client (dictionary) au tableau final
    If LCase(niveauDetail) = "client" Then
        r = 8
        Dim cle As Variant
        
        Application.EnableEvents = False
        
        For Each cle In dictClients.keys
            r = r + 1
            wshCAR_Liste_Agée.Cells(r, 2).value = cle ' Nom du client
            wshCAR_Liste_Agée.Cells(r, 3).value = dictClients(cle)(0) ' Total
            wshCAR_Liste_Agée.Cells(r, 4).value = dictClients(cle)(1) ' - de 30 jours
            wshCAR_Liste_Agée.Cells(r, 5).value = dictClients(cle)(2) ' 31 @ 60 jours
            wshCAR_Liste_Agée.Cells(r, 6).value = dictClients(cle)(3) ' 61 @ 90 jours
            wshCAR_Liste_Agée.Cells(r, 7).value = dictClients(cle)(4) ' + de 90 jours
        Next cle
        
        Application.EnableEvents = True

    End If
    
    'Tri alphabétique par nom de client
    derniereLigne = wshCAR_Liste_Agée.Cells(wshCAR_Liste_Agée.rows.count, "B").End(xlUp).row
    Set rngResultat = wshCAR_Liste_Agée.Range("B8:J" & derniereLigne)
    
    Application.EnableEvents = False
    
    Dim ordreTri As String
    If derniereLigne > 9 Then 'Le tri n'est peut-être pas nécessaire
        With wshCAR_Liste_Agée.Sort
            .SortFields.Clear
            If wshCAR_Liste_Agée.Range("D4").value = "Nom de client" Then
                .SortFields.Add _
                    key:=wshCAR_Liste_Agée.Range("B8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par nom de client
                ordreTri = "Ordre de nom de client"
            Else
                .SortFields.Add _
                    key:=wshCAR_Liste_Agée.Range("C8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par numéro de facture
                .SortFields.Add _
                    key:=wshCAR_Liste_Agée.Range("D8"), _
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
    
    derniereLigne = derniereLigne + 2
    
    Dim t(0 To 4) As Currency
    
    With wshCAR_Liste_Agée
        .columns("B:B").ColumnWidth = 50
        .columns("C:J").ColumnWidth = 13
        Select Case LCase(niveauDetail)
            Case "client"
                .Range("C9:G" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("C9:G" & derniereLigne).HorizontalAlignment = xlRight
                .Range("C" & derniereLigne).formula = "=Sum(C9:C" & derniereLigne - 2 & ")"
                t(0) = .Range("C" & derniereLigne).value
                .Range("D" & derniereLigne).formula = "=Sum(D9:D" & derniereLigne - 2 & ")"
                t(1) = .Range("D" & derniereLigne).value
                .Range("E" & derniereLigne).formula = "=Sum(E9:E" & derniereLigne - 2 & ")"
                t(2) = .Range("E" & derniereLigne).value
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                t(3) = .Range("F" & derniereLigne).value
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                t(4) = .Range("G" & derniereLigne).value
                .Range("C" & derniereLigne & ":G" & derniereLigne).Font.Bold = True
            Case "facture"
                .Range("C9:C" & derniereLigne).HorizontalAlignment = xlCenter
                .Range("D9:D" & derniereLigne).HorizontalAlignment = xlCenter
                .Range("E9:I" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("E9:I" & derniereLigne).HorizontalAlignment = xlRight
                .Range("E" & derniereLigne).formula = "=Sum(E9:E" & derniereLigne - 2 & ")"
                t(0) = .Range("E" & derniereLigne).value
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                t(1) = .Range("F" & derniereLigne).value
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                t(2) = .Range("G" & derniereLigne).value
                .Range("H" & derniereLigne).formula = "=Sum(H9:H" & derniereLigne - 2 & ")"
                t(3) = .Range("H" & derniereLigne).value
                .Range("I" & derniereLigne).formula = "=Sum(I9:I" & derniereLigne - 2 & ")"
                t(4) = .Range("I" & derniereLigne).value
                .Range("E" & derniereLigne & ":I" & derniereLigne).Font.Bold = True
            Case "transaction"
                .columns("C:E").HorizontalAlignment = xlCenter
                .columns("D").HorizontalAlignment = xlLeft
                .Range("F9:J" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("F9:J" & derniereLigne).HorizontalAlignment = xlRight
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                t(0) = .Range("F" & derniereLigne).value
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                t(1) = .Range("G" & derniereLigne).value
                .Range("H" & derniereLigne).formula = "=Sum(H9:H" & derniereLigne - 2 & ")"
                t(2) = .Range("H" & derniereLigne).value
                .Range("I" & derniereLigne).formula = "=Sum(I9:I" & derniereLigne - 2 & ")"
                t(3) = .Range("I" & derniereLigne).value
                .Range("J" & derniereLigne).formula = "=Sum(J9:J" & derniereLigne - 2 & ")"
                t(4) = .Range("J" & derniereLigne).value
                .Range("F" & derniereLigne & ":J" & derniereLigne).Font.Bold = True
        End Select
        .Range("B" & derniereLigne).value = "Totaux de la liste"
        .Range("B" & derniereLigne).Font.Bold = True
        derniereLigne = derniereLigne + 1
        
        'Ligne de pourcentages
        .Range("B" & derniereLigne).value = "Pourcentages"
        .Range("B" & derniereLigne & ":J" & derniereLigne).Font.Bold = True
        .Range("C" & derniereLigne & ":J" & derniereLigne).NumberFormat = "##0.00"
        .Range("C" & derniereLigne & ":J" & derniereLigne).HorizontalAlignment = xlRight
        Dim totalListe As Currency
        totalListe = t(0)
        If totalListe <> 0 Then
            Select Case LCase(niveauDetail)
                Case "client"
                    .Range("C" & derniereLigne).value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("D" & derniereLigne).value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("E" & derniereLigne).value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("F" & derniereLigne).value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("G" & derniereLigne).value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
                Case "facture"
                    .Range("E" & derniereLigne).value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("F" & derniereLigne).value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("G" & derniereLigne).value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("H" & derniereLigne).value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("I" & derniereLigne).value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
                Case "transaction"
                    .Range("F" & derniereLigne).value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("G" & derniereLigne).value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("H" & derniereLigne).value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("I" & derniereLigne).value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("J" & derniereLigne).value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
            End Select
        End If
    End With
    
    Application.EnableEvents = True

    DoEvents
    
    'Result print setup - 2024-08-31 @ 12:19
    lastUsedRow = derniereLigne
    
    Dim rngToPrint As Range:
    Select Case LCase(niveauDetail)
        Case "client"
            Set rngToPrint = wshCAR_Liste_Agée.Range("B9:G" & lastUsedRow)
        Case "facture"
            Set rngToPrint = wshCAR_Liste_Agée.Range("B9:I" & lastUsedRow)
        Case "transaction"
            Set rngToPrint = wshCAR_Liste_Agée.Range("B9:J" & lastUsedRow)
    End Select
    
    Application.EnableEvents = False

    Call Apply_Conditional_Formatting_Alternate(rngToPrint, 0, False)
    
    'Caractères pour le rapport
    With rngToPrint.Font
        .Name = "Aptos Narrow"
        .size = 10
    End With
    
    Application.EnableEvents = True
    
    DoEvents

    Dim header1 As String: header1 = "Liste âgée des comptes clients au " & wshCAR_Liste_Agée.Range("H4").value
    Dim header2 As String
    If LCase(niveauDetail) = "client" Then
        header2 = "1 ligne par client"
    ElseIf LCase(niveauDetail) = "facture" Then
        header2 = "1 ligne par Facture"
    Else
        header2 = "1 ligne par transaction"
    End If
    header2 = ordreTri & " - " & header2
    
    Call Simple_Print_Setup(wshCAR_Liste_Agée, rngToPrint, header1, header2, "$8:$8", "L")
    
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
    
    Call Log_Record("modCAR_Liste_Agée:CAR_Creer_Liste_Agee", startTime)
    
End Sub

