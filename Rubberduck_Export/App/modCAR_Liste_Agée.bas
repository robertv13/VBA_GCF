Attribute VB_Name = "modCAR_Liste_Ag�e"
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
Option Explicit

Sub CC_PreparerListeAgee_Click()

    Call CreerListeAgee

End Sub

Sub CreerListeAgee() '2024-09-08 @ 15:55

    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Liste_Ag�e:CreerListeAgee", "", 0)
   
    Application.ScreenUpdating = False
    
    'D�bloque la feuille
    ActiveSheet.Unprotect
    
    'Initialiser les feuilles n�cessaires
    Dim wsFactures As Worksheet: Set wsFactures = wshFAC_Comptes_Clients
    Dim wsPaiements As Worksheet: Set wsPaiements = wshENC_D�tails
    Dim wsR�gularisations As Worksheet: Set wsR�gularisations = wshCC_R�gularisations
    
    'Utilisation de la m�me feuille
    Dim rngResultat As Range
    Set rngResultat = wshCAR_Liste_Ag�e.Range("B8")
    Dim lastUsedRow As Long
    lastUsedRow = wshCAR_Liste_Ag�e.Cells(wshCAR_Liste_Ag�e.Rows.count, "B").End(xlUp).row
    If lastUsedRow > 7 Then
        Application.EnableEvents = False
        wshCAR_Liste_Ag�e.Unprotect
        wshCAR_Liste_Ag�e.Range("B8:J" & lastUsedRow + 5).Clear
        wshCAR_Liste_Ag�e.Protect UserInterfaceOnly:=True
        Application.EnableEvents = True
    End If
    
    'Niveau de d�tail
    Dim niveauDetail As String
    niveauDetail = wshCAR_Liste_Ag�e.Range("B4").value
    
    Application.EnableEvents = False
    
    'Ent�tes de colonnes en fonction du niveau de d�tail
    If LCase$(niveauDetail) = "client" Then
        wshCAR_Liste_Ag�e.Range("B8:G8").value = Array("Client", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Ag�e.Range("B8:G8"))
    End If

    'Ent�tes de colonnes en fonction du niveau de d�tail (Facture)
    If LCase$(niveauDetail) = "facture" Then
        wshCAR_Liste_Ag�e.Range("B8:I8").value = Array("Client", "No. Facture", "Date Facture", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Ag�e.Range("B8:I8"))
    End If

    'Ent�tes de colonnes en fonction du niveau de d�tail (Transaction)
    If LCase$(niveauDetail) = "transaction" Then
        wshCAR_Liste_Ag�e.Range("B8:J8").value = Array("Client", "No. Facture", "Type", "Date", "Montant", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Ag�e.Range("B8:J8"))
    End If

    Application.EnableEvents = True

    'Initialiser le dictionnaire pour les r�sultats (Nom du client, Solde)
    Dim dictClients As Object 'Utilisez un dictionnaire pour stocker les r�sultats
    Set dictClients = CreateObject("Scripting.Dictionary")
    
    'Date actuelle pour le calcul de l'�ge des factures
    Dim dateAujourdhui As Date
    dateAujourdhui = Date
    
    'Boucle sur les factures
    Dim DerniereLigne As Long
    DerniereLigne = wsFactures.Cells(wsFactures.Rows.count, 1).End(xlUp).row
    Dim rngFactures As Range
    Set rngFactures = wsFactures.Range("A3:A" & DerniereLigne) '2 lignes d'ent�te
    
    Dim client As String, numFacture As String
    Dim dateFacture As Date, dateDue As Date
    Dim montantFacture As Currency, montantPaye As Currency, montantRegul As Currency, montantRestant As Currency
    Dim trancheAge As String
    Dim ageFacture As Long, i As Long, r As Long
    
    Application.EnableEvents = False

    r = 8
    For i = 1 To rngFactures.Rows.count
        'R�cup�rer les donn�es de la facture directement du Range
        numFacture = CStr(rngFactures.Cells(i, fFacCCInvNo).value)
        'Do not process non Confirmed invoice
        If Fn_Get_Invoice_Type(numFacture) <> "C" Then
            GoTo Next_Invoice
        End If
        
        'Est-ce que la facture est � l'int�rieur de la date limite ?
        dateFacture = rngFactures.Cells(i, fFacCCInvoiceDate).value
        If rngFactures.Cells(i, fFacCCInvoiceDate).value > CDate(wshCAR_Liste_Ag�e.Range("H4").value) Then
            Debug.Print "#022 - Comparaison de date - " & rngFactures.Cells(i, fFacCCInvoiceDate).value & " .vs. " & wshCAR_Liste_Ag�e.Range("H4").value
            GoTo Next_Invoice
        End If
        
        client = rngFactures.Cells(i, fFacCCCodeClient).value
        'Obtenir le nom du client (MF) pour trier par nom de client plut�t que par code de client
        client = Fn_Get_Client_Name(client)
        dateDue = rngFactures.Cells(i, fFacCCDueDate).value
        montantFacture = CCur(rngFactures.Cells(i, fFacCCTotal).value)
        
        'Obtenir les paiements et r�gularisations pour cette facture
        montantPaye = Fn_Obtenir_Paiements_Facture(numFacture, wshCAR_Liste_Ag�e.Range("H4").value)
        montantRegul = Fn_Obtenir_R�gularisations_Facture(numFacture, wshCAR_Liste_Ag�e.Range("H4").value)
        
        montantRestant = montantFacture - montantPaye + montantRegul
        
        'Exclus les soldes de facture � 0,00 $ SI ET SEULMENT SI F4 = "NON"
        If UCase$(wshCAR_Liste_Ag�e.Range("F4").value) = "NON" And montantRestant = 0 Then
            GoTo Next_Invoice
        End If
        
        'Calcul de l'�ge de la facture
        ageFacture = WorksheetFunction.Max(dateAujourdhui - dateDue, 0)
        
        'D�termine la trancheAge d'�ge
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
                trancheAge = "Non d�fini"
        End Select
        
        Dim rngPaiements As Range
        Dim RowOffset As Long
        Dim tableau As Variant
        'Ajouter les donn�es au dictionnaire en fonction du niveau de d�tail
        Select Case LCase$(niveauDetail)
            Case "client"
                If Not dictClients.Exists(client) Then
                    dictClients.Add client, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
                End If
                tableau = dictClients(client) 'Obtenir le tableau a partir du dictionary
                
                'Ajouter le solde de la facture au total (0)
                tableau(0) = tableau(0) + montantRestant
                
                'Ajouter le montant restant � la trancheAge correspondante (1 @ 4)
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
                'Ajouter chaque facture avec son montant restant d�
                r = r + 1
                wshCAR_Liste_Ag�e.Cells(r, 2).value = client
                wshCAR_Liste_Ag�e.Cells(r, 3).value = numFacture
                wshCAR_Liste_Ag�e.Cells(r, 4).value = dateFacture
                wshCAR_Liste_Ag�e.Cells(r, 4).NumberFormat = wshAdmin.Range("B1").value
                wshCAR_Liste_Ag�e.Cells(r, 5).value = montantRestant
                Select Case trancheAge
                    Case "- de 30 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 6).value = montantRestant
                    Case "31 @ 60 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 7).value = montantRestant
                    Case "61 @ 90 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 8).value = montantRestant
                    Case "+ de 90 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 9).value = montantRestant
                End Select
                
            Case "transaction"
                'La facture en premier...
                r = r + 1
                wshCAR_Liste_Ag�e.Cells(r, 2).value = client
                wshCAR_Liste_Ag�e.Cells(r, 3).value = numFacture
                wshCAR_Liste_Ag�e.Cells(r, 4).value = "Facture"
                wshCAR_Liste_Ag�e.Cells(r, 5).value = dateFacture
                wshCAR_Liste_Ag�e.Cells(r, 5).NumberFormat = wshAdmin.Range("B1").value
                wshCAR_Liste_Ag�e.Cells(r, 6).value = montantFacture
                Select Case trancheAge
                    Case "- de 30 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 7).value = montantRestant
                    Case "31 @ 60 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 8).value = montantRestant
                    Case "61 @ 90 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 9).value = montantRestant
                    Case "+ de 90 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 10).value = montantRestant
                End Select
                
                'Transactions de paiements par la suite
                Dim rngPaiementsAssoc As Range
                Dim pmtFirstAddress As String
                'Obtenir tous les paiements pour la facture
                Set rngPaiementsAssoc = wsPaiements.Range("B:B").Find(numFacture, LookIn:=xlValues, LookAt:=xlWhole)
                If Not rngPaiementsAssoc Is Nothing Then
                    pmtFirstAddress = rngPaiementsAssoc.Address
                    Do
                        If rngPaiementsAssoc.offset(0, 2).value <= CDate(wshCAR_Liste_Ag�e.Range("H4").value) Then
                            r = r + 1
                            wshCAR_Liste_Ag�e.Cells(r, 2).value = client
                            wshCAR_Liste_Ag�e.Cells(r, 3).value = numFacture
                            wshCAR_Liste_Ag�e.Cells(r, 4).value = "Paiement"
                            wshCAR_Liste_Ag�e.Cells(r, 5).value = rngPaiementsAssoc.offset(0, 2).value
                            wshCAR_Liste_Ag�e.Cells(r, 6).value = -rngPaiementsAssoc.offset(0, 3).value 'Montant du paiement
                        End If
                        Set rngPaiementsAssoc = wsPaiements.Columns("B:B").FindNext(rngPaiementsAssoc)
                    Loop While Not rngPaiementsAssoc Is Nothing And rngPaiementsAssoc.Address <> pmtFirstAddress
                End If
                'Transactions de r�gularisations par la suite
                Dim rngR�gularisationAssoc As Range
                Dim regulFirstAddress As String
                'Obtenir toutes les r�gularisations pour la facture
                Set rngR�gularisationAssoc = wsR�gularisations.Range("B:B").Find(numFacture, LookIn:=xlValues, LookAt:=xlWhole)
                If Not rngR�gularisationAssoc Is Nothing Then
                    regulFirstAddress = rngR�gularisationAssoc.Address
                    Do
                        If rngR�gularisationAssoc.offset(0, 1).value <= CDate(wshCAR_Liste_Ag�e.Range("H4").value) Then
                            r = r + 1
                            wshCAR_Liste_Ag�e.Cells(r, 2).value = client
                            wshCAR_Liste_Ag�e.Cells(r, 3).value = numFacture
                            wshCAR_Liste_Ag�e.Cells(r, 4).value = "R�gularisation"
                            wshCAR_Liste_Ag�e.Cells(r, 5).value = rngR�gularisationAssoc.offset(0, 1).value
                            wshCAR_Liste_Ag�e.Cells(r, 6).value = rngR�gularisationAssoc.offset(0, 4).value + _
                                rngR�gularisationAssoc.offset(0, 5).value + _
                                rngR�gularisationAssoc.offset(0, 6).value + _
                                rngR�gularisationAssoc.offset(0, 7).value
                        End If
                        Set rngR�gularisationAssoc = wsR�gularisations.Columns("B:B").FindNext(rngR�gularisationAssoc)
                    Loop While Not rngR�gularisationAssoc Is Nothing And rngR�gularisationAssoc.Address <> regulFirstAddress
                End If
        End Select

Next_Invoice:
    Next i
    
    Application.EnableEvents = True
    
    'Si niveau de d�tail est "client", ajouter les soldes du client (dictionary) au tableau final
    If LCase$(niveauDetail) = "client" Then
        r = 8
        Dim cle As Variant
        
        Application.EnableEvents = False
        
        For Each cle In dictClients.keys
            r = r + 1
            wshCAR_Liste_Ag�e.Cells(r, 2).value = cle ' Nom du client
            wshCAR_Liste_Ag�e.Cells(r, 3).value = dictClients(cle)(0) ' Total
            wshCAR_Liste_Ag�e.Cells(r, 4).value = dictClients(cle)(1) ' - de 30 jours
            wshCAR_Liste_Ag�e.Cells(r, 5).value = dictClients(cle)(2) ' 31 @ 60 jours
            wshCAR_Liste_Ag�e.Cells(r, 6).value = dictClients(cle)(3) ' 61 @ 90 jours
            wshCAR_Liste_Ag�e.Cells(r, 7).value = dictClients(cle)(4) ' + de 90 jours
        Next cle
        
        Application.EnableEvents = True

    End If
    
    'Tri alphab�tique par nom de client
    DerniereLigne = wshCAR_Liste_Ag�e.Cells(wshCAR_Liste_Ag�e.Rows.count, "B").End(xlUp).row
    Set rngResultat = wshCAR_Liste_Ag�e.Range("B8:J" & DerniereLigne)
    
    Application.EnableEvents = False
    
    Dim ordreTri As String
    If DerniereLigne > 9 Then 'Le tri n'est peut-�tre pas n�cessaire
        With wshCAR_Liste_Ag�e.Sort
            .SortFields.Clear
            If wshCAR_Liste_Ag�e.Range("D4").value = "Nom de client" Then
                .SortFields.Add _
                    key:=wshCAR_Liste_Ag�e.Range("B8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par nom de client
                ordreTri = "Ordre de nom de client"
            Else
                .SortFields.Add _
                    key:=wshCAR_Liste_Ag�e.Range("C8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par num�ro de facture
                .SortFields.Add _
                    key:=wshCAR_Liste_Ag�e.Range("D8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier date de transaction
                ordreTri = "Ordre de num�ro de facture"
            End If
            .SetRange rngResultat
            .Header = xlYes
            .Apply
        End With
    End If
    
    DerniereLigne = DerniereLigne + 2
    
    Dim t(0 To 4) As Currency
    
    With wshCAR_Liste_Ag�e
        .Columns("B:B").ColumnWidth = 50
        .Columns("C:J").ColumnWidth = 13
        Select Case LCase$(niveauDetail)
            Case "client"
                .Range("C9:G" & DerniereLigne).NumberFormat = "#,##0.00 $"
                .Range("C9:G" & DerniereLigne).HorizontalAlignment = xlRight
                .Range("C" & DerniereLigne).formula = "=Sum(C9:C" & DerniereLigne - 2 & ")"
                t(0) = .Range("C" & DerniereLigne).value
                .Range("D" & DerniereLigne).formula = "=Sum(D9:D" & DerniereLigne - 2 & ")"
                t(1) = .Range("D" & DerniereLigne).value
                .Range("E" & DerniereLigne).formula = "=Sum(E9:E" & DerniereLigne - 2 & ")"
                t(2) = .Range("E" & DerniereLigne).value
                .Range("F" & DerniereLigne).formula = "=Sum(F9:F" & DerniereLigne - 2 & ")"
                t(3) = .Range("F" & DerniereLigne).value
                .Range("G" & DerniereLigne).formula = "=Sum(G9:G" & DerniereLigne - 2 & ")"
                t(4) = .Range("G" & DerniereLigne).value
                .Range("C" & DerniereLigne & ":G" & DerniereLigne).Font.Bold = True
            Case "facture"
                .Range("C9:C" & DerniereLigne).HorizontalAlignment = xlCenter
                .Range("D9:D" & DerniereLigne).HorizontalAlignment = xlCenter
                .Range("E9:I" & DerniereLigne).NumberFormat = "#,##0.00 $"
                .Range("E9:I" & DerniereLigne).HorizontalAlignment = xlRight
                .Range("E" & DerniereLigne).formula = "=Sum(E9:E" & DerniereLigne - 2 & ")"
                t(0) = .Range("E" & DerniereLigne).value
                .Range("F" & DerniereLigne).formula = "=Sum(F9:F" & DerniereLigne - 2 & ")"
                t(1) = .Range("F" & DerniereLigne).value
                .Range("G" & DerniereLigne).formula = "=Sum(G9:G" & DerniereLigne - 2 & ")"
                t(2) = .Range("G" & DerniereLigne).value
                .Range("H" & DerniereLigne).formula = "=Sum(H9:H" & DerniereLigne - 2 & ")"
                t(3) = .Range("H" & DerniereLigne).value
                .Range("I" & DerniereLigne).formula = "=Sum(I9:I" & DerniereLigne - 2 & ")"
                t(4) = .Range("I" & DerniereLigne).value
                .Range("E" & DerniereLigne & ":I" & DerniereLigne).Font.Bold = True
            Case "transaction"
                .Columns("C:E").HorizontalAlignment = xlCenter
                .Columns("D").HorizontalAlignment = xlLeft
                .Range("F9:J" & DerniereLigne).NumberFormat = "#,##0.00 $"
                .Range("F9:J" & DerniereLigne).HorizontalAlignment = xlRight
                .Range("F" & DerniereLigne).formula = "=Sum(F9:F" & DerniereLigne - 2 & ")"
                t(0) = .Range("F" & DerniereLigne).value
                .Range("G" & DerniereLigne).formula = "=Sum(G9:G" & DerniereLigne - 2 & ")"
                t(1) = .Range("G" & DerniereLigne).value
                .Range("H" & DerniereLigne).formula = "=Sum(H9:H" & DerniereLigne - 2 & ")"
                t(2) = .Range("H" & DerniereLigne).value
                .Range("I" & DerniereLigne).formula = "=Sum(I9:I" & DerniereLigne - 2 & ")"
                t(3) = .Range("I" & DerniereLigne).value
                .Range("J" & DerniereLigne).formula = "=Sum(J9:J" & DerniereLigne - 2 & ")"
                t(4) = .Range("J" & DerniereLigne).value
                .Range("F" & DerniereLigne & ":J" & DerniereLigne).Font.Bold = True
        End Select
        .Range("B" & DerniereLigne).value = "Totaux de la liste"
        .Range("B" & DerniereLigne).Font.Bold = True
        DerniereLigne = DerniereLigne + 1
        
        'Ligne de pourcentages
        .Range("B" & DerniereLigne).value = "Pourcentages"
        .Range("B" & DerniereLigne & ":J" & DerniereLigne).Font.Bold = True
        .Range("C" & DerniereLigne & ":J" & DerniereLigne).NumberFormat = "##0.00"
        .Range("C" & DerniereLigne & ":J" & DerniereLigne).HorizontalAlignment = xlRight
        Dim totalListe As Currency
        totalListe = t(0)
        If totalListe <> 0 Then
            Select Case LCase$(niveauDetail)
                Case "client"
                    .Range("C" & DerniereLigne).value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("D" & DerniereLigne).value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("E" & DerniereLigne).value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("F" & DerniereLigne).value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("G" & DerniereLigne).value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
                Case "facture"
                    .Range("E" & DerniereLigne).value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("F" & DerniereLigne).value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("G" & DerniereLigne).value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("H" & DerniereLigne).value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("I" & DerniereLigne).value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
                Case "transaction"
                    .Range("F" & DerniereLigne).value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("G" & DerniereLigne).value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("H" & DerniereLigne).value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("I" & DerniereLigne).value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("J" & DerniereLigne).value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
            End Select
        End If
    End With
    
    Application.EnableEvents = True

    DoEvents
    
    'Result print setup - 2024-08-31 @ 12:19
    lastUsedRow = DerniereLigne
    
    Dim rngToPrint As Range:
    Select Case LCase$(niveauDetail)
        Case "client"
            Set rngToPrint = wshCAR_Liste_Ag�e.Range("B9:G" & lastUsedRow)
        Case "facture"
            Set rngToPrint = wshCAR_Liste_Ag�e.Range("B9:I" & lastUsedRow)
        Case "transaction"
            Set rngToPrint = wshCAR_Liste_Ag�e.Range("B9:J" & lastUsedRow)
    End Select
    
    Application.EnableEvents = False

    Call modAppli_Utils.ApplyConditionalFormatting(rngToPrint, 0, False)
    
    'Caract�res pour le rapport
    With rngToPrint.Font
        .Name = "Aptos Narrow"
        .size = 10
    End With
    
    Application.EnableEvents = True
    
    DoEvents

    Dim header1 As String: header1 = "Liste �g�e des comptes clients au " & wshCAR_Liste_Ag�e.Range("H4").value
    Dim header2 As String
    If LCase$(niveauDetail) = "client" Then
        header2 = "1 ligne par client"
    ElseIf LCase$(niveauDetail) = "facture" Then
        header2 = "1 ligne par Facture"
    Else
        header2 = "1 ligne par transaction"
    End If
    header2 = ordreTri & " - " & header2
    
    Call Simple_Print_Setup(wshCAR_Liste_Ag�e, rngToPrint, header1, header2, "$8:$8", "L")
    
    Application.ScreenUpdating = True
    
    msgBox "La pr�paration de la liste �g�e est termin�e", vbInformation
    
    Application.EnableEvents = True
    
    'Lib�rer la m�moire
    Set cle = Nothing
    Set dictClients = Nothing
    Set rngFactures = Nothing
    Set rngPaiementsAssoc = Nothing
    Set rngResultat = Nothing
    Set rngToPrint = Nothing
    Set wsFactures = Nothing
    Set wsPaiements = Nothing
    
    Call Log_Record("modCAR_Liste_Ag�e:CreerListeAgee", "", startTime)
    
End Sub

Sub CAR_ListeAgee_AfficherMenuContextuel(ByVal target As Range) '2025-02-21 @ 19:10

    Dim menu As CommandBar
    Dim menuItem As CommandBarButton

    'Supprimer le menu contextuel personnalis� s'il existe d�j�
    On Error Resume Next
    Application.CommandBars("FactureMenu").Delete
    On Error GoTo 0
    
    'D�termine les coordonn�es de la colonne qui a �t� cliqu�e
    Dim numeroLigne As Long, numeroColonne As Long
    Call ExtraireLigneColonneCellule(target.Address, numeroLigne, numeroColonne)
    
    Dim numeroFacture As String
    numeroFacture = ActiveSheet.Cells(numeroLigne, "C").value
    If Trim$(numeroFacture) = "" Then
        Exit Sub
    End If
    
    'Cr�er un nouveau menu contextuel
    Set menu = Application.CommandBars.Add(Name:="FactureMenu", position:=msoBarPopup, Temporary:=True)

    'Option # 1 - Visualiser la facture
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.Caption = "Visualiser la facture (format PDF)"
        menuItem.OnAction = "'VisualiserFacturePDF """ & numeroFacture & """'"

    'Option # 2 - Visualiser la facture
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.Caption = "Envoi d'un rappel par courriel"
        menuItem.OnAction = "'EnvoyerRappelParCourriel """ & numeroFacture & """'"

    'Afficher le menu contextuel
    menu.ShowPopup

End Sub

Sub EnvoyerRappelParCourriel(noFact As String)

    Dim montantDu As String
    Dim message As String

    'Retrouver le code du client
    Dim codeClient As String
    Dim dateFact As Date
    Dim allCols As Variant
    allCols = Fn_Get_A_Row_From_A_Worksheet("FAC_Ent�te", noFact, fFacEInvNo)
    'V�rifier les r�sultats
    If IsArray(allCols) Then
        codeClient = allCols(fFacECustID)
        dateFact = allCols(fFacEDateFacture)
    Else
        msgBox "Enregistrement '" & noFact & "' non trouv�e !!!", vbCritical
        Exit Sub
    End If
    
    If codeClient = "" Then
        msgBox "Le code client pour cette facture est INVALIDE", vbCritical, "Information erron�e / manquante"
        Exit Sub
    End If
    
    'Retrouver le nom du client, le nom du contact et son adresse courriel
    Dim clientNom As String
    Dim clientContactFact As String
    Dim prenomContact As String
    Dim clientCourriel As String
    allCols = Fn_Get_A_Row_From_A_Worksheet("BD_Clients", codeClient, fClntFMClientID)
    'V�rifier les r�sultats
    If IsArray(allCols) Then
        clientNom = allCols(fClntFMClientNom)
        '�limine le ou les nom(s) de contact
        clientNom = Fn_Strip_Contact_From_Client_Name(clientNom)
        clientContactFact = Trim$(allCols(fClntFMContactFacturation)) + " "
        prenomContact = ""
        If InStr(clientContactFact, " ") <> 0 Then
            prenomContact = Left$(clientContactFact, InStr(clientContactFact, " ") - 1)
        End If
        '0 � 2 adresses courriel
        clientCourriel = allCols(fClntFMCourrielFacturation)
    Else
        msgBox "Enregistrement '" & codeClient & "' non trouv�e !!!", vbCritical
        Exit Sub
    End If
    
    'Retrouver le solde de la facture, le total des paiements & le total des r�gularisations
    Dim factSolde As Currency
    Dim factSommePmts As Currency
    Dim factSommeRegul As Currency
    allCols = Fn_Get_A_Row_From_A_Worksheet("FAC_Comptes_Clients", noFact, fFacCCInvNo)
    'V�rifier les r�sultats
    If IsArray(allCols) Then
        factSolde = allCols(fFacCCBalance)
        factSommePmts = allCols(fFacCCTotalPaid)
        factSommeRegul = allCols(fFacCCTotalRegul)
    Else
        msgBox "Enregistrement '" & noFact & "' non trouv�e !!!", vbCritical
        Exit Sub
    End If
    
    'V�rification pour �viter d'envoyer un rappel pour une facture � 0 $ ou cr�diteur
    If factSolde <= 0 Then
        msgBox "Il n'y a pas lieu d'envoyer un rappel" & vbNewLine & vbNewLine & _
                "pour cette facture. Solde � " & Format$(factSolde, "#,##0.00 $"), vbCritical, "Solde � 0,00 $ ou cr�diteur"
        Exit Sub
    End If
    
    'V�rifier si l'email est valide
    If Fn_ValiderCourriel(clientCourriel) = False Then
        msgBox "L'adresse courriel est vide OU invalide" & vbNewLine & vbNewLine & _
                "pour ce client.", vbExclamation, "Impossible d'envoyer un rappel (Adresse courriel invalide)"
        Exit Sub
    End If

    'Ajouter la copie de la facture (format PDF)
    Dim attachmentFullPathName As String
    attachmentFullPathName = wshAdmin.Range("F5").value & FACT_PDF_PATH & Application.PathSeparator & _
                     noFact & ".pdf"
    
    'V�rification de l'existence de la pi�ce jointe
    Dim fileExists As Boolean
    fileExists = Dir(attachmentFullPathName) <> ""
    If Not fileExists Then
        msgBox "La pi�ce jointe (Facture en format PDF) n'existe pas �" & vbNewLine & _
                    "l'emplacement sp�cifi�, soit " & attachmentFullPathName, vbCritical
        GoTo Exit_Sub
    End If
    
    'Chemin du template (.oft) de courriel
    Dim templateFullPathName As String
    templateFullPathName = Environ("appdata") & "\Microsoft\Templates\GCF_Rappel.oft"

    'V�rification de l'existence du template
    fileExists = Dir(templateFullPathName) <> ""
    If Not fileExists Then
        msgBox "Le gabarit 'GCF_Rappel.oft' est introuvable " & _
                    "� l'emplacement sp�cifi�, soit " & Environ("appdata") & "\Microsoft\Templates", _
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
    
    'V�rifier si Outlook est bien ouvert
    If OutlookApp Is Nothing Then
        msgBox "Impossible d'ouvrir Outlook. V�rifiez votre installation.", vbCritical
        Exit Sub
    End If

    Dim mailItem As Object
    Set mailItem = OutlookApp.CreateItemFromTemplate(templateFullPathName)
    On Error GoTo 0

    mailItem.Attachments.Add attachmentFullPathName

    'Ajuster les derniers param�tres du courriel
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
        
        .Subject = wshAdmin.Range("NomEntreprise") & " - Rappel pour facture impay�e - " & clientNom & " - Facture # " & noFact
        .Display  'Pour afficher le mail avant envoi (remplacez par .Send pour envoyer directement)
    End With

Exit_Sub:

    ' Nettoyer les objets
    Set mailItem = Nothing
    Set OutlookApp = Nothing
    
End Sub

Sub shpRetourMenuFacturation_Click()

    Call RetourMenuFacturation

End Sub

Sub RetourMenuFacturation()

    wshCAR_Liste_Ag�e.Visible = xlSheetHidden
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select

End Sub


