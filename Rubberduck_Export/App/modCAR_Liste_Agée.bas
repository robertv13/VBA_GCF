Attribute VB_Name = "modCAR_Liste_Ag�e"
Option Explicit

Sub shpPr�parationListe�g�e_Click()

    Call Cr�erListe�g�e

End Sub

Sub Cr�erListe�g�e() '2024-09-08 @ 15:55

    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Liste_Ag�e:Cr�erListe�g�e", "", 0)
   
    Application.ScreenUpdating = False
    
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
    niveauDetail = wshCAR_Liste_Ag�e.Range("B4").Value
    
    Application.EnableEvents = False
    
    'Ent�tes de colonnes en fonction du niveau de d�tail
    If LCase(niveauDetail) = "client" Then
        wshCAR_Liste_Ag�e.Range("B8:G8").Value = Array("Client", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Ag�e.Range("B8:G8"))
    End If

    'Ent�tes de colonnes en fonction du niveau de d�tail (Facture)
    If LCase(niveauDetail) = "facture" Then
        wshCAR_Liste_Ag�e.Range("B8:I8").Value = Array("Client", "No. Facture", "Date Facture", "Solde", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
        Call Make_It_As_Header(wshCAR_Liste_Ag�e.Range("B8:I8"))
    End If

    'Ent�tes de colonnes en fonction du niveau de d�tail (Transaction)
    If LCase(niveauDetail) = "transaction" Then
        wshCAR_Liste_Ag�e.Range("B8:J8").Value = Array("Client", "No. Facture", "Type", "Date", "Montant", "- de 30 jours", "31 @ 60 jours", "61 @ 90 jours", "+ de 90 jours")
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
    Dim derniereLigne As Long
    derniereLigne = wsFactures.Cells(wsFactures.Rows.count, 1).End(xlUp).row
    Dim rngFactures As Range
    Set rngFactures = wsFactures.Range("A3:A" & derniereLigne) '2 lignes d'ent�te
    
    Dim client As String, numFacture As String
    Dim dateFacture As Date, dateDue As Date
    Dim montantFacture As Currency, montantPaye As Currency, montantRegul As Currency, montantRestant As Currency
    Dim trancheAge As String
    Dim ageFacture As Long, i As Long, r As Long
    
    Application.EnableEvents = False

    r = 8
    For i = 1 To rngFactures.Rows.count
        'R�cup�rer les donn�es de la facture directement du Range
        numFacture = CStr(rngFactures.Cells(i, fFacCCInvNo).Value)
        'Do not process non Confirmed invoice
        If Fn_Get_Invoice_Type(numFacture) <> "C" Then
            GoTo Next_Invoice
        End If
        
        'Est-ce que la facture est � l'int�rieur de la date limite ?
        dateFacture = rngFactures.Cells(i, fFacCCInvoiceDate).Value
        If rngFactures.Cells(i, fFacCCInvoiceDate).Value > CDate(wshCAR_Liste_Ag�e.Range("H4").Value) Then
            Debug.Print "#022 - Comparaison de date - " & rngFactures.Cells(i, fFacCCInvoiceDate).Value & " .vs. " & wshCAR_Liste_Ag�e.Range("H4").Value
            GoTo Next_Invoice
        End If
        
        client = rngFactures.Cells(i, fFacCCCodeClient).Value
        'Obtenir le nom du client (MF) pour trier par nom de client plut�t que par code de client
        client = Fn_Get_Client_Name(client)
        dateDue = rngFactures.Cells(i, fFacCCDueDate).Value
        montantFacture = CCur(rngFactures.Cells(i, fFacCCTotal).Value)
        
        'Obtenir les paiements et r�gularisations pour cette facture
        montantPaye = Fn_Obtenir_Paiements_Facture(numFacture, wshCAR_Liste_Ag�e.Range("H4").Value)
        montantRegul = Fn_Obtenir_R�gularisations_Facture(numFacture, wshCAR_Liste_Ag�e.Range("H4").Value)
        
        montantRestant = montantFacture - montantPaye + montantRegul
        
        'Exclus les soldes de facture � 0,00 $ SI ET SEULMENT SI F4 = "NON"
        If UCase(wshCAR_Liste_Ag�e.Range("F4").Value) = "NON" And montantRestant = 0 Then
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
        Select Case LCase(niveauDetail)
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
                wshCAR_Liste_Ag�e.Cells(r, 2).Value = client
                wshCAR_Liste_Ag�e.Cells(r, 3).Value = numFacture
                wshCAR_Liste_Ag�e.Cells(r, 4).Value = dateFacture
                wshCAR_Liste_Ag�e.Cells(r, 4).NumberFormat = wshAdmin.Range("B1").Value
                wshCAR_Liste_Ag�e.Cells(r, 5).Value = montantRestant
                Select Case trancheAge
                    Case "- de 30 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 6).Value = montantRestant
                    Case "31 @ 60 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 7).Value = montantRestant
                    Case "61 @ 90 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 8).Value = montantRestant
                    Case "+ de 90 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 9).Value = montantRestant
                End Select
                
            Case "transaction"
                'La facture en premier...
                r = r + 1
                wshCAR_Liste_Ag�e.Cells(r, 2).Value = client
                wshCAR_Liste_Ag�e.Cells(r, 3).Value = numFacture
                wshCAR_Liste_Ag�e.Cells(r, 4).Value = "Facture"
                wshCAR_Liste_Ag�e.Cells(r, 5).Value = dateFacture
                wshCAR_Liste_Ag�e.Cells(r, 5).NumberFormat = wshAdmin.Range("B1").Value
                wshCAR_Liste_Ag�e.Cells(r, 6).Value = montantFacture
                Select Case trancheAge
                    Case "- de 30 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 7).Value = montantRestant
                    Case "31 @ 60 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 8).Value = montantRestant
                    Case "61 @ 90 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 9).Value = montantRestant
                    Case "+ de 90 jours"
                        wshCAR_Liste_Ag�e.Cells(r, 10).Value = montantRestant
                End Select
                
                'Transactions de paiements par la suite
                Dim rngPaiementsAssoc As Range
                Dim pmtFirstAddress As String
                'Obtenir tous les paiements pour la facture
                Set rngPaiementsAssoc = wsPaiements.Range("B:B").Find(numFacture, LookIn:=xlValues, LookAt:=xlWhole)
                If Not rngPaiementsAssoc Is Nothing Then
                    pmtFirstAddress = rngPaiementsAssoc.Address
                    Do
                        If rngPaiementsAssoc.offset(0, 2).Value <= CDate(wshCAR_Liste_Ag�e.Range("H4").Value) Then
                            r = r + 1
                            wshCAR_Liste_Ag�e.Cells(r, 2).Value = client
                            wshCAR_Liste_Ag�e.Cells(r, 3).Value = numFacture
                            wshCAR_Liste_Ag�e.Cells(r, 4).Value = "Paiement"
                            wshCAR_Liste_Ag�e.Cells(r, 5).Value = rngPaiementsAssoc.offset(0, 2).Value
                            wshCAR_Liste_Ag�e.Cells(r, 6).Value = -rngPaiementsAssoc.offset(0, 3).Value 'Montant du paiement
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
                        If rngR�gularisationAssoc.offset(0, 1).Value <= CDate(wshCAR_Liste_Ag�e.Range("H4").Value) Then
                            r = r + 1
                            wshCAR_Liste_Ag�e.Cells(r, 2).Value = client
                            wshCAR_Liste_Ag�e.Cells(r, 3).Value = numFacture
                            wshCAR_Liste_Ag�e.Cells(r, 4).Value = "R�gularisation"
                            wshCAR_Liste_Ag�e.Cells(r, 5).Value = rngR�gularisationAssoc.offset(0, 1).Value
                            wshCAR_Liste_Ag�e.Cells(r, 6).Value = rngR�gularisationAssoc.offset(0, 4).Value + _
                                rngR�gularisationAssoc.offset(0, 5).Value + _
                                rngR�gularisationAssoc.offset(0, 6).Value + _
                                rngR�gularisationAssoc.offset(0, 7).Value
                        End If
                        Set rngR�gularisationAssoc = wsR�gularisations.Columns("B:B").FindNext(rngR�gularisationAssoc)
                    Loop While Not rngR�gularisationAssoc Is Nothing And rngR�gularisationAssoc.Address <> regulFirstAddress
                End If
        End Select

Next_Invoice:
    Next i
    
    Application.EnableEvents = True
    
    'Si niveau de d�tail est "client", ajouter les soldes du client (dictionary) au tableau final
    If LCase(niveauDetail) = "client" Then
        r = 8
        Dim cle As Variant
        
        Application.EnableEvents = False
        
        For Each cle In dictClients.keys
            r = r + 1
            wshCAR_Liste_Ag�e.Cells(r, 2).Value = cle ' Nom du client
            wshCAR_Liste_Ag�e.Cells(r, 3).Value = dictClients(cle)(0) ' Total
            wshCAR_Liste_Ag�e.Cells(r, 4).Value = dictClients(cle)(1) ' - de 30 jours
            wshCAR_Liste_Ag�e.Cells(r, 5).Value = dictClients(cle)(2) ' 31 @ 60 jours
            wshCAR_Liste_Ag�e.Cells(r, 6).Value = dictClients(cle)(3) ' 61 @ 90 jours
            wshCAR_Liste_Ag�e.Cells(r, 7).Value = dictClients(cle)(4) ' + de 90 jours
        Next cle
        
        Application.EnableEvents = True

    End If
    
    'Tri alphab�tique par nom de client
    derniereLigne = wshCAR_Liste_Ag�e.Cells(wshCAR_Liste_Ag�e.Rows.count, "B").End(xlUp).row
    Set rngResultat = wshCAR_Liste_Ag�e.Range("B8:J" & derniereLigne)
    
    Application.EnableEvents = False
    
    Dim ordreTri As String
    If derniereLigne > 9 Then 'Le tri n'est peut-�tre pas n�cessaire
        With wshCAR_Liste_Ag�e.Sort
            .SortFields.Clear
            If wshCAR_Liste_Ag�e.Range("D4").Value = "Nom de client" Then
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
    
    derniereLigne = derniereLigne + 2
    
    Dim t(0 To 4) As Currency
    
    With wshCAR_Liste_Ag�e
        .Columns("B:B").ColumnWidth = 50
        .Columns("C:J").ColumnWidth = 13
        Select Case LCase(niveauDetail)
            Case "client"
                .Range("C9:G" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("C9:G" & derniereLigne).HorizontalAlignment = xlRight
                .Range("C" & derniereLigne).formula = "=Sum(C9:C" & derniereLigne - 2 & ")"
                t(0) = .Range("C" & derniereLigne).Value
                .Range("D" & derniereLigne).formula = "=Sum(D9:D" & derniereLigne - 2 & ")"
                t(1) = .Range("D" & derniereLigne).Value
                .Range("E" & derniereLigne).formula = "=Sum(E9:E" & derniereLigne - 2 & ")"
                t(2) = .Range("E" & derniereLigne).Value
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                t(3) = .Range("F" & derniereLigne).Value
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                t(4) = .Range("G" & derniereLigne).Value
                .Range("C" & derniereLigne & ":G" & derniereLigne).Font.Bold = True
            Case "facture"
                .Range("C9:C" & derniereLigne).HorizontalAlignment = xlCenter
                .Range("D9:D" & derniereLigne).HorizontalAlignment = xlCenter
                .Range("E9:I" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("E9:I" & derniereLigne).HorizontalAlignment = xlRight
                .Range("E" & derniereLigne).formula = "=Sum(E9:E" & derniereLigne - 2 & ")"
                t(0) = .Range("E" & derniereLigne).Value
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                t(1) = .Range("F" & derniereLigne).Value
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                t(2) = .Range("G" & derniereLigne).Value
                .Range("H" & derniereLigne).formula = "=Sum(H9:H" & derniereLigne - 2 & ")"
                t(3) = .Range("H" & derniereLigne).Value
                .Range("I" & derniereLigne).formula = "=Sum(I9:I" & derniereLigne - 2 & ")"
                t(4) = .Range("I" & derniereLigne).Value
                .Range("E" & derniereLigne & ":I" & derniereLigne).Font.Bold = True
            Case "transaction"
                .Columns("C:E").HorizontalAlignment = xlCenter
                .Columns("D").HorizontalAlignment = xlLeft
                .Range("F9:J" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("F9:J" & derniereLigne).HorizontalAlignment = xlRight
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                t(0) = .Range("F" & derniereLigne).Value
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                t(1) = .Range("G" & derniereLigne).Value
                .Range("H" & derniereLigne).formula = "=Sum(H9:H" & derniereLigne - 2 & ")"
                t(2) = .Range("H" & derniereLigne).Value
                .Range("I" & derniereLigne).formula = "=Sum(I9:I" & derniereLigne - 2 & ")"
                t(3) = .Range("I" & derniereLigne).Value
                .Range("J" & derniereLigne).formula = "=Sum(J9:J" & derniereLigne - 2 & ")"
                t(4) = .Range("J" & derniereLigne).Value
                .Range("F" & derniereLigne & ":J" & derniereLigne).Font.Bold = True
        End Select
        .Range("B" & derniereLigne).Value = "Totaux de la liste"
        .Range("B" & derniereLigne).Font.Bold = True
        derniereLigne = derniereLigne + 1
        
        'Ligne de pourcentages
        .Range("B" & derniereLigne).Value = "Pourcentages"
        .Range("B" & derniereLigne & ":J" & derniereLigne).Font.Bold = True
        .Range("C" & derniereLigne & ":J" & derniereLigne).NumberFormat = "##0.00"
        .Range("C" & derniereLigne & ":J" & derniereLigne).HorizontalAlignment = xlRight
        Dim totalListe As Currency
        totalListe = t(0)
        If totalListe <> 0 Then
            Select Case LCase(niveauDetail)
                Case "client"
                    .Range("C" & derniereLigne).Value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("D" & derniereLigne).Value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("E" & derniereLigne).Value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("F" & derniereLigne).Value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("G" & derniereLigne).Value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
                Case "facture"
                    .Range("E" & derniereLigne).Value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("F" & derniereLigne).Value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("G" & derniereLigne).Value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("H" & derniereLigne).Value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("I" & derniereLigne).Value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
                Case "transaction"
                    .Range("F" & derniereLigne).Value = Format$(Round(t(0) / totalListe, 4), "##0.00 %")
                    .Range("G" & derniereLigne).Value = Format$(Round(t(1) / totalListe, 4), "##0.00 %")
                    .Range("H" & derniereLigne).Value = Format$(Round(t(2) / totalListe, 4), "##0.00 %")
                    .Range("I" & derniereLigne).Value = Format$(Round(t(3) / totalListe, 4), "##0.00 %")
                    .Range("J" & derniereLigne).Value = Format$(Round(t(4) / totalListe, 4), "##0.00 %")
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

    Dim header1 As String: header1 = "Liste �g�e des comptes clients au " & wshCAR_Liste_Ag�e.Range("H4").Value
    Dim header2 As String
    If LCase(niveauDetail) = "client" Then
        header2 = "1 ligne par client"
    ElseIf LCase(niveauDetail) = "facture" Then
        header2 = "1 ligne par Facture"
    Else
        header2 = "1 ligne par transaction"
    End If
    header2 = ordreTri & " - " & header2
    
    Call Simple_Print_Setup(wshCAR_Liste_Ag�e, rngToPrint, header1, header2, "$8:$8", "L")
    
    Application.ScreenUpdating = True
    
    MsgBox "La pr�paration de la liste �g�e est termin�e", vbInformation
    
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
    
    Call Log_Record("modCAR_Liste_Ag�e:Cr�erListe�g�e", "", startTime)
    
End Sub

Sub shpRetourMenuFacturation_Click()

    Call RetourMenuFacturation

End Sub

Sub RetourMenuFacturation()

    wshCAR_Liste_Ag�e.Visible = xlSheetHidden
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select

End Sub


