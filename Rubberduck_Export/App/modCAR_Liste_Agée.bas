Attribute VB_Name = "modCAR_Liste_Agée"
Option Explicit

Sub CAR_Creer_Liste_Agee() '2024-09-08 @ 15:55

    Debug.Print "#005: " & Format$(Now(), "yyyy/mm/dd")
    
    'Initialiser les feuilles nécessaires
    Dim wsFactures As Worksheet
    Set wsFactures = ThisWorkbook.Sheets("FAC_Comptes_Clients")
    Dim wsPaiements As Worksheet
    Set wsPaiements = ThisWorkbook.Sheets("ENC_Détails")
    
    'Utilisation de la même feuille
    Dim rngResultat As Range
    Set rngResultat = wshCAR_Liste_Agée.Range("B8")
    Dim lastUsedRow As Long
    lastUsedRow = wshCAR_Liste_Agée.Cells(wshCAR_Liste_Agée.rows.count, "B").End(xlUp).Row
    If lastUsedRow > 7 Then
        Application.EnableEvents = False
        wshCAR_Liste_Agée.Range("B8:J" & lastUsedRow + 5).clear
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
    derniereLigne = wsFactures.Cells(wsFactures.rows.count, "A").End(xlUp).Row
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
            Debug.Print "#076 - Facture rejetée '" & numFacture & "', car son Statut (C_AC) n'est pas 'C' (Confirmée)"
            GoTo Next_Invoice
        End If
        'Est-ce que la facture est à l'intérieur e la date limite ?
        dateFacture = CDate(rngFactures.Cells(i, 2).value)
        If rngFactures.Cells(i, 2) > wshCAR_Liste_Agée.Range("H4").value Then
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
                    dictClients.add client, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0))
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
                wshCAR_Liste_Agée.Cells(r, 4).value = dateFacture
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
                wshCAR_Liste_Agée.Cells(r, 5).value = dateFacture
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
    derniereLigne = wshCAR_Liste_Agée.Cells(wshCAR_Liste_Agée.rows.count, "B").End(xlUp).Row
    Set rngResultat = wshCAR_Liste_Agée.Range("B8:J" & derniereLigne)
    
    Application.EnableEvents = False
    
    Dim ordreTri As String
    If derniereLigne > 9 Then 'Le tri n'est peut-être pas nécessaire
        With wshCAR_Liste_Agée.Sort
            .SortFields.clear
            If wshCAR_Liste_Agée.Range("D4").value = "Nom de client" Then
                .SortFields.add _
                    key:=wshCAR_Liste_Agée.Range("B8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par nom de client
                ordreTri = "Ordre de nom de client"
            Else
                .SortFields.add _
                    key:=wshCAR_Liste_Agée.Range("C8"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par numéro de facture
                .SortFields.add _
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
    
    Application.EnableEvents = True
    
    derniereLigne = derniereLigne + 2
    
    Application.EnableEvents = False
    
    With wshCAR_Liste_Agée
        .columns("B:B").ColumnWidth = 55
        .columns("C:I").ColumnWidth = 12
        Select Case LCase(niveauDetail)
            Case "client"
                .Range("C9:G" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("C9:G" & derniereLigne).HorizontalAlignment = xlRight
                .Range("C" & derniereLigne).formula = "=Sum(C9:C" & derniereLigne - 2 & ")"
                .Range("D" & derniereLigne).formula = "=Sum(D9:D" & derniereLigne - 2 & ")"
                .Range("E" & derniereLigne).formula = "=Sum(E9:E" & derniereLigne - 2 & ")"
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                .Range("C" & derniereLigne & ":G" & derniereLigne).Font.Bold = True
            Case "facture"
                .Range("C9:C" & derniereLigne).HorizontalAlignment = xlCenter
                .Range("D9:D" & derniereLigne).HorizontalAlignment = xlCenter
                .Range("E9:I" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("E9:I" & derniereLigne).HorizontalAlignment = xlRight
                .Range("E" & derniereLigne).formula = "=Sum(E9:E" & derniereLigne - 2 & ")"
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                .Range("H" & derniereLigne).formula = "=Sum(H9:H" & derniereLigne - 2 & ")"
                .Range("I" & derniereLigne).formula = "=Sum(I9:I" & derniereLigne - 2 & ")"
                .Range("E" & derniereLigne & ":I" & derniereLigne).Font.Bold = True
            Case "transaction"
                .columns("C:E").HorizontalAlignment = xlCenter
                .columns("D").HorizontalAlignment = xlLeft
                .Range("F9:J" & derniereLigne).NumberFormat = "#,##0.00 $"
                .Range("F9:J" & derniereLigne).HorizontalAlignment = xlRight
                .Range("F" & derniereLigne).formula = "=Sum(F9:F" & derniereLigne - 2 & ")"
                .Range("G" & derniereLigne).formula = "=Sum(G9:G" & derniereLigne - 2 & ")"
                .Range("H" & derniereLigne).formula = "=Sum(H9:H" & derniereLigne - 2 & ")"
                .Range("I" & derniereLigne).formula = "=Sum(I9:I" & derniereLigne - 2 & ")"
                .Range("J" & derniereLigne).formula = "=Sum(J9:J" & derniereLigne - 2 & ")"
                .Range("F" & derniereLigne & ":J" & derniereLigne).Font.Bold = True
        End Select
        .Range("B" & derniereLigne).value = "Totaux de la liste"
        .Range("B" & derniereLigne).Font.Bold = True
    End With
    
    Application.EnableEvents = True

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
    
    With rngToPrint.Font
        .name = "Segoe UI"
        .size = 9
    End With
    
    Application.EnableEvents = True

    Dim header1 As String: header1 = "Liste âgée des comptes clients"
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

End Sub

Sub zClean_Up_Summary_Area(ws As Worksheet)

    Application.EnableEvents = False
    
    'Cleanup the summary area (columns K to Q)
    ws.Range("J:P").clear
    'Erase any checkbox left over
    Call Delete_CheckBox
    
    Application.EnableEvents = True

End Sub

Sub zApply_Conditional_Formatting_Alternate_On_Column_H(rng As Range, lastUsedRow As Long)

    Dim ws As Worksheet: Set ws = wshCAR_Liste_Agée
    
    'Loop each cell in column C to find Totals row
    Dim totalRange As Range, cell As Range
    For Each cell In rng
        If InStr(1, cell.value, "Total ", vbTextCompare) > 0 Then
            If totalRange Is Nothing Then
                Set totalRange = ws.Cells(cell.Row, 8) 'Column H
            Else
                Set totalRange = Union(totalRange, ws.Cells(cell.Row, 8))
            End If
        End If
    Next cell
    
    'Check if any total rows were found
    rng.FormatConditions.delete

    'Define conditional formatting rules for the total rows
    If Not totalRange Is Nothing Then
        'Clear existing conditional formatting rules in the totalRange
        totalRange.FormatConditions.delete
        
        'Define conditional formatting rules for the totalRange
        With totalRange.FormatConditions
    
            'Rule for values > 50 (Highest priority)
            .add Type:=xlCellValue, Operator:=xlGreater, Formula1:="50"
            With .item(.count)
                .Interior.Color = RGB(255, 0, 0) 'Red color
            End With
    
            'Rule for values > 25
            .add Type:=xlCellValue, Operator:=xlGreater, Formula1:="25"
            With .item(.count)
                .Interior.Color = RGB(255, 165, 0) 'Orange color
            End With
    
            'Rule for values > 10
            .add Type:=xlCellValue, Operator:=xlGreater, Formula1:="10"
            With .item(.count)
                .Interior.Color = RGB(255, 255, 0) 'Yellow color
            End With
    
            'Rule for values > 5
            .add Type:=xlCellValue, Operator:=xlGreater, Formula1:="5"
            With .item(.count)
                .Interior.Color = RGB(144, 238, 144) 'Light green color
            End With
        End With
    End If
    
End Sub

Sub zBuild_Hours_Summary(rowSelected As Long)

    If rowSelected < 7 Then Exit Sub
    
    Dim ws As Worksheet: Set ws = wshCAR_Liste_Agée
    
    'Determine the last row used
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    
    'Clear the Hours Summary area
    Call Clean_Up_Summary_Area(ws)
    
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
    Dim i As Long, saveR As Long
    rowSelected = rowSelected + 1 'Summary starts on the next line (first line of expanded lines)
    saveR = rowSelected
    i = rowSelected
    Do Until Cells(i, 5) = ""
        If Cells(i, 6).value <> "" Then
            If dictHours.Exists(Cells(i, 6).value) Then
                dictHours(Cells(i, 6).value) = dictHours(Cells(i, 6).value) + Cells(i, 8).value
            Else
                dictHours.add Cells(i, 6).value, Cells(i, 8).value
            End If
        End If
        i = i + 1
    Loop

    Dim prof As Variant
    Dim profID As Long
    Dim tauxHoraire As Currency
    
    Application.EnableEvents = False
    
    ws.Range("O" & rowSelected).value = 0 'Reset the total WIP value
    For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) ' Sort dictionary by hours in descending order
        Cells(rowSelected, 10).value = prof
        Dim strProf As String
        strProf = prof
        profID = Fn_GetID_From_Initials(strProf)
        Cells(rowSelected, 11).HorizontalAlignment = xlRight
        Cells(rowSelected, 11).NumberFormat = "#,##0.00"
        Cells(rowSelected, 11).value = dictHours(prof)
        tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("H3").value)
        Cells(rowSelected, 12).value = tauxHoraire
        Cells(rowSelected, 13).NumberFormat = "#,##0.00$"
        Cells(rowSelected, 13).FormulaR1C1 = "=RC[-2]*RC[-1]"
        Cells(rowSelected, 13).HorizontalAlignment = xlRight
        rowSelected = rowSelected + 1
    Next prof
    
    'Sort the summary by rate (descending value) if required
    If rowSelected - 1 > saveR Then
        Dim rngSort As Range
        Set rngSort = ws.Range(ws.Cells(saveR, 10), ws.Cells(rowSelected - 1, 13))
        rngSort.Sort Key1:=ws.Cells(saveR, 13), Order1:=xlDescending, Header:=xlNo
    End If
    
    'Hours Total
    Dim rTotal As Long
    rTotal = rowSelected
    With Cells(rTotal, 11)
        .HorizontalAlignment = xlRight
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
'        .value = Format(t, "#,##0.00")
        .Font.Bold = True
    End With
    
    'Fees Total
    With Cells(rowSelected, 13)
        .HorizontalAlignment = xlRight
'        .value = Format(tdollars, "#,##0.00$")
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
        .Font.Bold = True
    End With
    
    With Range("J" & saveR & ":M" & rowSelected).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    With Range("K" & rowSelected & ", M" & rowSelected)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With

    'Save the TOTAL WIP value
    With ws.Range("N" & saveR)
        .value = "Valeur TEC:"
        .Font.Italic = True
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("O" & saveR)
        .NumberFormat = "#,##0.00 $"
        .value = ws.Range("M" & rowSelected).value
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    'Create a visual clue if amounts are different
    With ws.Range("O" & rowSelected)
        Dim formula As String
        formula = "=IF(M" & rowSelected & " <> O" & saveR & ", M" & rowSelected & "-O" & saveR & ",""""" & ")"
        Application.EnableEvents = False
        .formula = formula
        .NumberFormat = "#,##0.00 $"
        Application.EnableEvents = True
    End With
    
    Call Add_And_Modify_Checkbox(saveR, rowSelected)
    
    Application.EnableEvents = False

    'Clean up - 2024-07-11 @ 15:20
    Set dictHours = Nothing
    Set rngSort = Nothing
    Set ws = Nothing
    
End Sub
    
Sub zBring_In_Existing_Invoice_Requests(activeLastUsedRow As Long)

    Dim wsSource As Worksheet: Set wsSource = wshFAC_Projets_Entête
    Dim sourceLastUsedRow As Long
    sourceLastUsedRow = wsSource.Range("A9999").End(xlUp).Row
    
    Dim wsActive As Worksheet: Set wsActive = wshCAR_Liste_Agée
    Dim rngTotal As Range: Set rngTotal = wsActive.Range("C1:C" & activeLastUsedRow)
    
    'Analyze all Invoice Requests (one row at the time)
    
    Dim clientName As String
    Dim clientID As String
    Dim honoTotal As Double
    Dim result As Variant
    Dim i As Long, r As Long
    For i = 2 To sourceLastUsedRow
        If wsSource.Cells(i, 26).value <> "True" Then
            clientName = wsSource.Cells(i, 2).value
            clientID = wsSource.Cells(i, 3).value
            honoTotal = wsSource.Cells(i, 5).value
            'Using XLOOKUP to find the result directly
            result = Application.WorksheetFunction.XLookup("Total " & clientName, _
                                                           rngTotal, _
                                                           rngTotal, _
                                                           "Not Found", _
                                                           0, _
                                                           1)
            
            If result <> "Not Found" Then
                r = Application.WorksheetFunction.Match(result, rngTotal, 0)
                wsActive.Cells(r, 4).value = honoTotal
                wsActive.Cells(r, 4).NumberFormat = "#,##0.00 $"
            End If
        End If
    Next i

End Sub

Sub zFAC_Projets_Détails_Add_Record_To_DB(clientID As String, fr As Long, lr As Long, ByRef projetID As Long) 'Write a record to MASTER.xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Analyse:FAC_Projet_Détails_Add_Record_To_DB", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Détails"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    
    'Initialize recordset
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    
    'First SQL - SQL query to find the maximum value in the first column
    Dim strSQL As String
    strSQL = "SELECT MAX(ProjetID) AS MaxValue FROM [" & destinationTab & "$]"
    rs.Open strSQL, conn

    'Get the maximum value
    Dim MaxValue As Long
    If IsNull(rs.Fields("MaxValue").value) Then
        'Handle empty table (assign a default value, e.g., 1)
        projetID = 1
    Else
        projetID = rs.Fields("MaxValue").value + 1
    End If
    
    'Close the previous recordset (no longer needed)
    rs.Close
    
    'Second SQL - SQL query to add the new records
    strSQL = "SELECT * FROM [" & destinationTab & "$] WHERE 1=0"
    rs.Open strSQL, conn, 2, 3
    
    'Read all line from CAR_Analyse
    Dim dateTEC As String, TimeStamp As String
    Dim l As Long
    For l = fr To lr
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields("ProjetID").value = projetID
            rs.Fields("NomClient").value = wshCAR_Liste_Agée.Range("C" & l).value
            rs.Fields("ClientID").value = clientID
            rs.Fields("TECID").value = wshCAR_Liste_Agée.Range("A" & l).value
            rs.Fields("ProfID").value = wshCAR_Liste_Agée.Range("B" & l).value
            dateTEC = Format$(wshCAR_Liste_Agée.Range("E" & l).value, "dd/mm/yyyy")
            rs.Fields("Date").value = dateTEC
            rs.Fields("Prof").value = wshCAR_Liste_Agée.Range("F" & l).value
            rs.Fields("estDetruite") = 0 'Faux
            rs.Fields("Heures").value = CDbl(wshCAR_Liste_Agée.Range("H" & l).value)
            TimeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
            rs.Fields("TimeStamp").value = TimeStamp
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Open the MASTER file to clone the format to newly added lines
'    Call Clone_Last_Line_Formatting_For_New_Records(destinationFileName, destinationTab, (lr - fr + 1))
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modCAR_Analyse:FAC_Projet_Détails_Add_Record_To_DB()", startTime)

End Sub

Sub zFAC_Projets_Détails_Add_Record_Locally(clientID As String, fr As Long, lr As Long, projetID As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Analyse:FAC_Projet_Détails_Add_Record_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in FAC_Projets_Détails?
    Dim lastUsedRow As Long, rn As Long
    lastUsedRow = wshFAC_Projets_Détails.Range("A99999").End(xlUp).Row
    rn = lastUsedRow + 1
    
    Dim dateTEC As String, TimeStamp As String
    Dim i As Long
    For i = fr To lr
        wshFAC_Projets_Détails.Range("A" & rn).value = projetID
        wshFAC_Projets_Détails.Range("B" & rn).value = wshCAR_Liste_Agée.Range("C" & i).value
        wshFAC_Projets_Détails.Range("C" & rn).value = clientID
        wshFAC_Projets_Détails.Range("D" & rn).value = wshCAR_Liste_Agée.Range("A" & i).value
        wshFAC_Projets_Détails.Range("E" & rn).value = wshCAR_Liste_Agée.Range("B" & i).value
        dateTEC = Format$(wshCAR_Liste_Agée.Range("E" & i).value, "dd/mm/yyyy")
        wshFAC_Projets_Détails.Range("F" & rn).value = dateTEC
        wshFAC_Projets_Détails.Range("G" & rn).value = wshCAR_Liste_Agée.Range("F" & i).value
        wshFAC_Projets_Détails.Range("H" & rn).value = wshCAR_Liste_Agée.Range("H" & i).value
        wshFAC_Projets_Détails.Range("I" & rn).value = "FAUX"
        TimeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
        wshFAC_Projets_Détails.Range("J" & rn).value = TimeStamp
        rn = rn + 1
    Next i
    
    Call Log_Record("modCAR_Analyse:FAC_Projet_Détails_Add_Record_Locally()", startTime)

    Application.ScreenUpdating = True

End Sub

Sub zSoft_Delete_If_Value_Is_Found_In_Master_Details(filepath As String, _
                                                    sheetName As String, _
                                                    columnName As String, _
                                                    valueToFind As Variant) '2024-07-19 @ 15:31
    'Create a new ADODB connection
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    'Open the connection to the closed workbook
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filepath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'Update the rows to mark as deleted (soft delete)
    Dim strSQL As String
    strSQL = "UPDATE [" & sheetName & "$] SET estDetruite = -1 WHERE [" & columnName & "] = '" & Replace(valueToFind, "'", "''") & "'"
    cn.Execute strSQL
    
    'Close the connection
    cn.Close
    Set cn = Nothing
    
End Sub

Sub zFAC_Projets_Entête_Add_Record_To_DB(projetID As Long, _
                                        nomClient As String, _
                                        clientID As String, _
                                        dte As String, _
                                        hono As Double, _
                                        ByRef arr As Variant) 'Write a record to MASTER.xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Analyse:FAC_Projet_Entête_Add_Record_To_DB", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Entête"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "$] WHERE 1=0"
    
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conn, 2, 3
    
    Dim TimeStamp As String
    Dim c As Long
    Dim l As Long
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("ProjetID").value = projetID
        rs.Fields("NomClient").value = nomClient
        rs.Fields("ClientID").value = clientID
        rs.Fields("Date").value = dte
        rs.Fields("HonoTotal").value = hono
        For c = 1 To UBound(arr, 1)
            rs.Fields("Prof" & c).value = arr(c, 1)
            rs.Fields("Hres" & c).value = arr(c, 2)
            rs.Fields("TauxH" & c).value = arr(c, 3)
            rs.Fields("Hono" & c).value = arr(c, 4)
        Next c
        rs.Fields("estDétruite").value = 0 'Faux
        TimeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
        rs.Fields("TimeStamp").value = TimeStamp
    rs.update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Open the MASTER file to clone the format to newly added lines
'    Call Clone_Last_Line_Formatting_For_New_Records(destinationFileName, destinationTab, 1)
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modCAR_Analyse:FAC_Projet_Entête_Add_Record_To_DB()", startTime)

End Sub

Sub zFAC_Projets_Entête_Add_Record_Locally(projetID As Long, nomClient As String, clientID As String, dte As String, hono As Double, ByRef arr As Variant) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Analyse:FAC_Projet_Entête_Add_Record_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in FAC_Projets_Détails?
    Dim lastUsedRow As Long, rn As Long
    lastUsedRow = wshFAC_Projets_Entête.Range("A99999").End(xlUp).Row
    rn = lastUsedRow + 1
    
    Dim dateTEC As String, TimeStamp As String
    wshFAC_Projets_Entête.Range("A" & rn).value = projetID
    wshFAC_Projets_Entête.Range("B" & rn).value = nomClient
    wshFAC_Projets_Entête.Range("C" & rn).value = clientID
    wshFAC_Projets_Entête.Range("D" & rn).value = dte
    wshFAC_Projets_Entête.Range("E" & rn).value = hono
    'Assign values from the array to the worksheet using .Cells
    Dim i As Long, j As Long
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            wshFAC_Projets_Entête.Cells(rn, 6 + (i - 1) * UBound(arr, 2) + j - 1).value = arr(i, j)
        Next j
    Next i
    wshFAC_Projets_Entête.Range("Z" & rn).value = "FAUX"
    TimeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
    wshFAC_Projets_Entête.Range("AA" & rn).value = TimeStamp
    
    Call Log_Record("modCAR_Analyse:FAC_Projet_Entête_Add_Record_Locally()", startTime)

    Application.ScreenUpdating = True

End Sub

Sub zSoft_Delete_If_Value_Is_Found_In_Master_Entete(filepath As String, _
                                                   sheetName As String, _
                                                   columnName As String, _
                                                   valueToFind As Variant) '2024-07-19 @ 15:31
    'Create a new ADODB connection
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    'Open the connection to the closed workbook
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filepath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'Update the rows to mark as deleted (soft delete)
    Dim strSQL As String
    strSQL = "UPDATE [" & sheetName & "$] SET estDétruite = -1 WHERE [" & columnName & "] = '" & Replace(valueToFind, "'", "''") & "'"
    cn.Execute strSQL
    
    'Close the connection
    cn.Close
    Set cn = Nothing
    
End Sub

Sub zAdd_And_Modify_Checkbox(startRow As Long, lastRow As Long)
    
    'Set your worksheet (adjust this to match your worksheet name)
    Dim ws As Worksheet: Set ws = wshCAR_Liste_Agée
    
    'Define the range for the summary
    Dim summaryRange As Range
    Set summaryRange = ws.Range(ws.Cells(startRow, 10), ws.Cells(lastRow, 13)) 'Columns J to M
    
    'Add an ActiveX checkbox next to the summary in column O
    Dim checkBox As OLEObject
    With ws
        Set checkBox = .OLEObjects.add(ClassType:="Forms.CheckBox.1", _
                    Left:=.Cells(lastRow, 14).Left + 5, _
                    Top:=.Cells(lastRow, 14).Top, width:=80, Height:=16)
        
        'Modify checkbox properties
        With checkBox.Object
            .Caption = "On facture"
            .Font.size = 11  'Set font size
            .Font.Bold = True  'Set font bold
            .ForeColor = RGB(0, 0, 255)  'Set font color (Blue)
            .BackColor = RGB(200, 255, 200)  'Set background color (Light Green)
        End With
    End With
    
End Sub

Sub zDelete_CheckBox()

    'Set the worksheet
    Dim ws As Worksheet: Set ws = wshCAR_Liste_Agée
    
    'Check if any CheckBox exists and then delete it/them
    Dim Sh As Shape
    For Each Sh In ws.Shapes
        If InStr(Sh.name, "CheckBox") Then
            Sh.delete
        End If
    Next Sh
    
End Sub

Sub zGroups_SubTotals_Collapse_A_Client(r As Long)
    
    'Set the worksheet you want to work on
    Dim ws As Worksheet: Set ws = wshCAR_Liste_Agée

    'Loop through each row starting at row r
    Dim saveR As Long
    saveR = r
    Do While wshCAR_Liste_Agée.Range("A" & r).value <> ""
        r = r + 1
    Loop

    r = r - 1
    ws.rows(saveR & ":" & r).EntireRow.Hidden = True
    
End Sub

Sub zClear_Fees_Summary_And_CheckBox()

    'Clean the Fees Summary Area
    Dim ws As Worksheet: Set ws = wshCAR_Liste_Agée
    Application.EnableEvents = False
    ws.Range("J7:O9999").clear
    Application.EnableEvents = True
    
    'Clear any leftover CheckBox
    Dim Sh As Shape
    For Each Sh In ws.Shapes
        If InStr(Sh.name, "CheckBox") Then
            Sh.delete
        End If
    Next Sh

End Sub


