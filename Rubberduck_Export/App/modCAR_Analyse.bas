Attribute VB_Name = "modCAR_Analyse"
Option Explicit

Sub CAR_Creer_Liste_Agee()

    'Initialiser les feuilles nécessaires
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
        wsResultat.Range("A1:F1").Value = Array("Client", "Solde", "0-30 jours", "31-60 jours", "61-90 jours", "90+ jours")
        Call Make_It_As_Header(wsResultat.Range("A1:F1"))
    End If

    'Entêtes de colonnes en fonction du niveau de détail (Facture)
    If LCase(niveauDetail) = "facture" Then
        wsResultat.Range("A1:H1").Value = Array("Client", "No. Facture", "Date Facture", "Solde", "0-30 jours", "31-60 jours", "61-90 jours", "90+ jours")
        Call Make_It_As_Header(wsResultat.Range("A1:H1"))
    End If

    'Entêtes de colonnes en fonction du niveau de détail (Transaction)
    If LCase(niveauDetail) = "transaction" Then
        wsResultat.Range("A1:H1").Value = Array("Client", "No. Facture", "Date", "Montant", "0-30 jours", "31-60 jours", "61-90 jours", "90+ jours")
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
        client = rngFactures.Cells(i, 4).Value
        'Obtenir le nom du client pour trier par nom de client plutôt que par code de client
        client = Fn_Get_Client_Name(client)
        numFacture = rngFactures.Cells(i, 1).Value
        dateFacture = rngFactures.Cells(i, 2).Value
        dateDue = rngFactures.Cells(i, 7).Value
        montantFacture = CCur(rngFactures.Cells(i, 8).Value)
        
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
                wsResultat.Cells(r, 1).Value = client
                wsResultat.Cells(r, 2).Value = numFacture
                wsResultat.Cells(r, 3).Value = dateFacture
                wsResultat.Cells(r, 4).Value = montantRestant
                Select Case tranche
                    Case "0-30 jours"
                        wsResultat.Cells(r, 5).Value = montantRestant
                    Case "31-60 jours"
                        wsResultat.Cells(r, 6).Value = montantRestant
                    Case "61-90 jours"
                        wsResultat.Cells(r, 7).Value = montantRestant
                    Case "90+ jours"
                        wsResultat.Cells(r, 8).Value = montantRestant
                End Select
                
            Case "transaction"
                'Type Facture
                r = r + 1
                wsResultat.Cells(r, 1).Value = client
                wsResultat.Cells(r, 2).Value = numFacture
                wsResultat.Cells(r, 3).Value = dateFacture
                wsResultat.Cells(r, 4).Value = montantFacture
                Select Case tranche
                    Case "0-30 jours"
                        wsResultat.Cells(r, 5).Value = montantRestant
                    Case "31-60 jours"
                        wsResultat.Cells(r, 6).Value = montantRestant
                    Case "61-90 jours"
                        wsResultat.Cells(r, 7).Value = montantRestant
                    Case "90+ jours"
                        wsResultat.Cells(r, 8).Value = montantRestant
                End Select
                
                'Transactions de type paiements
                Dim rngPaiementsAssoc As Range
                Dim firstAddress As String
                Set rngPaiementsAssoc = wsPaiements.Range("B:B").Find(numFacture, LookIn:=xlValues, lookat:=xlWhole)
                If Not rngPaiementsAssoc Is Nothing Then
                    firstAddress = rngPaiementsAssoc.Address
                        Do
                        r = r + 1
                        wsResultat.Cells(r, 1).Value = client
                        wsResultat.Cells(r, 2).Value = numFacture
                        wsResultat.Cells(r, 3).Value = rngPaiementsAssoc.Offset(0, 2).Value
                        wsResultat.Cells(r, 4).Value = -rngPaiementsAssoc.Offset(0, 3).Value ' Montant du paiement
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
            wsResultat.Cells(i, 1).Value = cle ' Nom du client
            wsResultat.Cells(i, 2).Value = dictClients(cle)(0) ' Total
            wsResultat.Cells(i, 3).Value = dictClients(cle)(1) ' 0-30 jours
            wsResultat.Cells(i, 4).Value = dictClients(cle)(2) ' 31-60 jours
            wsResultat.Cells(i, 5).Value = dictClients(cle)(3) ' 61-90 jours
            wsResultat.Cells(i, 6).Value = dictClients(cle)(4) ' 90+ jours
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
    
    Call Simple_Print_Setup(wsResultat, rngToPrint, header1, header2, "L")
    
    MsgBox "La préparation est terminé" & vbNewLine & vbNewLine & "Voir la feuille 'X_Liste_Âgée_CAR'", vbInformation
    
    ThisWorkbook.Worksheets("X_Liste_Âgée_CAR").Activate
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set wsResultat = Nothing

End Sub

Sub CAR_Sort_Group_And_Subtotal() '2024-08-29 @ 22:24

'    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Analyse:CAR_Sort_Group_And_Subtotal",0)
    
    Application.ScreenUpdating = False
    
    Dim wsDest As Worksheet: Set wsDest = wshCAR_Liste_Agée
    
    'Remove existing subtotals in the destination worksheet
    wsDest.Cells.RemoveSubtotal
    
    'Clear the worksheet from row 5 until the last row used
    Dim destLastUsedRow As Long
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "B").End(xlUp).Row
    If destLastUsedRow < 5 Then destLastUsedRow = 5
    wsDest.Range("A5:H" & destLastUsedRow).clear
    
    'Build the dictionnary (Code, Nom du client) from Client's Master File
    Dim wsClientsMF As Worksheet: Set wsClientsMF = wshBD_Clients
    Dim lastUsedRowClient
    lastUsedRowClient = wsClientsMF.Cells(wsClientsMF.rows.count, "B").End(xlUp).Row
    Dim dictClients As Dictionary
    Set dictClients = New Dictionary
    Dim i As Long
    For i = 2 To lastUsedRowClient
        dictClients.add CStr(wsClientsMF.Cells(i, 2).Value), wsClientsMF.Cells(i, 1).Value
    Next i

    'Calculate the center of the used range
    Dim centerX As Double, centerY As Double
    centerX = 430
    centerY = 60

    'Set the dimensions of the progress bar
    Dim barWidth As Double, barHeight As Double
    barWidth = 300
    barHeight = 25  'Height of the progress bar

    'Create the background shape of the progress bar positioned in the center of the visible range
    Dim progressBarBg As Shape
    Set progressBarBg = ActiveSheet.Shapes.AddShape(msoShapeRectangle, centerX - barWidth / 3, centerY - barHeight / 2, barWidth, barHeight)
    progressBarBg.Fill.ForeColor.RGB = RGB(255, 255, 255)  ' White background
    progressBarBg.line.Visible = msoTrue  'Show the border of the progress bar
    progressBarBg.TextFrame.HorizontalAlignment = xlHAlignCenter
    progressBarBg.TextFrame.VerticalAlignment = xlVAlignCenter
    progressBarBg.TextFrame.Characters.Font.size = 14
    progressBarBg.TextFrame.Characters.Font.Color = RGB(0, 0, 0) 'Black font
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à 0 %"
    
    'Create the fill shape of the progress bar
    Dim progressBarFill As Shape
    Set progressBarFill = ActiveSheet.Shapes.AddShape(msoShapeRectangle, centerX - barWidth / 3, centerY - barHeight / 2, 0, barHeight)
    progressBarFill.Fill.ForeColor.RGB = RGB(0, 255, 0)  ' Green fill color
    progressBarFill.Fill.Transparency = 0.6  'Set transparency to 60%
    progressBarFill.line.Visible = msoFalse  'Hide the border of the fill
    
    'Update the progress bar fill
    progressBarFill.width = 0.15 * barWidth  '15 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.15, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    Dim lastUsedRow As Long, firstEmptyCol As Long
    
    'Set the source worksheet, lastUsedRow and lastUsedCol
    Dim wsSource As Worksheet: Set wsSource = wshFAC_Comptes_Clients
    'Find the last row with data in the source worksheet
    lastUsedRow = wsSource.Cells(wsSource.rows.count, "A").End(xlUp).Row
    'Find the first empty column from the left in the source worksheet
    firstEmptyCol = 1
    Do Until IsEmpty(wsSource.Cells(2, firstEmptyCol))
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    'Update the progress bar fill
    progressBarFill.width = 0.2 * barWidth  '20 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.2, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    Dim r As Long
    r = 6
    Application.EnableEvents = False
    For i = 3 To lastUsedRow
        'Conditions for exclusion (adjust as needed)
        If wsSource.Cells(i, 14).Value <> "VRAI" And _
            wsSource.Cells(i, 12).Value <> "VRAI" And _
            wsSource.Cells(i, 10).Value = "VRAI" Then
                If wsSource.Cells(i, ftecDate).Value <= wsDest.Range("H3").Value Then
                    'Get clients's name from MasterFile
                    Dim codeClient As String, nomClientFromMF As String
                    codeClient = wsSource.Cells(i, ftecClient_ID).Value
                    nomClientFromMF = dictClients(codeClient)
                    
                    wsDest.Cells(r, 1).Value = wsSource.Cells(i, ftecCAR_ID).Value
                    wsDest.Cells(r, 2).Value = wsSource.Cells(i, ftecProf_ID).Value
                    wsDest.Cells(r, 3).Value = nomClientFromMF
                    wsDest.Cells(r, 5).Value = wsSource.Cells(i, ftecDate).Value
                    wsDest.Cells(r, 6).Value = wsSource.Cells(i, ftecProf).Value
                    wsDest.Cells(r, 7).Value = wsSource.Cells(i, ftecDescription).Value
                    wsDest.Cells(r, 8).Value = wsSource.Cells(i, ftecHeures).Value
                    wsDest.Cells(r, 8).NumberFormat = "#,##0.00"
                    wsDest.Cells(r, 9).Value = wsSource.Cells(i, ftecCommentaireNote).Value
                    r = r + 1
                End If
        End If
    Next i
    Application.EnableEvents = False
   
    'Update the progress bar fill
    progressBarFill.width = 0.45 * barWidth  '45 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.45, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
   
    'Find the last row in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).Row

    'Sort by Client_ID (column E) and Date (column D) in the destination worksheet
    wsDest.Sort.SortFields.clear
    wsDest.Sort.SortFields.add key:=wsDest.Range("C6:C" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("E6:E" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("B6:B" & destLastUsedRow), Order:=xlAscending
    
    With wsDest.Sort
        .SetRange wsDest.Range("A6:I" & destLastUsedRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Update the progress bar fill
    progressBarFill.width = 0.6 * barWidth  '60 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.6, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    'Add subtotals for hours (column H) at each change in nomClientMF (column C) in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).Row
    Application.DisplayAlerts = False
    wsDest.Range("A5:H" & destLastUsedRow).Subtotal GroupBy:=3, Function:=xlSum, _
            TotalList:=Array(8), Replace:=True, PageBreaks:=False, SummaryBelowData:=False
    Application.DisplayAlerts = True
    wsDest.Range("A:B").EntireColumn.Hidden = True

    'Group the data to show subtotals in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).Row
    wsDest.Outline.ShowLevels RowLevels:=2
    
    'Add a formula to sum the billed amounts at the top row
    wsDest.Range("D6").formula = "=SUM(D7:D" & destLastUsedRow & ")"
    wsDest.Range("D6").NumberFormat = "#,##0.00 $"
    
    'Update the progress bar fill
    progressBarFill.width = 0.75 * barWidth  '75 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.75, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    'Change the format of the top row (Total General)
    With wsDest.Range("C6:D6")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .Color = -16776961
            .TintAndShade = 0
            .Bold = True
            .size = 12
        End With
    End With
    
    'Change the format of the top row (Hours)
    With wsDest.Range("H6")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .Bold = True
            .size = 12
        End With
    End With
    
    'Change the format of all Client's Total rows
    For r = 6 To destLastUsedRow
        If wsDest.Range("A" & r).Value = "" Then
            With wsDest.Range("C" & r).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With wsDest.Range("C" & r).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            With wsDest.Range("C" & r)
'                If InStr(.Value, "Total ") = 1 Then
'                    .Value = Mid(.Value, 7)
'                End If
                If .Value = "Total général" Then
                    .Value = "G r a n d   T o t a l"
                End If
            End With
        End If
    Next r
    
    'Update the progress bar fill
    progressBarFill.width = 0.85 * barWidth  '85 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.85, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    'Set conditional formats for total hours (Client's total)
    Dim rngTotals As Range: Set rngTotals = wsDest.Range("C7:C" & destLastUsedRow)
    Call Apply_Conditional_Formatting_Alternate_On_Column_H(rngTotals, destLastUsedRow)
    
    'Bring in all the invoice requests
    Call Bring_In_Existing_Invoice_Requests(destLastUsedRow)
    
    'Clean up the summary aera of the worksheet
    Call Clean_Up_Summary_Area(wsDest)
    
    'Update the progress bar fill
    progressBarFill.width = 0.95 * barWidth   '95 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.95, "0%")
    
    'Introduce a small delay to ensure the worksheet is fully updated
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01")) '2024-07-23 @ 16:13 - Slowdown the application
        
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    progressBarBg.delete
    progressBarFill.delete
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Active le volet inférieur (Pane 2) et défile pour positionner la ligne 7 en haut de ce volet
    With ActiveWindow.Panes(2)
        .ScrollRow = 7
    End With
    
    'Optionnel : Sélectionne la cellule I7
'    Range("I7").Select
    
'    Application.StatusBar = ""

'    Call Log_Record("modCAR_Analyse:CAR_Sort_Group_And_Subtotal()", startTime)

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
        If InStr(1, cell.Value, "Total ", vbTextCompare) > 0 Then
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
        If Cells(i, 6).Value <> "" Then
            If dictHours.Exists(Cells(i, 6).Value) Then
                dictHours(Cells(i, 6).Value) = dictHours(Cells(i, 6).Value) + Cells(i, 8).Value
            Else
                dictHours.add Cells(i, 6).Value, Cells(i, 8).Value
            End If
        End If
        i = i + 1
    Loop

    Dim prof As Variant
    Dim profID As Long
    Dim tauxHoraire As Currency
    
    Application.EnableEvents = False
    
    ws.Range("O" & rowSelected).Value = 0 'Reset the total WIP value
    For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) ' Sort dictionary by hours in descending order
        Cells(rowSelected, 10).Value = prof
        Dim strProf As String
        strProf = prof
        profID = Fn_GetID_From_Initials(strProf)
        Cells(rowSelected, 11).HorizontalAlignment = xlRight
        Cells(rowSelected, 11).NumberFormat = "#,##0.00"
        Cells(rowSelected, 11).Value = dictHours(prof)
        tauxHoraire = Fn_Get_Hourly_Rate(profID, ws.Range("H3").Value)
        Cells(rowSelected, 12).Value = tauxHoraire
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
'        .Value = Format(t, "#,##0.00")
        .Font.Bold = True
    End With
    
    'Fees Total
    With Cells(rowSelected, 13)
        .HorizontalAlignment = xlRight
'        .Value = Format(tdollars, "#,##0.00$")
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
        .Value = "Valeur TEC:"
        .Font.Italic = True
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("O" & saveR)
        .NumberFormat = "#,##0.00 $"
        .Value = ws.Range("M" & rowSelected).Value
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
        If wsSource.Cells(i, 26).Value <> "True" Then
            clientName = wsSource.Cells(i, 2).Value
            clientID = wsSource.Cells(i, 3).Value
            honoTotal = wsSource.Cells(i, 5).Value
            'Using XLOOKUP to find the result directly
            result = Application.WorksheetFunction.XLookup("Total " & clientName, _
                                                           rngTotal, _
                                                           rngTotal, _
                                                           "Not Found", _
                                                           0, _
                                                           1)
            
            If result <> "Not Found" Then
                r = Application.WorksheetFunction.Match(result, rngTotal, 0)
                wsActive.Cells(r, 4).Value = honoTotal
                wsActive.Cells(r, 4).NumberFormat = "#,##0.00 $"
            End If
        End If
    Next i

End Sub

Sub zFAC_Projets_Détails_Add_Record_To_DB(clientID As String, fr As Long, lr As Long, ByRef projetID As Long) 'Write a record to MASTER.xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR_Analyse:FAC_Projet_Détails_Add_Record_To_DB", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
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
    If IsNull(rs.Fields("MaxValue").Value) Then
        'Handle empty table (assign a default value, e.g., 1)
        projetID = 1
    Else
        projetID = rs.Fields("MaxValue").Value + 1
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
            rs.Fields("ProjetID").Value = projetID
            rs.Fields("NomClient").Value = wshCAR_Liste_Agée.Range("C" & l).Value
            rs.Fields("ClientID").Value = clientID
            rs.Fields("TECID").Value = wshCAR_Liste_Agée.Range("A" & l).Value
            rs.Fields("ProfID").Value = wshCAR_Liste_Agée.Range("B" & l).Value
            dateTEC = Format$(wshCAR_Liste_Agée.Range("E" & l).Value, "dd/mm/yyyy")
            rs.Fields("Date").Value = dateTEC
            rs.Fields("Prof").Value = wshCAR_Liste_Agée.Range("F" & l).Value
            rs.Fields("estDetruite") = 0 'Faux
            rs.Fields("Heures").Value = CDbl(wshCAR_Liste_Agée.Range("H" & l).Value)
            TimeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
            rs.Fields("TimeStamp").Value = TimeStamp
        rs.update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Open the MASTER file to clone the format to newly added lines
    Call Clone_Last_Line_Formatting_For_New_Records(destinationFileName, destinationTab, (lr - fr + 1))
    
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
        wshFAC_Projets_Détails.Range("A" & rn).Value = projetID
        wshFAC_Projets_Détails.Range("B" & rn).Value = wshCAR_Liste_Agée.Range("C" & i).Value
        wshFAC_Projets_Détails.Range("C" & rn).Value = clientID
        wshFAC_Projets_Détails.Range("D" & rn).Value = wshCAR_Liste_Agée.Range("A" & i).Value
        wshFAC_Projets_Détails.Range("E" & rn).Value = wshCAR_Liste_Agée.Range("B" & i).Value
        dateTEC = Format$(wshCAR_Liste_Agée.Range("E" & i).Value, "dd/mm/yyyy")
        wshFAC_Projets_Détails.Range("F" & rn).Value = dateTEC
        wshFAC_Projets_Détails.Range("G" & rn).Value = wshCAR_Liste_Agée.Range("F" & i).Value
        wshFAC_Projets_Détails.Range("H" & rn).Value = wshCAR_Liste_Agée.Range("H" & i).Value
        wshFAC_Projets_Détails.Range("I" & rn).Value = "FAUX"
        TimeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
        wshFAC_Projets_Détails.Range("J" & rn).Value = TimeStamp
        rn = rn + 1
    Next i
    
    Call Log_Record("modCAR_Analyse:FAC_Projet_Détails_Add_Record_Locally()", startTime)

    Application.ScreenUpdating = True

End Sub

Sub zSoft_Delete_If_Value_Is_Found_In_Master_Details(filePath As String, _
                                                    sheetName As String, _
                                                    columnName As String, _
                                                    valueToFind As Variant) '2024-07-19 @ 15:31
    'Create a new ADODB connection
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    'Open the connection to the closed workbook
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
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
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
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
        rs.Fields("ProjetID").Value = projetID
        rs.Fields("NomClient").Value = nomClient
        rs.Fields("ClientID").Value = clientID
        rs.Fields("Date").Value = dte
        rs.Fields("HonoTotal").Value = hono
        For c = 1 To UBound(arr, 1)
            rs.Fields("Prof" & c).Value = arr(c, 1)
            rs.Fields("Hres" & c).Value = arr(c, 2)
            rs.Fields("TauxH" & c).Value = arr(c, 3)
            rs.Fields("Hono" & c).Value = arr(c, 4)
        Next c
        rs.Fields("estDétruite").Value = 0 'Faux
        TimeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
        rs.Fields("TimeStamp").Value = TimeStamp
    rs.update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Open the MASTER file to clone the format to newly added lines
    Call Clone_Last_Line_Formatting_For_New_Records(destinationFileName, destinationTab, 1)
    
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
    wshFAC_Projets_Entête.Range("A" & rn).Value = projetID
    wshFAC_Projets_Entête.Range("B" & rn).Value = nomClient
    wshFAC_Projets_Entête.Range("C" & rn).Value = clientID
    wshFAC_Projets_Entête.Range("D" & rn).Value = dte
    wshFAC_Projets_Entête.Range("E" & rn).Value = hono
    'Assign values from the array to the worksheet using .Cells
    Dim i As Long, j As Long
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            wshFAC_Projets_Entête.Cells(rn, 6 + (i - 1) * UBound(arr, 2) + j - 1).Value = arr(i, j)
        Next j
    Next i
    wshFAC_Projets_Entête.Range("Z" & rn).Value = "FAUX"
    TimeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
    wshFAC_Projets_Entête.Range("AA" & rn).Value = TimeStamp
    
    Call Log_Record("modCAR_Analyse:FAC_Projet_Entête_Add_Record_Locally()", startTime)

    Application.ScreenUpdating = True

End Sub

Sub zSoft_Delete_If_Value_Is_Found_In_Master_Entete(filePath As String, _
                                                   sheetName As String, _
                                                   columnName As String, _
                                                   valueToFind As Variant) '2024-07-19 @ 15:31
    'Create a new ADODB connection
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    'Open the connection to the closed workbook
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
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
    Do While wshCAR_Liste_Agée.Range("A" & r).Value <> ""
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


