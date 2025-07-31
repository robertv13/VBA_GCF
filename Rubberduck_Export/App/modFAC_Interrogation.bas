Attribute VB_Name = "modFAC_Interrogation"
'@IgnoreModule SetAssignmentWithIncompatibleObjectType

Option Explicit

Sub shpAfficherFactures_Click()

    Call AfficherListeDesFactures

End Sub

Sub shpSaisirAutreClient_Click()

    Call SaisirAutreClient

End Sub

Sub AfficherListeDesFactures()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:AfficherListeDesFactures", vbNullString, 0)
    
    Application.EnableEvents = False
    With wshFAC_Interrogation.Range("B9:O33")
        .ClearContents
        .Font.Color = vbBlack
        .Font.Bold = False
    End With
    Application.EnableEvents = True
    
    Dim ws As Worksheet: Set ws = wshFAC_Interrogation
    
    Application.ScreenUpdating = False
    
    Dim clientName As String: clientName = ws.Range("D4").Value
    Dim dateFrom As Date: dateFrom = ws.Range("G6").Value
    Dim dateTo As Date: dateTo = ws.Range("I6").Value
    
    'What is the ID for the selected client ?
    Dim myInfo() As Variant
    Dim rng As Range: Set rng = wsdBD_Clients.Range("dnrClients_Names_Only")
    myInfo = Fn_Find_Data_In_A_Range(rng, 1, clientName, fClntFMClientID)
    If myInfo(1) = vbNullString Then
        MsgBox "Je ne peux retrouver ce client dans ma liste de clients", vbCritical
        GoTo Clean_Exit
    End If
    
    Dim codeClient As String
    codeClient = myInfo(3)
    
    Call FAC_Get_Invoice_Client_AF(codeClient)
    
    Call modFAC_Interrogation.CopierListeFacturesVersWorksheet(dateFrom, dateTo)
    
    'Ajuste les 2 boutons
    Dim shp As Shape
    Set shp = wshFAC_Interrogation.Shapes("shpAfficherFactures")
    shp.Visible = False
    Set shp = wshFAC_Interrogation.Shapes("shpSaisirAutreClient")
    shp.Top = 70
    shp.Visible = True
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:AfficherListeDesFactures", vbNullString, startTime)

Clean_Exit:
    
    'Libérer la mémoire
    Set rng = Nothing
    Set shp = Nothing
    Set ws = Nothing
    
    DoEvents
    
End Sub

Sub SaisirAutreClient()

    Call FAC_Historique_Clear_All_Cells
    
    Dim shp As Shape
    Set shp = wshFAC_Interrogation.Shapes("shpSaisirAutreClient")
    shp.Visible = False

End Sub

Sub FAC_Get_Invoice_Client_AF(codeClient As String) '2024-06-27 @ 15:27

    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    
    'wsdFAC_Entete_AF#1

    With ws
        'Effacer les données de la dernière utilisation
        .Range("Y6:Y10").ClearContents
        .Range("Y6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
        'Définir le range pour la source des données en utilisant un tableau
        Dim rngData As Range
        Set rngData = .Range("l_tbl_FAC_Entete[#All]")
        .Range("Y7").Value = rngData.Address
        
        'Définir le range des critères
        Dim rngCriteria As Range
        Set rngCriteria = .Range("Y2:Y3")
        .Range("Y3").Value = codeClient
        .Range("Y8").Value = rngCriteria.Address
        
        'Définir le range des résultats et effacer avant le traitement
        Dim rngResult As Range
        Set rngResult = .Range("AA1").CurrentRegion
        rngResult.offset(2, 0).Clear
        Set rngResult = .Range("AA2:AV2")
        .Range("Y9").Value = rngResult.Address
        
        rngData.AdvancedFilter _
                    action:=xlFilterCopy, _
                    criteriaRange:=rngCriteria, _
                    CopyToRange:=rngResult, _
                    Unique:=False
          
        'Quels sont les résultats ?
        Dim lastResultRow As Long
        lastResultRow = .Cells(.Rows.count, "AA").End(xlUp).Row
        .Range("Y10").Value = lastResultRow - 2 & " lignes"
         
        'Est-il nécessaire de trier les résultats ?
        If lastResultRow < 4 Then Exit Sub
        With .Sort 'Sort - Invoice Date
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("AA3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortTextAsNumbers 'Sort Based On Invoice Number
            .SetRange ws.Range("AA3:AV" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
     End With

    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing

End Sub

Sub CopierListeFacturesVersWorksheet(dateMin As Date, dateMax As Date)

    Dim ws As Worksheet: Set ws = wsdFAC_Entete
    
    'Détermine la dernière utilisée dans les résultats de AF_1 dans wsdFAC_Entete
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AA").End(xlUp).Row
    If lastUsedRow < 3 Then Exit Sub 'Nothing to display
    
    Dim arr() As Variant
    ReDim arr(1 To 250, 0 To 13)
    Dim resultArr As Variant
    
    With ws
        Dim i As Long, r As Long
        For i = 3 To lastUsedRow
            'Vérification de la date de facture -ET- si la facture est bel et bien confirmée
            If .Range("AB" & i).Value >= dateMin And .Range("AB" & i).Value <= dateMax Then
                r = r + 1
                arr(r, 0) = .Range("AC" & i).Value 'ACouC
                arr(r, 1) = .Range("AA" & i).Value 'Invoice number
                arr(r, 2) = .Range("AB" & i).Value 'Invoice Date
                arr(r, 3) = .Range("AJ" & i).Value 'Fees
                arr(r, 4) = .Range("AL" & i).Value 'Misc. 1
                arr(r, 5) = .Range("AN" & i).Value 'Misc. 2
                arr(r, 6) = .Range("AP" & i).Value 'Misc. 3
                arr(r, 7) = .Range("AR" & i).Value 'GST $
                arr(r, 8) = .Range("AT" & i).Value 'PST $
                arr(r, 9) = .Range("AV" & i).Value 'Deposit
                arr(r, 10) = .Range("AU" & i).Value 'AR_Total
                arr(r, 11) = modFunctions.Fn_Get_Invoice_Total_Payments_AF(.Range("AA" & i).Value) - _
                                modFunctions.Fn_Get_Invoice_Total_Regul_AF(.Range("AA" & i).Value)
                arr(r, 12) = CDate(modFunctions.Fn_Get_A_Cell_From_A_Worksheet("FAC_Comptes_Clients", .Range("AA" & i).Value, fFacCCInvNo, fFacCCDueDate))
                'Obtenir les TEC facturés par cette facture
                arr(r, 13) = modFunctions.Fn_Get_TEC_Total_Invoice_AF(.Range("AA" & i).Value, "Dollars")
            End If
        Next i
    End With
    
    If r = 0 Then
        MsgBox "Il n'y a aucune facture pour la période recherchée", vbExclamation
        GoTo Clean_Exit
    End If
    
    'Transfer the arr to the worksheet, after resizing it
    Call RedimensionnerTableau2D(arr, r, 13)

    Application.EnableEvents = False
    
    Dim ACouC As String
    With wshFAC_Interrogation
        For i = 1 To UBound(arr, 1)
            ACouC = arr(i, 0)
            .Range("C" & i + 8).Value = arr(i, 1)
            .Range("D" & i + 8).Value = Format$(arr(i, 2), wsdADMIN.Range("B1").Value)
            .Range("E" & i + 8).Value = arr(i, 3)
            .Range("F" & i + 8).Value = arr(i, 13)
            .Range("G" & i + 8).Value = arr(i, 4)
            .Range("H" & i + 8).Value = arr(i, 5)
            .Range("I" & i + 8).Value = arr(i, 6)
            .Range("J" & i + 8).Value = arr(i, 7)
            .Range("K" & i + 8).Value = arr(i, 8)
            .Range("L" & i + 8).Value = arr(i, 9)
            .Range("M" & i + 8).Value = arr(i, 10)
            If arr(i, 10) - arr(i, 11) > 0 Then
                .Range("N" & i + 8).Value = Format$(WorksheetFunction.Max(0, Date - arr(i, 12)), "# ###")
            End If
            .Range("O" & i + 8).Value = arr(i, 10) - arr(i, 11) 'Balance
            If ACouC = "AC" Then
                With wshFAC_Interrogation.Range("B" & i + 8 & ":O" & i + 8)
                    .Font.Color = vbRed
                    .Font.Bold = True
                End With
                wshFAC_Interrogation.Range("B" & i + 8).Value = "AC"
            End If
        Next i
    End With
    
    lastUsedRow = i + 8
    
    Application.EnableEvents = True

Clean_Exit:

    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub FAC_Historique_Clear_All_Cells()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:FAC_Historique_Clear_All_Cells", vbNullString, 0)
    
    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    On Error Resume Next
    ActiveSheet.Unprotect
    On Error GoTo 0
    With wshFAC_Interrogation
        .Range("D4:H4, D6:E6").ClearContents
        .Range("G6, I6").ClearContents
        .Range("B9:R33").ClearContents
        Application.EnableEvents = True
        .Activate
        .Range("D4").Select
    End With
    
    With ActiveSheet
        .Protect userInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:FAC_Historique_Clear_All_Cells", vbNullString, startTime)

End Sub

Sub shp_FAC_Historique_Exit_Click()

    Call FAC_Historique_Back_To_FAC_Menu

End Sub

Sub FAC_Historique_Back_To_FAC_Menu()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:FAC_Historique_Back_To_FAC_Menu", vbNullString, 0)
    
    wshFAC_Interrogation.Visible = xlSheetHidden
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:FAC_Historique_Back_To_FAC_Menu", vbNullString, startTime)

End Sub

Sub FAC_Historique_Montrer_Bouton_Afficher()

    Dim shp As Shape: Set shp = wshFAC_Interrogation.Shapes("shpAfficherFactures")
    
    Application.EnableEvents = False
    
    If IsDate(wshFAC_Interrogation.Range("G6").Value) And _
        IsDate(wshFAC_Interrogation.Range("I6").Value) And _
        Trim$(wshFAC_Interrogation.Range("D4").Value) <> vbNullString Then
        shp.Top = 70
        shp.Visible = True
        Set shp = wshFAC_Interrogation.Shapes("shpSaisirAutreClient")
        shp.Visible = False
    Else
        shp.Visible = False
    End If
    
    Application.EnableEvents = True

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub FAC_Historique_Montrer_Bouton_AutreClient()

    Dim shp As Shape: Set shp = wshFAC_Interrogation.Shapes("shpSaisirAutreClient")
    
    Application.EnableEvents = False
    
    shp.Top = 70
    shp.Visible = True
    
    Application.EnableEvents = True

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub FAC_Interrogation_AfficherMenuContextuel(ByVal Target As Range) '2025-01-28 @ 10:19

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
    
    'Créer un nouveau menu contextuel
    Set menu = Application.CommandBars.Add(Name:="FactureMenu", position:=msoBarPopup, Temporary:=True)

    'Ajout de l'option 1 au menu contextuel
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.Caption = "Visualiser la facture (format PDF)"
        menuItem.OnAction = "'VisualiserFacturePDF """ & numeroFacture & """'"

    'Ajout de l'option 2 au menu contextuel
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.Caption = "TEC + Honoraires de la facture"
        menuItem.OnAction = "'TEC_HONO_Facture """ & Target.Address & """'"

    'Ajout de l'option 3 au menu contextuel
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.Caption = "TEC détaillé pour la facture"
        menuItem.OnAction = "'ObtenirListeTECFactures """ & Target.Address & """'"

    'Ajout de l'option 4 au menu contextuel
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
        menuItem.Caption = "Transactions des Comptes-Clients"
        menuItem.OnAction = "'TransactionComptesClients """ & Target.Address & """'"

    'Afficher le menu contextuel
    menu.ShowPopup
    
End Sub

Sub VisualiserFacturePDF(noFact As String)

    'The invoice number is in column C (3rd column) pour les 2 feuilles...
    Dim fullPDFFileName As String
    fullPDFFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gFACT_PDF_PATH & _
                            Application.PathSeparator & noFact & ".pdf"
    
    'Ouvrir la version PDF de la facture, si elle existe
    If Dir(fullPDFFileName) <> vbNullString Then
        'Le fichier existe, on peut lancer la commande Shell pour l'ouvrir
        Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " & Chr$(34) & fullPDFFileName & Chr$(34), vbNormalFocus
    Else
        'Le fichier n'existe pas, afficher un message d'erreur
        MsgBox "La version PDF de cette facture n'existe pas" & vbNewLine & vbNewLine & _
                                                        fullPDFFileName, vbExclamation, "Fichier PDF introuvable"
    End If
    
End Sub

Sub TEC_HONO_Facture(adresse As String)

    Dim numeroLigne As Long, numeroColonne As Long
    Call ExtraireLigneColonneCellule(adresse, numeroLigne, numeroColonne)

    Dim ws As Worksheet: Set ws = wshFAC_Interrogation
    
    'The invoice number is in column C (3rd column)
    Dim invNo As String
    invNo = ws.Cells(numeroLigne, 3).Value
    
    'Nom du client et date de facture
    Dim nomClient As String
    nomClient = ws.Range("D4").Value
    Dim dateFacture As Date
    dateFacture = Format$(ws.Cells(numeroLigne, 4).Value, wsdADMIN.Range("B1").Value)
    
    Call ObtenirFactureInfos(invNo, nomClient, dateFacture)
    
End Sub

Sub TransactionComptesClients(adresse As String)

    Dim numeroLigne As Long, numeroColonne As Long
    Call ExtraireLigneColonneCellule(adresse, numeroLigne, numeroColonne)

    Dim ws As Worksheet: Set ws = wshFAC_Interrogation
    
    'The invoice number is in column C (3rd column)
    Dim invNo As String
    invNo = ws.Cells(numeroLigne, 3).Value
    
    'Nom du client et date de facture
    Dim nomClient As String
    nomClient = ws.Range("D4").Value
    Dim dateFacture As Date
    dateFacture = ws.Cells(numeroLigne, 4).Value
    
    Call ObtenirTransactionsCC(invNo, nomClient, dateFacture)

End Sub

Sub ExtraireLigneColonneCellule(adresse As String, ByRef numeroLigne As Long, ByRef numeroColonne As Long)

    Dim cellule As Range
    Set cellule = ActiveSheet.Range(adresse)
    
    numeroLigne = cellule.row
    numeroColonne = cellule.Column
    
End Sub

Sub ObtenirFactureInfos(noFact As String, nomClient As String, dateFacture As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:ObtenirFactureInfos", vbNullString, 0)
    
    Call AfficherNouvelleFeuille_Stats(noFact, nomClient, dateFacture)
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:ObtenirFactureInfos", vbNullString, startTime)

End Sub

Sub AfficherNouvelleFeuille_Stats(invNo As String, nomClient As String, dateFacture As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:AfficherNouvelleFeuille_Stats", vbNullString, 0)
    
    If invNo = vbNullString Then
        Exit Sub
    End If
    
    Dim sheetName As String
    sheetName = "FactureInfo_" & invNo
    
    ' Référence à la première feuille
    Dim wsSelection As Worksheet
    Set wsSelection = wshFAC_Interrogation
    
    'Vérifier si la feuille existe déjà
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    'Si la feuille existe, la supprimer
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Créer une nouvelle feuille
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = sheetName
    
    'Entête de la feuille
    ws.Range("B1:J1").Merge
    With ws.Range("B1")
        .Value = "Informations sur les TEC & les Honoraires"
        .Font.size = 22
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("D4").Value = "No. de Facture:"
    ws.Range("D4").HorizontalAlignment = xlRight
    ws.Range("D4").Font.Italic = True
    ws.Range("D4").Font.size = 9
    
    ws.Range("E4").Value = invNo
    ws.Range("E4").HorizontalAlignment = xlCenter
    ws.Range("E4").Font.Bold = True
    ws.Range("E4").Font.size = 12
    
    ws.Range("D6").Value = "Nom du Client:"
    ws.Range("D6").HorizontalAlignment = xlRight
    ws.Range("D6").Font.Italic = True
    ws.Range("D6").Font.size = 9
    
    If Len(nomClient) > 59 Then nomClient = Left$(nomClient, 59) & "..."
    
    With ws.Range("E6")
        .Value = nomClient
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.size = 12
    End With
    
    ws.Range("G4").Value = "Date de Facture:"
    ws.Range("G4").HorizontalAlignment = xlRight
    ws.Range("G4").Font.Italic = True
    ws.Range("G4").Font.size = 9
    
    ws.Range("H4").Value = dateFacture
    ws.Range("H4").HorizontalAlignment = xlCenter
    ws.Range("H4").Font.Bold = True
    ws.Range("H4").Font.size = 12
    
    'Obtenir les valeurs des tableaux pour TEC et HONORAIRES
    Dim tableauTEC As Variant
    tableauTEC = ObtenirTableauTEC(invNo)
    Dim tableauHonoraires As Variant
    tableauHonoraires = ObtenirTableauHonoraires(invNo)

    ws.Range("D8").Value = "Travaux en cours"
    ws.Range("D8").Font.Italic = True
    ws.Range("D8").Font.Bold = True
    ws.Range("D8").Font.size = 11
    
    'Remplir le tableau TEC (9 x 4)
    Dim rOffset As Integer
    rOffset = 8
    Dim cOffset As Integer
    cOffset = 4
    Dim lastRowUsed As Long, nbItemTEC As Long
    Dim totHres As Currency, totValeur As Currency
    Dim i As Integer, j As Integer
    For i = LBound(tableauTEC, 1) To UBound(tableauTEC, 1)
        If tableauTEC(i, 1) <> vbNullString Then
            For j = 1 To UBound(tableauTEC, 2)
                If j <> 1 Then
                    ws.Cells(i + rOffset, j + cOffset).Value = CCur(tableauTEC(i, j))
                Else
                    ws.Cells(i + rOffset, j + cOffset).Value = tableauTEC(i, j)
                End If
            Next j
            totHres = totHres + tableauTEC(i, 2)
            totValeur = totValeur + tableauTEC(i, 4)
            lastRowUsed = i + rOffset
            nbItemTEC = nbItemTEC + 1
        End If
    Next i
    
    If nbItemTEC > 0 Then
        With ws.Range("F" & lastRowUsed).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With ws.Range("H" & lastRowUsed).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    
        'Totaux du premier tableau
        lastRowUsed = lastRowUsed + 1
        ws.Cells(lastRowUsed, 2 + cOffset).Value = totHres
        ws.Range("F" & lastRowUsed).Font.Bold = True
        ws.Cells(lastRowUsed, 4 + cOffset).Value = totValeur
        ws.Range("H" & lastRowUsed).Font.Bold = True
        
        'Mise en forme du premier tableau
        ws.Range("E9:E" & lastRowUsed).HorizontalAlignment = xlCenter
        ws.Range("F9:H" & lastRowUsed).HorizontalAlignment = xlRight
        ws.Range("F9:F" & lastRowUsed).NumberFormat = "##0.00"
        ws.Range("G9:H" & lastRowUsed).NumberFormat = "###,##0.00 $"
        
        'S'assurer que les valeurs sont de vrais valeurs numériques ???
        Dim rng As Range
        Set rng = ws.Range("F9:F" & lastRowUsed)
        Call ConvertirEnNumerique(rng)
        Set rng = ws.Range("G9:G" & lastRowUsed - 1)
        Call ConvertirEnNumerique(rng)
        Set rng = ws.Range("H9:H" & lastRowUsed)
        Call ConvertirEnNumerique(rng)
        
        rOffset = lastRowUsed + 2
    Else
        MsgBox "Je n'ai AUCUNE information sur les TEC" & _
                vbNewLine & vbNewLine & "Pour cette facture", _
                vbOKOnly, "Facture '" & vbNullString & invNo & "'"
    End If
    
    With ws.Range("D" & rOffset)
        .Value = "Honoraires"
        .Font.Italic = True
        .Font.Bold = True
        .Font.size = 11
    End With
    
    'Remplir le tableau HONORAIRES (9 x 4)
    totHres = 0
    totValeur = 0
    Dim premiereLigne As Integer, nbItemHono As Long
    premiereLigne = lastRowUsed + 3
    
    For i = LBound(tableauHonoraires, 1) To UBound(tableauHonoraires, 1)
        If tableauHonoraires(i, 1) <> vbNullString Then
            For j = 1 To UBound(tableauHonoraires, 2)
                ws.Cells(i + rOffset, j + cOffset).Value = tableauHonoraires(i, j)
            Next j
            lastRowUsed = i + rOffset
            nbItemHono = nbItemHono + 1
            totHres = totHres + CCur(tableauHonoraires(i, 2))
            totValeur = totValeur + CCur(tableauHonoraires(i, 4))
        End If
    Next i

    If nbItemHono > 0 Then
        With ws.Range("F" & lastRowUsed).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With ws.Range("H" & lastRowUsed).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        'Totaux du deuxième tableau
        lastRowUsed = lastRowUsed + 1
        ws.Cells(lastRowUsed, 2 + cOffset).Value = totHres
        ws.Range("F" & lastRowUsed).Font.Bold = True
        ws.Cells(lastRowUsed, 4 + cOffset).Value = totValeur
        ws.Range("H" & lastRowUsed).Font.Bold = True
        
        'Mise en forme du deuxième tableau
        ws.Range("E" & premiereLigne & ":E" & lastRowUsed).HorizontalAlignment = xlCenter
        ws.Range("F" & premiereLigne & ":H" & lastRowUsed).HorizontalAlignment = xlRight
        ws.Range("F" & premiereLigne & ":F" & lastRowUsed).NumberFormat = "##0.00"
        ws.Range("G" & premiereLigne & ":G" & lastRowUsed).NumberFormat = "###,##0.00 $"
        ws.Range("H" & premiereLigne & ":H" & lastRowUsed).NumberFormat = "###,##0.00 $"
        
        'S'assurer que les valeurs sont de vrais valeurs numériques ???
        Set rng = ws.Range("F" & premiereLigne & ":F" & lastRowUsed)
        Call ConvertirEnNumerique(rng)
        Set rng = ws.Range("G" & premiereLigne & ":G" & lastRowUsed - 1)
        Call ConvertirEnNumerique(rng)
        Set rng = ws.Range("H" & premiereLigne & ":H" & lastRowUsed)
        Call ConvertirEnNumerique(rng)
    Else
        MsgBox "Je n'ai AUCUNE information sur les honoraires" & _
        vbNewLine & vbNewLine & "Pour cette facture", _
        vbOKOnly, "Facture '" & vbNullString & invNo & "'"
    End If
    
    'Rien d'imprimé
    If nbItemTEC = 0 And nbItemHono = 0 Then
        Call RetourFeuilleSelection_Stats
        Exit Sub
    End If
    
    ws.Range("B1:J" & lastRowUsed + 1).VerticalAlignment = xlCenter
    ws.Columns("B").ColumnWidth = 5
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 11
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F:H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J").ColumnWidth = 5
    
    lastRowUsed = lastRowUsed + 2
    
    'Couleur de fond de feuille
    With ws.Range("B1:J" & lastRowUsed + 3)
        .Interior.Color = gCOULEUR_BASE_FACTURATION
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
    End With
    
   'Bordure blanche
   With ActiveSheet.Range("C3:I" & lastRowUsed - 1)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End With
    
    'Ajouter les boutons Imprimer et Retour
    Dim rngImprimer As Range
    Dim rngRetour As Range
    Set rngImprimer = ws.Range("C" & lastRowUsed + 1)
    Set rngRetour = ws.Range("H" & lastRowUsed + 1)
    Call AjouterBoutons_Stats(ws, wsSelection, rngImprimer, rngRetour)
    
    'Afficher la feuille nouvellement créée
    ws.Activate
    
    'Libérer la mémoire
    Set rng = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:AfficherNouvelleFeuille_Stats", vbNullString, startTime)

End Sub

Function ObtenirTableauTEC(numeroFacture As String) As Variant

    Dim tableauTEC(1 To 9, 1 To 4) As String
    
    Dim ws As Worksheet: Set ws = wsdFAC_Details
    
    Dim hresTEC As Currency
    
    hresTEC = Fn_Get_TEC_Total_Invoice_AF(numeroFacture, "Heures")
    
    'Utilisation du AF#1 généré dans la procédure précédente
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "K").End(xlUp).Row
    If lastUsedRow > 2 Then
        Dim r As Integer
        Dim indice As Integer
        Dim hres As Currency, taux As Currency, valeur As Currency
        Dim prof As String
        For r = 2 To lastUsedRow
            If InStr(ws.Cells(r, 11).Value, "*** - [Sommaire des TEC] pour la facture") = 1 Then
                'Exemple de remplissage du tableau TEC avec des données en fonction du numéro de facture
                indice = indice + 1
                prof = Trim$(Mid$(ws.Cells(r, 11).Value, 44))
                hres = ws.Cells(r, 12).Value
                taux = ws.Cells(r, 13).Value
                valeur = ws.Cells(r, 14).Value
                tableauTEC(indice, 1) = prof
                tableauTEC(indice, 2) = hres
                tableauTEC(indice, 3) = taux
                tableauTEC(indice, 4) = valeur
            End If
        Next r
    End If
    
    ObtenirTableauTEC = tableauTEC
    
End Function

Function ObtenirTableauHonoraires(numeroFacture As String) As Variant

    Dim tableauHonoraires(1 To 9, 1 To 4) As String
    
    Dim wsFees As Worksheet: Set wsFees = wsdFAC_Sommaire_Taux
    
    'Déterminer la dernière ligne utilisée
    Dim lastUsedRow As Long
    lastUsedRow = wsFees.Cells(wsFees.Rows.count, 1).End(xlUp).Row
    
    'Création d'une plage qui contient toutes les lignes pour la facture
    Dim cell As Range
    Set cell = wsFees.Range("A2:A" & lastUsedRow).Find(What:=numeroFacture, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Avons-nous trouvé quelque chose ?
    Dim firstAddress As String
    Dim indice As Integer
    Dim prof As String
    Dim hres As Currency, taux As Currency
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Application.EnableEvents = False
        Do
            'Lire les données
            If wsFees.Cells(cell.row, 4).Value <> 0 Then
                indice = indice + 1
                prof = wsFees.Cells(cell.row, 3).Value
                hres = wsFees.Cells(cell.row, 4).Value
                taux = wsFees.Cells(cell.row, 5).Value
                tableauHonoraires(indice, 1) = prof
                tableauHonoraires(indice, 2) = hres
                tableauHonoraires(indice, 3) = taux
                tableauHonoraires(indice, 4) = Round(hres * taux, 2)
            End If
            'On passe à la ligne suivante de la plage
            Set cell = wsFees.Range("A2:A" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set wsFees = Nothing
    
    ObtenirTableauHonoraires = tableauHonoraires
    
End Function

Function ObtenirTransCC(numeroFacture As String) As Variant

    Dim tableauCC(1 To 100, 1 To 7) As String
    
    'Feuilles nécessaires
    Dim wsFactures As Worksheet: Set wsFactures = wsdFAC_Comptes_Clients
    Dim wsPaiements As Worksheet: Set wsPaiements = wsdENC_Details
    Dim wsRégularisations As Worksheet: Set wsRégularisations = wsdCC_Regularisations
    
    'Obtenir les informations sur la facture (wshComptes_Clients)
    Dim ligneFacture As Long
    ligneFacture = TrouverLigneFacture(wsFactures, numeroFacture)
    Dim montantFacture As Currency
    Dim dateFacture As Date, dateDue As Date
    dateFacture = Format$(wsFactures.Cells(ligneFacture, fFacCCInvoiceDate).Value, wsdADMIN.Range("B1").Value)
    montantFacture = wsFactures.Cells(ligneFacture, fFacCCTotal).Value
    dateDue = wsFactures.Cells(ligneFacture, fFacCCDueDate).Value
    
    'Obtenir les paiements et régularisations pour cette facture
    Dim montantPaye As Currency, montantRegul As Currency, montantRestant As Currency
    montantPaye = Fn_Obtenir_Paiements_Facture(numeroFacture, #12/31/2999#)
    montantRegul = Fn_Obtenir_Régularisations_Facture(numeroFacture, #12/31/2999#)
    
    montantRestant = montantFacture - montantPaye + montantRegul
    
    'Date actuelle pour le calcul de l'âge des factures
    Dim dateAujourdhui As Date
    dateAujourdhui = Date
    
    'Calcul de l'âge de la facture et de la tranche d'âge
    Dim ageFacture As Long
    ageFacture = WorksheetFunction.Max(dateAujourdhui - dateDue, 0)
           
    Dim trancheAge As Integer
    Select Case ageFacture
        Case 0 To 30
            trancheAge = 1
        Case 31 To 60
            trancheAge = 2
        Case 61 To 90
            trancheAge = 3
        Case Is > 90
            trancheAge = 4
        Case Else
            trancheAge = 5
    End Select
    
    Dim i As Integer
    i = i + 1
    tableauCC(i, 1) = "Facture"
    tableauCC(i, 2) = CDbl(dateFacture)
    tableauCC(i, 3) = montantFacture
    tableauCC(i, 3 + trancheAge) = montantRestant
    
    'Obtenir tous les paiements pour la facture
    Dim rngPaiementsAssoc As Range
    Dim pmtFirstAddress As String
    Set rngPaiementsAssoc = wsPaiements.Range("B:B").Find(numeroFacture, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngPaiementsAssoc Is Nothing Then
        pmtFirstAddress = rngPaiementsAssoc.Address
        Do
            i = i + 1
            tableauCC(i, 1) = "Paiement"
            tableauCC(i, 2) = CDbl(rngPaiementsAssoc.offset(0, 2).Value)
            tableauCC(i, 3) = -rngPaiementsAssoc.offset(0, 3).Value 'Montant du paiement
            Set rngPaiementsAssoc = wsPaiements.Columns("B:B").FindNext(rngPaiementsAssoc)
        Loop While Not rngPaiementsAssoc Is Nothing And rngPaiementsAssoc.Address <> pmtFirstAddress
    End If
    
    'Obtenir toutes les régularisations pour la facture
    Dim rngRégularisationAssoc As Range
    Dim regulFirstAddress As String
    Set rngRégularisationAssoc = wsRégularisations.Range("B:B").Find(numeroFacture, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngRégularisationAssoc Is Nothing Then
        regulFirstAddress = rngRégularisationAssoc.Address
        Do
            i = i + 1
            tableauCC(i, 1) = "Régularisation"
            tableauCC(i, 2) = Format$(rngRégularisationAssoc.offset(0, 1).Value, wsdADMIN.Range("B1").Value)
            tableauCC(i, 3) = rngRégularisationAssoc.offset(0, 4).Value + _
                rngRégularisationAssoc.offset(0, 5).Value + _
                rngRégularisationAssoc.offset(0, 6).Value + _
                rngRégularisationAssoc.offset(0, 7).Value
            Set rngRégularisationAssoc = wsRégularisations.Columns("B:B").FindNext(rngRégularisationAssoc)
        Loop While Not rngRégularisationAssoc Is Nothing And rngRégularisationAssoc.Address <> regulFirstAddress
    End If
    
    ObtenirTransCC = tableauCC
    
End Function

Sub AjouterBoutons_Stats(ws As Worksheet, wsSelection As Worksheet, rngImprimer As Range, rngRetour As Range)
    Dim btnImprimer As Shape
    Dim btnRetour As Shape
    
    ' Ajouter un bouton pour imprimer à la position de rngImprimer
    Set btnImprimer = ws.Shapes.AddFormControl(xlButtonControl, rngImprimer.Left, rngImprimer.Top, 103, 30)
    With btnImprimer
        .Name = "btnImprimer"
        .TextFrame.Characters.text = "Imprimer"
        .TextFrame.Characters.Font.size = 14
        .TextFrame.Characters.Font.Bold = True
        .OnAction = "ImprimerStatistiques"
    End With
    
    ' Ajouter un bouton pour retourner à la feuille de sélection à la position de rngRetour
    Set btnRetour = ws.Shapes.AddFormControl(xlButtonControl, rngRetour.Left, rngRetour.Top, 103, 30)
    With btnRetour
        .Name = "btnRetour"
        .TextFrame.Characters.text = "Retour"
        .TextFrame.Characters.Font.size = 14
        .TextFrame.Characters.Font.Bold = True
        .OnAction = "RetourFeuilleSelection_Stats"
    End With
    
End Sub

Sub AjouterBoutons_CC(ws As Worksheet, wsSelection As Worksheet, rngImprimer As Range, rngRetour As Range)

    Dim btnImprimer As Shape
    Dim btnRetour As Shape
    
    ' Ajouter un bouton pour imprimer à la position de rngImprimer
    Set btnImprimer = ws.Shapes.AddFormControl(xlButtonControl, rngImprimer.Left, rngImprimer.Top, 103, 30)
    With btnImprimer
        .Name = "btnImprimer"
        .TextFrame.Characters.text = "Imprimer"
        .TextFrame.Characters.Font.size = 14
        .TextFrame.Characters.Font.Bold = True
        .OnAction = "ImprimerComptesClients"
    End With
    
    ' Ajouter un bouton pour retourner à la feuille de sélection à la position de rngRetour
    Set btnRetour = ws.Shapes.AddFormControl(xlButtonControl, rngRetour.Left, rngRetour.Top, 103, 30)
    With btnRetour
        .Name = "btnRetour"
        .TextFrame.Characters.text = "Retour"
        .TextFrame.Characters.Font.size = 14
        .TextFrame.Characters.Font.Bold = True
        .OnAction = "RetourFeuilleSelection_CC"
    End With
    
End Sub

Sub ImprimerStatistiques()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row
    lastUsedRow = lastUsedRow + 6
    Dim plage As Range
    Set plage = ws.Range("B1:J" & lastUsedRow)
    
    Call ImprimerInfosTECetHonoraires(plage)

End Sub

Sub ImprimerComptesClients()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
    lastUsedRow = lastUsedRow + 5
    Dim plage As Range
    Set plage = ws.Range("A1:K" & lastUsedRow)
    
    Call ImprimerTransCC(plage)

End Sub

Sub ImprimerInfosTECetHonoraires(ByVal plage As Range)

    With plage.Worksheet.PageSetup
        .PrintArea = plage.Address
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .CenterVertically = True
        .LeftFooter = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
    End With

    'Afficher un aperçu avant impression
    plage.Worksheet.PrintPreview
    
End Sub

Sub ImprimerTransCC(ByVal plage As Range)

    With plage.Worksheet.PageSetup
        .PrintArea = plage.Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .CenterVertically = True
        .LeftFooter = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
    End With

    'Afficher un aperçu avant impression
    plage.Worksheet.PrintPreview
    
End Sub

Sub RetourFeuilleSelection_Stats()

    Call SupprimerFeuillesFactureInfo
    
    Dim wsSelection As Worksheet
    Set wsSelection = wshFAC_Interrogation
    wsSelection.Activate
    
End Sub

Sub RetourFeuilleSelection_CC()

    Call SupprimerFeuillesFactureCC
    
    Dim wsSelection As Worksheet
    Set wsSelection = wshFAC_Interrogation
    wsSelection.Activate
    
End Sub

Sub SupprimerFeuillesFactureInfo()

    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 12) = "FactureInfo_" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
End Sub

Sub SupprimerFeuillesFactureCC()

    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 10) = "FactureCC_" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
End Sub

Sub ObtenirListeTECFactures(adresse As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:ObtenirListeTECFactures", vbNullString, 0)
    
    Dim numeroLigne As Long, numeroColonne As Long
    Call ExtraireLigneColonneCellule(adresse, numeroLigne, numeroColonne)

    Dim ws As Worksheet: Set ws = wshFAC_Interrogation
    
    'The invoice number is in column C (3rd column)
    Dim invNo As String
    invNo = ws.Cells(numeroLigne, 3).Value
    
    'Nom du client et date de facture
    Dim nomClient As String
    nomClient = ws.Range("D4").Value
    Dim dateFacture As Date
    dateFacture = ws.Cells(numeroLigne, 4).Value
    
    'Utilisation d'un AdvancedFilter directement dans TEC_Local (BI:BX)
    Call ObtenirListeTECFacturésFiltreAvancé(invNo)

    Set ws = wsdTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).Row
    
    'Est-ce que nous avons des TEC pour cette facture ?
    If lastUsedRow < 3 Then
        MsgBox "Il n'y a aucun TEC associé à la facture '" & invNo & "'"
    Else
        Call PreparerRapportTECFacturés(invNo)
    End If
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:ObtenirListeTECFactures", vbNullString, startTime)
    
End Sub

Sub PreparerRapportTECFactures()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:PreparerRapportTECFactures", vbNullString, 0)
    
    'Assigner la feuille du rapport
    Dim strRapport As String
    strRapport = "Rapport TEC facturés"
    Dim wsRapport As Worksheet: Set wsRapport = wshTECFacture
    wsRapport.Cells.Clear
    
    'Désactiver les mises à jour de l'écran et autres alertes
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    'Mettre en forme la feuille de rapport
    With wsRapport
        ' Titre du rapport
        .Range("A1").Value = "TEC facturés pour la facture '" & invNo & "'"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.size = 12
        
        'Ajouter une date de génération du rapport
        .Range("A2").Value = "Date de création : " & Format$(Date, wsdADMIN.Range("B1").Value)
        .Range("A2").Font.Italic = True
        .Range("A2").Font.size = 10
        
        'Entête du rapport (A4:D4)
        .Range("A4").Value = "Date"
        .Range("B4").Value = "Prof."
        .Range("C4").Value = "Description"
        .Range("D4").Value = "Heures"
        With .Range("A4:D4")
            .Font.size = 9
            .Font.Bold = True
            .Font.Italic = True
            .Font.Color = vbWhite
            .HorizontalAlignment = xlCenter
        End With
        
        'Utilisation du AdvancedFilter # 3 sur la feuille TEC_Local
        Dim wsSource As Worksheet
        Set wsSource = wsdTEC_Local 'Utilisation des résultats du AF (BJ:BY)
        
        'Copier quelques données de la source
        Dim rngResult As Range
        Set rngResult = wsSource.Range("BJ1").CurrentRegion.offset(2, 0)
        'Redimensionner la plage après l'offset pour avoir que les données (pas d'entête)
        Set rngResult = rngResult.Resize(rngResult.Rows.count - 2)
        'Transfert des données vers un tableau
        Dim tableau As Variant
        tableau = rngResult.Value
        
        'Créer un tableau pour les résultats
        Dim output() As Variant
        ReDim output(1 To UBound(tableau, 1), 1 To 4)
        Dim r As Long
        
        Dim i As Long
        For i = LBound(tableau, 1) To UBound(tableau, 1)
            r = r + 1
            output(r, 1) = tableau(i, 4)
            output(r, 2) = tableau(i, 3)
            output(r, 3) = tableau(i, 7)
            output(r, 4) = tableau(i, 8)
        Next i

        'Copier le tableau dans la feuille du rapport  partir de la ligne 5, colonne 1
        .Range(.Cells(5, 1), .Cells(5 + UBound(output, 1) - 1, 1 + UBound(output, 2) - 1)).Value = output
        'Ligne dans la feuille du rapport
        r = 5 + UBound(output, 1) - 1
        
        'Corps du rapport
        .Range("A5:D" & r).VerticalAlignment = xlCenter
        With .Range("A4:D4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        'Ajouter une bordure aux données
        .Range("A4:D" & r).Borders.LineStyle = xlContinuous
        With .Range("A5:D" & r).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Range("A5:D" & r).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        
        .Range("A4:D" & r).Font.Name = "Aptos Narrow"
        .Range("A4:D" & r).Font.size = 10
        
        .Columns("A").ColumnWidth = 10
        .Range("A4:A" & r).HorizontalAlignment = xlCenter
        
        .Columns("B").ColumnWidth = 6
        .Range("B4:B" & r).HorizontalAlignment = xlCenter
        
        .Columns("C").ColumnWidth = 72
        .Columns("C").WrapText = True
        
        .Columns("D").ColumnWidth = 7
        .Columns("D").NumberFormat = "##0.00"
        
    End With

    'Configurer la mise en page pour l'impression ou l'export en PDF
    With wsRapport.PageSetup
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        
        'Ajuster la marge des en-têtes et pieds de page (1 cm)
        .HeaderMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(1)
        
        .Orientation = xlPortrait 'Portrait
        .FitToPagesWide = 1 'Ajuster sur une page en largeur
        .FitToPagesTall = False ' Ne pas ajuster en hauteur
        .PrintArea = "A1:D" & r ' Définir la zone d'impression
        .CenterHorizontally = True ' Centrer horizontalement
        .CenterVertically = False ' Centrer verticalement
    End With
    
    'On se déplace à la feuille contenant le rapport
    wsRapport.Visible = xlSheetVisible
    wsRapport.Activate
    
    MsgBox "Le rapport a été généré sur la feuille " & strRapport
    
    'Libérer la mémoire
    Set rngResult = Nothing
    Set wsRapport = Nothing
    Set wsSource = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:PreparerRapportTECFactures", vbNullString, startTime)
    
End Sub

Sub ObtenirListeTECFacturésFiltreAvancé(noFact As String) '2024-10-20 @ 11:11

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:ObtenirListeTECFacturésFiltreAvancé", vbNullString, 0)

    'Utilisation de la feuille TEC_Local
    Dim ws As Worksheet: Set ws = wsdTEC_Local
    
    'wsdTEC_Local_AF#3
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'AdvancedFilter par Numéro de Facture
    
    'Effacer les données de la dernière utilisation
    ws.Range("BH6:BH10").ClearContents
    ws.Range("BH6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_TEC_Local[#All]")
    ws.Range("BH7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("BH2:BH3")
    ws.Range("BH3").Value = CStr(noFact)
    ws.Range("BH8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("BJ1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("BJ2:BY2")
    ws.Range("BH9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Qu'avons-nous comme résultat ?
    Dim lastResultRow As Long
    lastResultRow = ws.Cells(ws.Rows.count, "BJ").End(xlUp).Row
    ws.Range("BH10").Value = lastResultRow - 2 & " lignes"
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 3 Then
        With ws.Sort 'Sort - Date, ProfID, TECID
            .SortFields.Clear
            'First sort On Date
            .SortFields.Add key:=ws.Range("BM3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On ProfID
            .SortFields.Add key:=ws.Range("BK3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.Add key:=ws.Range("BJ3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange ws.Range("BJ3:BY" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Free memory
    Set rngData = Nothing
    Set rngCriteria = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:ObtenirListeTECFacturésFiltreAvancé", vbNullString, startTime)
    
End Sub

Sub shpRetournerInterrogationFacture_Click()

    Call RetournerInterrogationFacture

End Sub

Sub RetournerInterrogationFacture()

    Application.ScreenUpdating = False
    wshFAC_Interrogation.Activate
    Application.ScreenUpdating = True
    
End Sub

Sub ObtenirTransactionsCC(invNo As String, nomClient As String, dateFacture As Date)

    Call AfficherNouvelleFeuille_CC(invNo, nomClient, dateFacture)

End Sub

Sub AfficherNouvelleFeuille_CC(invNo As String, nomClient As String, dateFacture As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:AfficherNouvelleFeuille_CC", vbNullString, 0)
    
    If invNo = vbNullString Then
        Exit Sub
    End If
    
    Dim sheetName As String
    sheetName = "FactureCC_" & invNo
    
    ' Référence à la première feuille
    Dim wsSelection As Worksheet
    Set wsSelection = wshFAC_Interrogation
    
    'Vérifier si la feuille existe déjà
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    'Si la feuille existe, la supprimer
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Créer une nouvelle feuille
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = sheetName
    
    'Entête de la feuille
    ws.Range("A1:K1").Merge
    With ws.Range("A1")
        .Value = "Transactions des Comptes-Clients"
        .Font.size = 22
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("C4").Value = "No. de Facture:"
    ws.Range("C4").HorizontalAlignment = xlRight
    ws.Range("C4").Font.Italic = True
    ws.Range("C4").Font.size = 9
    
    ws.Range("D4").Value = invNo
    ws.Range("D4").HorizontalAlignment = xlCenter
    ws.Range("D4").Font.Bold = True
    ws.Range("D4").Font.size = 12
    
    ws.Range("C6").Value = "Nom du Client:"
    ws.Range("C6").HorizontalAlignment = xlRight
    ws.Range("C6").Font.Italic = True
    ws.Range("C6").Font.size = 9
    
    If Len(nomClient) > 59 Then nomClient = Left$(nomClient, 59) & "..."
    
    With ws.Range("D6")
        .Value = nomClient
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.size = 12
    End With
    
    ws.Range("H4").Value = "Date de Facture:"
    ws.Range("H4").HorizontalAlignment = xlRight
    ws.Range("H4").Font.Italic = True
    ws.Range("H4").Font.size = 9
    
    ws.Range("I4").Value = dateFacture
    ws.Range("I4").HorizontalAlignment = xlCenter
    ws.Range("I4").Font.Bold = True
    ws.Range("I4").Font.size = 12
    
    ws.Range("C8").Value = "Type trans."
    ws.Range("D8").Value = "Date trans."
    ws.Range("E8").Value = "Montant"
    ws.Range("F8").Value = "- 30 jours."
    ws.Range("G8").Value = "31 @ 60 jours"
    ws.Range("H8").Value = "61 @ 90 jours"
    ws.Range("I8").Value = "+ de 90 jours"
    ws.Range("C8:I8").HorizontalAlignment = xlCenter
    ws.Range("C8:I8").Font.Bold = True
    ws.Range("C8:I8").Font.Italic = True
    ws.Range("J4").Font.size = 9
    
    'Obtenir les transactions pour la facture
    Dim tableauCC As Variant
    tableauCC = ObtenirTransCC(invNo)

    'Transférer le tableauCC dans la plage
    Dim rOffset As Integer
    rOffset = 8
    Dim cOffset As Integer
    cOffset = 2
    Dim lastRowUsed As Long
'    Dim totHres As Currency, totValeur As Currency
    Dim i As Integer, j As Integer
    For i = LBound(tableauCC, 1) To UBound(tableauCC, 1)
        If tableauCC(i, 1) <> vbNullString Then
            For j = 1 To UBound(tableauCC, 2)
                If j = 2 Then
                    ws.Cells(i + rOffset, j + cOffset).Value = "'" & Format$(tableauCC(i, j), wsdADMIN.Range("B1").Value)
                Else
                    ws.Cells(i + rOffset, j + cOffset).Value = tableauCC(i, j)
                End If
            Next j
            lastRowUsed = i + rOffset
        Else
            Exit For
        End If
    Next i

    'S'assurer que les valeurs sont de vrais valeurs numériques ???
    Dim rng As Range
    Set rng = ws.Range("E9:I" & lastRowUsed)
    Call ConvertirEnNumerique(rng)
    
    'Effacer les soldes à zéro sur la première ligne
    For i = 6 To 9
        If ws.Cells(9, i).Value = 0 Then
            ws.Cells(9, i).Value = vbNullString
        End If
    Next i

    'Efface l'âge des transactions
    ws.Range("F10:I" & lastRowUsed).Clear
    
    'Mise en forme du tableau
    ws.Range("C9:D" & lastRowUsed).HorizontalAlignment = xlCenter
    ws.Range("E9:I" & lastRowUsed).HorizontalAlignment = xlRight
    ws.Range("E9:I" & lastRowUsed).NumberFormat = "###,##0.00 $"

    ws.Range("A1:K" & lastRowUsed + 1).VerticalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 3
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E:I").ColumnWidth = 15
    ws.Columns("J").ColumnWidth = 3
    ws.Columns("K").ColumnWidth = 5

    lastRowUsed = lastRowUsed + 2

    'Couleur de fond de feuille & cadre extérieur
    With ws.Range("A1:K" & lastRowUsed + 3)
        .Interior.Color = gCOULEUR_BASE_FACTURATION
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
    End With

    'Entête un peu plus foncé
    With ws.Range("C8:I8").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

    'Zone de données sans couleur
    With ws.Range("C9:I" & lastRowUsed - 2).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Bordure blanche
    With ActiveSheet.Range("B3:J" & lastRowUsed - 1)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End With

    'Ajouter les boutons Imprimer et Retour
    Dim rngImprimer As Range
    Dim rngRetour As Range
    Set rngImprimer = ws.Range("B" & lastRowUsed + 1)
    Set rngRetour = ws.Range("I" & lastRowUsed + 1)
    Call AjouterBoutons_CC(ws, wsSelection, rngImprimer, rngRetour)

    'Afficher la feuille nouvellement créée
    ws.Activate
    
    'Libérer la mémoire
    Set rng = Nothing
    Set rngImprimer = Nothing
    Set rngRetour = Nothing
    Set ws = Nothing
    Set wsSelection = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:AfficherNouvelleFeuille_CC", vbNullString, startTime)

End Sub

Sub PreparerRapportTECFacturés(numeroFacture As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:PreparerRapportTECFacturés", vbNullString, 0)
    
    'Assigner la feuille du rapport
    Dim strRapport As String
    strRapport = "Rapport TEC facturés"
    Dim wsRapport As Worksheet: Set wsRapport = wshTECFacture
    On Error Resume Next
    wsRapport.Unprotect
    On Error GoTo 0
    wsRapport.Cells.Clear
    
    'Désactiver les mises à jour de l'écran et autres alertes
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    'Mettre en forme la feuille de rapport
    With wsRapport
        ' Titre du rapport
        .Range("A1").Value = "TEC facturés pour la facture '" & numeroFacture & "'"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.size = 12
        
        'Ajouter une date de génération du rapport
        .Range("A2").Value = "Date de création : " & Format$(Date, "dd/mm/yyyy")
        .Range("A2").Font.Italic = True
        .Range("A2").Font.size = 10
        
        'Entête du rapport (A4:D4)
        .Range("A4").Value = "Date"
        .Range("B4").Value = "Prof."
        .Range("C4").Value = "Description"
        .Range("D4").Value = "Heures"
        With .Range("A4:D4")
            .Font.size = 9
            .Font.Bold = True
            .Font.Italic = True
            .Font.Color = vbWhite
            .HorizontalAlignment = xlCenter
        End With
        
        'Utilisation du AdvancedFilter # 3 sur la feuille TEC_Local
        Dim wsSource As Worksheet
        Set wsSource = wsdTEC_Local 'Utilisation des résultats du AF (BJ:BY)
        
        'Copier quelques données de la source
        Dim rngResult As Range
        Set rngResult = wsSource.Range("BJ1").CurrentRegion.offset(2, 0)
        'Redimensionner la plage après l'offset pour avoir que les données (pas d'entête)
        Set rngResult = rngResult.Resize(rngResult.Rows.count - 2)
        'Transfert des données vers un tableau
        Dim tableau As Variant
        tableau = rngResult.Value
        
        'Créer un tableau pour les résultats
        Dim output() As Variant
        ReDim output(1 To UBound(tableau, 1), 1 To 4)
        Dim r As Long
        
        Dim i As Long
        For i = LBound(tableau, 1) To UBound(tableau, 1)
            r = r + 1
            output(r, 1) = tableau(i, 4)
            output(r, 2) = tableau(i, 3)
            output(r, 3) = tableau(i, 7)
            output(r, 4) = tableau(i, 8)
        Next i

        'Copier le tableau dans la feuille du rapport  partir de la ligne 5, colonne 1
        .Range(.Cells(5, 1), .Cells(5 + UBound(output, 1) - 1, 1 + UBound(output, 2) - 1)).Value = output
        'Ligne dans la feuille du rapport
        r = 5 + UBound(output, 1) - 1
        
        'Corps du rapport
        .Range("A5:D" & r).VerticalAlignment = xlCenter
        With .Range("A4:D4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        'Ajouter une bordure aux données
        .Range("A4:D" & r).Borders.LineStyle = xlContinuous
        With .Range("A5:D" & r).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Range("A5:D" & r).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        
        .Range("A4:D" & r).Font.Name = "Aptos Narrow"
        .Range("A4:D" & r).Font.size = 10
        
        .Columns("A").ColumnWidth = 10
        .Range("A4:A" & r).HorizontalAlignment = xlCenter
        
        .Columns("B").ColumnWidth = 6
        .Range("B4:B" & r).HorizontalAlignment = xlCenter
        
        .Columns("C").ColumnWidth = 72
        .Columns("C").WrapText = True
        
        .Columns("D").ColumnWidth = 7
        .Columns("D").NumberFormat = "##0.00"
        
    End With

    'Configurer la mise en page pour l'impression ou l'export en PDF
    With wsRapport.PageSetup
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        
        'Ajuster la marge des en-têtes et pieds de page (1 cm)
        .HeaderMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(1)
        
        .Orientation = xlPortrait 'Portrait
        .FitToPagesWide = 1 'Ajuster sur une page en largeur
        .FitToPagesTall = False ' Ne pas ajuster en hauteur
        .PrintArea = "A1:D" & r ' Définir la zone d'impression
        .CenterHorizontally = True ' Centrer horizontalement
        .CenterVertically = False ' Centrer verticalement
    End With
    
    'On se déplace à la feuille contenant le rapport
    wsRapport.Visible = xlSheetVisible
    wsRapport.Activate
    
    MsgBox "Le rapport a été généré sur la feuille " & strRapport
    
    'Libérer la mémoire
    Set rngResult = Nothing
    Set wsRapport = Nothing
    Set wsSource = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Interrogation:PreparerRapportTECFacturés", vbNullString, startTime)
    
End Sub

Sub shpExitDetailTEC_Click()

    ActiveSheet.Visible = xlSheetHidden

End Sub


