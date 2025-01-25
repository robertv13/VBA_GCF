Attribute VB_Name = "modFAC_Historique"
Option Explicit

Sub shp_Affiche_Factures_Click()

    Call Affiche_Liste_Factures

End Sub

Sub Affiche_Liste_Factures()

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshFAC_Historique:Affiche_Liste_Factures", 0)
    
    wshFAC_Historique.Range("C9:O33").ClearContents
    
    Call Remove_All_PDF_Icons
    
    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Application.ScreenUpdating = False
    
    Dim clientName As String: clientName = ws.Range("D4").Value
    Dim dateFrom As Date: dateFrom = ws.Range("G6").Value
    Dim dateTo As Date: dateTo = ws.Range("I6").Value
    
    'What is the ID for the selected client ?
    Dim myInfo() As Variant
    Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
    myInfo = Fn_Find_Data_In_A_Range(rng, 1, clientName, fClntFMClientID)
    If myInfo(1) = "" Then
        MsgBox "Je ne peux retrouver ce client dans ma liste de clients", vbCritical
        GoTo Clean_Exit
    End If
    
    Dim codeClient As String
    codeClient = myInfo(3)
    
    Call FAC_Get_Invoice_Client_AF(codeClient)
    
    Call Copy_List_Of_Invoices_to_Worksheet(dateFrom, dateTo)
    
    Application.ScreenUpdating = True
    
    Dim shp As Shape: Set shp = wshFAC_Historique.Shapes("shpAfficheFactures")
    shp.Visible = False
    
    Call Log_Record("wshFAC_Historique:Affiche_Liste_Factures", startTime)

Clean_Exit:
    
    'Libérer la mémoire
    Set rng = Nothing
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub FAC_Get_Invoice_Client_AF(codeClient As String) '2024-06-27 @ 15:27

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'wshFAC_Entête_AF#1

    With ws
    
        'Effacer les données de la dernière utilisation
        .Range("Y6:Y10").ClearContents
        .Range("Y6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
        'Définir le range pour la source des données en utilisant un tableau
        Dim rngData As Range
        Set rngData = .Range("l_tbl_FAC_Entête[#All]")
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
        lastResultRow = .Cells(.Rows.count, "Z").End(xlUp).row
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

Sub Copy_List_Of_Invoices_to_Worksheet(dateMin As Date, dateMax As Date)

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Détermine la dernière utilisée dans les résultats de AF_1 dans wshFAC_Entête
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "Z").End(xlUp).row
    If lastUsedRow < 3 Then Exit Sub 'Nothing to display
    
    Dim arr() As Variant
    ReDim arr(1 To 250, 1 To 13)
    Dim resultArr As Variant
    
    With ws
        Dim i As Long, r As Long
        Debug.Print vbNewLine & "For/Next loop - " & 3 & " to " & lastUsedRow
        For i = 3 To lastUsedRow
            'Vérification de la date de facture -ET- si la facture est bel et bien confirmée
            If .Range("AB" & i).Value >= dateMin And .Range("AB" & i).Value <= dateMax Then
                If .Range("AC" & i).Value = "C" Then
                    r = r + 1
                    arr(r, 1) = .Range("AA" & i).Value  'Invoice number
                    arr(r, 2) = .Range("AB" & i).Value 'Invoice Date
                    arr(r, 3) = .Range("AJ" & i).Value 'Fees
                    arr(r, 4) = .Range("AL" & i).Value 'Misc. 1
                    arr(r, 5) = .Range("AN" & i).Value 'Misc. 2
                    arr(r, 6) = .Range("AP" & i).Value 'Misc. 3
                    arr(r, 7) = .Range("AR" & i).Value 'GST $
                    arr(r, 8) = .Range("AT" & i).Value 'PST $
                    arr(r, 9) = .Range("AV" & i).Value 'Deposit
                    arr(r, 10) = .Range("AU" & i).Value 'AR_Total
                    arr(r, 11) = Fn_Get_Invoice_Total_Payments_AF(.Range("AA" & i).Value)
                    arr(r, 12) = Fn_Get_Invoice_Due_Date(.Range("AA" & i).Value)
                    'Obtenir les TEC facturés par cette facture
                    arr(r, 13) = Fn_Get_TEC_Total_Invoice_AF(.Range("AA" & i).Value, "Dollars")
                End If
            End If
        Next i
    End With
    
    If r = 0 Then
        MsgBox "Il n'y a aucune facture pour la période recherchée", vbExclamation
        GoTo Clean_Exit
    End If
    
    'Transfer the arr to the worksheet, after resizing it
    Call Array_2D_Resizer(arr, r, 13)

    Application.EnableEvents = False
    
    With wshFAC_Historique
        For i = 1 To UBound(arr, 1)
            .Range("C" & i + 8).Value = arr(i, 1)
            .Range("D" & i + 8).Value = Format$(arr(i, 2), wshAdmin.Range("B1").Value)
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
                .Range("N" & i + 8).Value = Format$(WorksheetFunction.Max(0, Now() - arr(i, 12)), "# ###")
            End If
            .Range("O" & i + 8).Value = arr(i, 10) - arr(i, 11) 'Balance
        Next i
    End With
    
    lastUsedRow = i + 8
    Call Remove_All_PDF_Icons
    If lastUsedRow >= 9 Then
        Call Insert_PDF_Icons(lastUsedRow)
    End If
    
    Application.EnableEvents = True

Clean_Exit:

    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub Insert_PDF_Icons(lastUsedRow As Long)

    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Dim i As Long
    Dim iconPath As String
    iconPath = wshAdmin.Range("F5").Value & Application.PathSeparator & "Resources\AdobeAcrobatReader.png"
    
    Dim pic As Picture
    Dim cell As Range
    
    'Loop through each row and insert the icon if there is data in column C
    For i = 9 To lastUsedRow
        If ws.Cells(i, 3).Value <> "" Then
            Set cell = ws.Cells(i, 16) 'Set the cell where the icon should be inserted (column P)
            
            'Insert the icon
            Set pic = ws.Pictures.Insert(iconPath)
'            Debug.Print "#056 - " & pic.width, pic.Height
            With pic
                .Top = cell.Top + 1
                .Left = cell.Left + 3
                .Height = cell.Height - 5
                .Width = cell.Width - 5
                .Placement = xlMoveAndSize
                .OnAction = "Display_PDF_Invoice"
            End With
        End If
    Next i
    
    'Libérer la mémoire
    Set cell = Nothing
    Set pic = Nothing
    Set ws = Nothing
    
End Sub

Sub Display_PDF_Invoice()

    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Dim rowNumber As Long
    Dim fullPDFFileName As String
    
    'Determine which icon was clicked and get the corresponding row number
    Dim targetCell As Range
    Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    rowNumber = targetCell.row
    
    'Assuming the invoice number is in column E (5th column)
    fullPDFFileName = wshAdmin.Range("F5").Value & FACT_PDF_PATH & _
                            Application.PathSeparator & ws.Cells(rowNumber, 3).Value & ".pdf"
    
    'Ouvrir la version PDF de la facture, si elle existe
    If Dir(fullPDFFileName) <> "" Then
        'Le fichier existe, on peut lancer la commande Shell pour l'ouvrir
        Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " & Chr(34) & fullPDFFileName & Chr(34), vbNormalFocus
    Else
        'Le fichier n'existe pas, afficher un message d'erreur
        MsgBox "La version PDF de cette facture n'existe pas" & vbNewLine & vbNewLine & _
                                                        fullPDFFileName, vbExclamation, "Fichier PDF introuvable"
    End If
    
    'Libérer la mémoire
    Set targetCell = Nothing
    Set ws = Nothing
    
End Sub

Sub Remove_All_PDF_Icons() 'RMV - 2024-07-24 @ 19:58

    Dim ws As Worksheet: Set ws = wshFAC_Historique
    
    Dim pic As Picture
    For Each pic In ws.Pictures
        pic.Delete
    Next pic
    
    'Libérer la mémoire
    Set pic = Nothing
    Set ws = Nothing
    
End Sub

Sub FAC_Historique_Clear_All_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Historique:FAC_Historique_Clear_All_Cells", 0)
    
    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    ActiveSheet.Unprotect
    With wshFAC_Historique
        .Range("D4:H4, D6:E6").ClearContents
        .Range("G6, I6").ClearContents
        .Range("C9:R33").ClearContents
        Call Remove_All_PDF_Icons
        Application.EnableEvents = True
        wshFAC_Historique.Activate
        wshFAC_Historique.Range("D4").Select
    End With
    
    With ActiveSheet
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

    Call Log_Record("modFAC_Historique:FAC_Historique_Clear_All_Cells", startTime)

End Sub

Sub shp_FAC_Historique_Exit_Click()

    Call FAC_Historique_Back_To_FAC_Menu

End Sub

Sub FAC_Historique_Back_To_FAC_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Historique:FAC_Historique_Back_To_FAC_Menu", 0)
    
    wshFAC_Historique.Visible = xlSheetHidden
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call Log_Record("modFAC_Historique:FAC_Historique_Back_To_FAC_Menu", startTime)

End Sub

Sub FAC_Historique_Montrer_Bouton()

    Dim shp As Shape: Set shp = wshFAC_Historique.Shapes("shpAfficheFactures")
    
    Application.EnableEvents = False
    
    If IsDate(wshFAC_Historique.Range("G6").Value) And _
        IsDate(wshFAC_Historique.Range("I6").Value) And _
        Trim(wshFAC_Historique.Range("D4").Value) <> "" Then
        shp.Visible = True
    Else
        shp.Visible = False
    End If
    
    Application.EnableEvents = True

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub
