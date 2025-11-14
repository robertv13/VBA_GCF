Attribute VB_Name = "modFAC_Brouillon"
'@Folder("Saisie_Facture")

Option Explicit

Private invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Private lastRow As Long, lastResultRow As Long, resultRow As Long

Sub CreerNouvelleFactureBrouillon() 'Clear contents
    
    Dim startTime As Double: startTime = Timer
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:CreerNouvelleFactureBrouillon", vbNullString, 0)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Masquer la forme (détail TEC) si elle est présente
    On Error Resume Next
    Dim shapeTextBox As Shape
    Set shapeTextBox = wshFAC_Brouillon.Shapes("shpTECInfo")
    If Not shapeTextBox Is Nothing Then
        shapeTextBox.Visible = msoFalse
    End If
    On Error GoTo 0
    
    'Are we entering a NEW invoice ?
    If wshFAC_Brouillon.Range("B27").Value = False Then
    
        With wshFAC_Brouillon
            .Range("B5").Value = "FAUX"
            .Range("B24").Value = True
            .Range("K3:L7,O3,O5").ClearContents 'Clear cells for a new Invoice
            .Range("O6").Value = Fn_ProchainNumeroFacture
            
            On Error Resume Next
                DoEvents
                Unload ufFraisDivers
                Unload ufNonBillableTime '2025-01-14 @ 17:25
            On Error GoTo 0
            
            Call EffacerTECAffiches
            
            Call MettreEnPlaceCellulesFACBrouillon
            
            Application.EnableEvents = False
            .Range("B20").Value = vbNullString
            .Range("B24").Value = False
            .Range("B26").Value = False
            .Range("B51, B52, B53, B54").Value = vbNullString 'Requests for invoice
            .Range("R44:T48").ClearContents 'Hours & Fees summary from request for invoice
            Application.EnableEvents = True
        End With
        
        With wshFAC_Finale
            Application.EnableEvents = False
            .Range("B21,B23:C27,E28").ClearContents
            .Range("A34:F68").ClearContents
            .Range("E28").Value = wshFAC_Brouillon.Range("O6").Value 'Invoice #
            .Range("B69:F81").ClearContents 'NOT the formulas
            .Range("L79").Value = vbNullString
            .Range("L81").Value = vbNullString
            Application.EnableEvents = True
            
            Call MettreEnPlaceToutesLesCellules
        
        End With
        
        Application.EnableEvents = False
        wshFAC_Brouillon.Range("B16").Value = False 'Does not see billed charges
        Application.EnableEvents = True
        
        Application.ScreenUpdating = True
        
        'Ensure all pending events could be processed
        DoEvents

        'Save button is disabled UNTIL the invoice is saved
        Call CacherBoutonSauvegarder
        
        gFlagEtapeFacture = 0
    
        'Ensure all pending events could be processed
        DoEvents

        'Do we have pending requests to invoice ?
        Dim lo As ListObject '2025-06-01 @ 06:07
        Set lo = wsdFAC_Projets_Entete.ListObjects("l_tbl_FAC_Projets_Entete")
        
        Dim liveOne As Long
        Dim i As Long
        
        If Fn_TableauContientDesDonnees(lo) Then
            Dim nomColonne As ListColumn
            Set nomColonne = lo.ListColumns("estDetruite")
        
            For i = 1 To lo.ListRows.count
                Dim val As Variant
'                If lo.Range(i, 1).Value <> "" Then 'Ligne 1 du data est TOUJOURS vide (ADO vs Tableau structuré) '2025-06-30 @ 09:18
                    val = nomColonne.DataBodyRange.Cells(i, 1).Value
                    If UCase$(val) = "FAUX" Or val = 0 Then
                        liveOne = liveOne + 1
                    End If
'                End If
            Next i
        End If
        
        'Bring the visible area to the top
        wshFAC_Brouillon.Range("E3").Select

        If liveOne Then
            ufListeProjetsFacture.show
        End If
        
        Dim projetID As Long
        If wshFAC_Brouillon.Range("B51").Value <> vbNullString Then
            Application.EnableEvents = False
            projetID = CLng(wshFAC_Brouillon.Range("B52").Value)
            'Obtenir l'entête pour ce projet de facture
            Dim lastUsedRow As Long
            lastUsedRow = wsdFAC_Projets_Entete.Cells(wsdFAC_Projets_Entete.Rows.count, 1).End(xlUp).Row
            Dim rngToSearch As Range: Set rngToSearch = wsdFAC_Projets_Entete.Range("A1:A" & lastUsedRow)
            Dim result As Variant
            result = Application.WorksheetFunction.XLookup(projetID, _
                                                           rngToSearch, _
                                                           rngToSearch, _
                                                           "Not Found", _
                                                           0, _
                                                           1)

            If result <> "Not Found" Then
                Dim matchedRow As Long
                matchedRow = Application.Match(projetID, rngToSearch, 0)
                Dim arr() As Variant
                ReDim arr(1 To 5, 1 To 3)
                Dim ii As Long
                For ii = 1 To 5
                    arr(ii, 1) = wsdFAC_Projets_Entete.Cells(matchedRow, (ii - 1) * 4 + 6).Value
                    arr(ii, 2) = wsdFAC_Projets_Entete.Cells(matchedRow, (ii - 1) * 4 + 7).Value
                    arr(ii, 3) = wsdFAC_Projets_Entete.Cells(matchedRow, (ii - 1) * 4 + 8).Value
                Next ii
                'Update the summary for billing
                'Transfer data to the worksheet
                Application.EnableEvents = False
                Dim r As Long: r = 44
                For ii = 44 To 48
                    If arr(ii - 43, 1) <> vbNullString And arr(ii - 43, 2) <> 0 Then
                        wshFAC_Brouillon.Range("R" & r).Value = arr(ii - 43, 1)
                        wshFAC_Brouillon.Range("S" & r).Value = arr(ii - 43, 2)
                        If wshFAC_Brouillon.Range("S" & r).Value <> 0 Then
                            With wshFAC_Brouillon.Range("S" & r).Interior
                                .Pattern = xlNone
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                        wshFAC_Brouillon.Range("S" & r).NumberFormat = "#,##0.00"
                        wshFAC_Brouillon.Range("T" & r).Value = arr(ii - 43, 3)
                        wshFAC_Brouillon.Range("T" & r).NumberFormat = "#,##0.00 $"
                        Dim s As String
                        s = "=if(R" & r & "<>"""", S" & r & "*T" & r & ")"
                        wshFAC_Brouillon.Range("U" & r).formula = "=if(R" & r & "<>"""", S" & r & "*T" & r & ")"
                        r = r + 1
                   End If
                Next ii
            End If
            
            'Calcul du total des heures & des honoraires
            wshFAC_Brouillon.Range("S49").formula = "=sum(S44:S48)"
            wshFAC_Brouillon.Range("U49").formula = "=sum(U44:U48)"
            
            'The total fees amount id determined by the fees summary
            wshFAC_Brouillon.Range("O47").Value = wshFAC_Brouillon.Range("U49").Value
            
            wshFAC_Brouillon.Range("E3").Value = wshFAC_Brouillon.Range("B51").Value
            Call ChangerNomDuClient(wshFAC_Brouillon.Range("B51").Value)
            
            Application.EnableEvents = False
            
            'Utilisation de la date du projet de facture
            wshFAC_Brouillon.Range("O3").Value = Format$(Date, wsdADMIN.Range("B1").Value)
            Call FACBrouillonDate_Change(wshFAC_Brouillon.Range("O3").Value)
            
            wshFAC_Brouillon.Range("L11").Select '2025-11-01 @ 14:04
'            wshFAC_Brouillon.Range("O9").Select
            
            Application.EnableEvents = True
        Else
            Application.EnableEvents = True
            wshFAC_Brouillon.Select
            wshFAC_Brouillon.Range("E3").Select 'Start inputing values for a NEW invoice
        End If
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set rngToSearch = Nothing
    Set shapeTextBox = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:CreerNouvelleFactureBrouillon", vbNullString, startTime)

End Sub

Sub ChangerNomDuClient(clientName As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ChangerNomDuClient", clientName, 0)
    
    'Aller chercher le vrai nom de client de 2 sources selon le mode de facturation
    Dim allCols As Variant
    If wshFAC_Brouillon.Range("B52").Value = vbNullString Then
        allCols = Fn_ObtenirLigneDeFeuille("BD_Clients", clientName, fClntFMNomClientPlusNomClientSystème)
    Else
        allCols = Fn_ObtenirLigneDeFeuille("BD_Clients", clientName, fClntFMClientNom)
    End If
    
    'Vérifier le résultat retourné
    If IsArray(allCols) Then
        Application.EnableEvents = False
        clientName = allCols(1)
        wshFAC_Brouillon.Range("E3").Value = clientName
        Application.EnableEvents = True
    Else
        wshFAC_Brouillon.Range("E3").Value = vbNullString
        MsgBox "Valeur non trouvée !!!", vbCritical
        wshFAC_Brouillon.Range("E3").Select
    End If
    
    Dim clientNamePurged As String
    clientNamePurged = clientName
    Do While InStr(clientNamePurged, "[") > 0 And InStr(clientNamePurged, "]") > 0
        clientNamePurged = Fn_Strip_Contact_From_Client_Name(clientNamePurged)
    Loop
        
    Application.EnableEvents = False
    wshFAC_Brouillon.Range("B18").Value = allCols(fClntFMClientID)
    Application.EnableEvents = True
    
    With wshFAC_Brouillon
        Application.EnableEvents = False
        .Range("K3").Value = allCols(fClntFMContactFacturation)
        .Range("K4").Value = clientNamePurged
        .Range("K5").Value = allCols(fClntFMAdresse1) 'Adresse1
        If allCols(fClntFMAdresse2) <> vbNullString Then
            .Range("K6").Value = allCols(fClntFMAdresse2) 'Adresse2
            .Range("K7").Value = allCols(fClntFMVille) & ", " & _
                                 allCols(fClntFMProvince) & ", " & _
                                 allCols(fClntFMCodePostal) 'Ville, Province & Code postal
        Else
            .Range("K6").Value = allCols(fClntFMVille) & ", " & _
                                 allCols(fClntFMProvince) & ", " & _
                                 allCols(fClntFMCodePostal) 'Ville, Province & Code postal
            .Range("K7").Value = vbNullString
        End If
        Application.EnableEvents = True
    End With
    
    With wshFAC_Finale
        Application.EnableEvents = False
        .Range("B23").Value = allCols(fClntFMContactFacturation)
        .Range("B24").Value = clientNamePurged
        If Trim(.Range("B23").Value) = Trim(.Range("B24").Value) Then 'Contact = Nom du client 2025-11-01 @ 05:36
            .Range("B23").Value = ""
        End If
        .Range("B25").Value = allCols(fClntFMAdresse1) 'Adresse1
        If Trim$(allCols(fClntFMAdresse2)) <> vbNullString Then
            .Range("B26").Value = allCols(fClntFMAdresse2) 'Adresse2
            .Range("B27").Value = allCols(fClntFMVille) & ", " & _
                                  allCols(fClntFMProvince) & ", " & _
                                  allCols(fClntFMCodePostal) 'Ville, Province & Code postal
        Else
            .Range("B26").Value = allCols(fClntFMVille) & ", " & _
                                  allCols(fClntFMProvince) & ", " & _
                                  allCols(fClntFMCodePostal) 'Ville, Province & Code postal
            .Range("B27").Value = vbNullString
        End If
        If Trim$(.Range("B26").Value) = ", ," Then
            .Range("B26").Value = vbNullString
        End If
        If Trim$(.Range("B27").Value) = ", ," Then
            .Range("B27").Value = vbNullString
        End If
        Application.EnableEvents = True
    End With
    
    Call EffacerTECAffiches
    
    Call ObtenirTECNonFacturablePourClient
    
    Call ObtenirTECNonFacturableDansUserForm
    
    Application.EnableEvents = True '2025-02-01 @ 06:36

Clean_Exit:

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ChangerNomDuClient" & wshFAC_Brouillon.Range("B18").Value & "'", vbNullString, startTime)
    
End Sub

Sub FACBrouillonDate_Change(d As String)

    Application.EnableEvents = False
    
    If InStr(wshFAC_Brouillon.Range("O6").Value, "-") = 0 Then
        Dim Y As String
        Y = Right$(year(d), 2)
        wshFAC_Brouillon.Range("O6").Value = Y & "-" & wshFAC_Brouillon.Range("O6").Value
        wshFAC_Finale.Range("E28").Value = wshFAC_Brouillon.Range("O6").Value
    End If
    
    'Must Get GST & PST rates and store them in wshFAC_Brouillon 'B' column at that date
    Dim DateTaxRates As Date
    DateTaxRates = CDate(d)
    wshFAC_Brouillon.Range("B29").Value = Fn_Get_Tax_Rate(DateTaxRates, "TPS")
    wshFAC_Brouillon.Range("B30").Value = Fn_Get_Tax_Rate(DateTaxRates, "TVQ")
        
    'Adjust hourly rate base on the date
    Dim lastUsedProfInSummary As Long
    lastUsedProfInSummary = wshFAC_Brouillon.Cells(wshFAC_Brouillon.Rows.count, "W").End(xlUp).Row
    
    Dim dateTauxHoraire As Date
    dateTauxHoraire = CDate(d)
    Dim i As Long
    For i = 25 To lastUsedProfInSummary
        Dim profID As Long
        profID = wshFAC_Brouillon.Range("W" & i).Value
        Dim hRate As Currency
        hRate = Fn_Get_Hourly_Rate(profID, dateTauxHoraire)
        wshFAC_Brouillon.Range("T" & i).Value = hRate
    Next i
    
    'Get all TEC for the client at a certain date
    Dim cutoffDate As Date
    cutoffDate = CDate(d)
    Call ObtenirTousLesTECPourClientAvecAF(cutoffDate, False)
    
    Dim rng As Range: Set rng = wshFAC_Brouillon.Range("L11")

    On Error Resume Next
    wshFAC_Brouillon.Range("L11").Select
    On Error GoTo 0
    
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set rng = Nothing
    
End Sub

Sub chkMontrerTECDejaFactures_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ckbMontrerTECDejaFactures_Click", vbNullString, 0)
    
    Dim cutoffDate As Date
    cutoffDate = CDate(wshFAC_Brouillon.Range("O3").Value)
    
    If wshFAC_Brouillon.Range("B16").Value = True Then
        Call ObtenirTousLesTECPourClientAvecAF(cutoffDate, True)
    Else
        Call ObtenirTousLesTECPourClientAvecAF(cutoffDate, False)
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ckbMontrerTECDejaFactures_Click", vbNullString, startTime)

End Sub

Sub MettreEnPlaceCellulesFACBrouillon()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:MettreEnPlaceCellulesFACBrouillon", vbNullString, 0)

    Application.EnableEvents = False
    
    With wshFAC_Brouillon
        .Range("B9").Value = False
        .Range("O9").Value = vbNullString 'Clear the template code
        .Range("L11:O45").ClearContents
        .Range("J47:P60").ClearContents
        
        Call AjusterLesLibellesFACBrouillon(.Range("K47"), "FAC_Label_SubTotal_1")
        Call AjusterLesLibellesFACBrouillon(.Range("K51"), "FAC_Label_SubTotal_2")
        Call AjusterLesLibellesFACBrouillon(.Range("K52"), "FAC_Label_TPS")
        Call AjusterLesLibellesFACBrouillon(.Range("K53"), "FAC_Label_TVQ")
        Call AjusterLesLibellesFACBrouillon(.Range("K55"), "FAC_Label_GrandTotal")
        Call AjusterLesLibellesFACBrouillon(.Range("K57"), "FAC_Label_Deposit")
        Call AjusterLesLibellesFACBrouillon(.Range("K59"), "FAC_Label_AmountDue")
        
        'Establish Formulas
        .Range("O47").formula = "=U35"                                   'Fees sub-total from hours summary
        .Range("O47").Font.Bold = True
        
        .Range("M48").Value = wsdADMIN.Range("FAC_Label_Frais_1").Value   'Misc. # 1 - Descr.
        .Range("O48").Value = vbNullString                                          'Misc. # 1 - Amount
        .Range("M49").Value = wsdADMIN.Range("FAC_Label_Frais_2").Value   'Misc. # 2 - Descr.
        .Range("O49").Value = vbNullString                                          'Misc. # 2 - Amount
        .Range("M50").Value = wsdADMIN.Range("FAC_Label_Frais_3").Value   'Misc. # 3 - Descr.
        .Range("O50").Value = vbNullString                                          'Misc. # 3 - Amount
        
        .Range("O51").formula = "=sum(O47:O50)"                           'Sub-total
        .Range("O51").Font.Bold = True
        
        .Range("N52").Value = wshFAC_Brouillon.Range("B29").Value         'GST Rate
        .Range("N52").NumberFormat = "0.00%"
        .Range("O52").formula = "=round(o51*n52,2)"                     'GST Amnt
        .Range("N53").Value = wshFAC_Brouillon.Range("B30").Value       'PST Rate
        .Range("N53").NumberFormat = "0.000%"
        .Range("O53").formula = "=round(o51*n53,2)"                     'PST Amnt
        .Range("O55").formula = "=sum(o51:o54)"                         'Grand Total"
        .Range("O57").Value = vbNullString
        .Range("O59").formula = "=O55-O57"                              'Deposit Amount
        
        'ON élimine les cellules qui pourraient avoir du vert pâle...
        With .Range("E3:F3,O3,O9,L11:N45,O48:O50,M48:M50").Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        'Setup the hours summary to handle different rates
        Call PreparerSommaireDesHeures
        
    End With
    
    Application.EnableEvents = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:MettreEnPlaceCellulesFACBrouillon", vbNullString, startTime)

End Sub

Sub shpOuvrirEtCopierAncienneFacture_Click()

    Call OuvrirEtCopierAncienneFacture
    
End Sub

Sub OuvrirEtCopierAncienneFacture()

    'Définir le chemin complet du répertoire des fichiers Excel par client
    Dim DossierCopieFactureExcel As String
    DossierCopieFactureExcel = wsdADMIN.Range("PATH_DATA_FILES").Value & gFACT_EXCEL_PATH & Application.PathSeparator
    If Len(Dir(DossierCopieFactureExcel, vbDirectory)) = 0 Then
        MsgBox "Le dossier est introuvable : " & DossierCopieFactureExcel, vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    ChDrive DossierCopieFactureExcel
    ChDir DossierCopieFactureExcel
    On Error GoTo 0

    'Step 1 - Open the Excel file
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx", , "Fichier Excel à ouvrir")
        
    If UCase(filePath) = "FALSE" Or UCase(filePath) = "FAUX" Then Exit Sub 'User canceled

    Dim wbSource As Workbook: Set wbSource = Workbooks.Open(filePath)
    Dim wsSource As Worksheet: Set wsSource = wbSource.Sheets(wbSource.Sheets.count) 'Position to the last worksheet

    'Step 2 - Let the user selects the cells to be copied
    MsgBox "SVP, sélectionnez les cellules à copier," & vbNewLine & vbNewLine _
         & "et par la suite, pesez sur <Enter>.", vbInformation
    On Error Resume Next
    Dim rngSource As Range
    Set rngSource = Application.InputBox("Sélectionnez les cellules à copier", Type:=8)
    On Error GoTo 0

    If rngSource Is Nothing Then
        MsgBox "Aucune cellule de sélectionnée. L'Opération est annulée.", vbExclamation
        wbSource.Close SaveChanges:=False
        GoTo Exit_Sub
    End If

    'Step 3 - Copy the selected cells
    rngSource.Copy
    If rngSource.MergeCells Then
        'Unmerged cells
        rngSource.UnMerge
    End If

    'Step 4 - Paste the copied cells at a predefined location
    Application.EnableEvents = False

    With wshFAC_Brouillon
        .Unprotect
        .Range("L11:L" & 11 + rngSource.Rows.count - 1).Value = rngSource.Value
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
'        .EnableSelection = xlUnlockedCells
    End With

    Application.EnableEvents = True
    Application.CutCopyMode = False

    'Step 5 - Close and release the Excel file
    wbSource.Close SaveChanges:=False

Exit_Sub:

    'Libérer la mémoire
    On Error Resume Next
    Set rngSource = Nothing
    Set wbSource = Nothing
    Set wsSource = Nothing
    On Error GoTo 0

End Sub

Sub AjusterLesLibellesFACBrouillon(r As Range, l As String)

    r.Value = wsdADMIN.Range(l).Value
    If wsdADMIN.Range(l & "_Bold").Value = "OUI" Then r.Font.Bold = True

End Sub

Sub shpDeplacerVersFraisDivers_Click()

    Call DeplacerVersFraisDivers

End Sub

Sub DeplacerVersFraisDivers()
    
    ActiveWindow.SmallScroll Down:=6
    wshFAC_Brouillon.Range("O47").Select 'Hours Summary
    
End Sub

Sub EffacerTECAffiches()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:EffacerTECAffiches", vbNullString, 0)
    
    Application.EnableEvents = False
    
    Dim lastRow As Long
    lastRow = wshFAC_Brouillon.Cells(wshFAC_Brouillon.Rows.count, "F").End(xlUp).Row 'First line of data is at row 7
    If lastRow > 6 Then
        'Verrouiller les cellules des descriptions des TEC - 2025-03-02 @ 21:53
        wshFAC_Brouillon.Unprotect
        wshFAC_Brouillon.Range("F7:F" & lastRow).Locked = True
        wshFAC_Brouillon.Protect UserInterfaceOnly:=True
        wshFAC_Brouillon.Range("D7:I" & lastRow + 2).ClearContents
        Call EffacerCasesACocherFACBrouillon(lastRow - 2)
    End If
    
    Application.EnableEvents = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:EffacerTECAffiches", vbNullString, startTime)

End Sub

Sub ObtenirTousLesTECPourClientAvecAF(d As Date, includeBilledTEC As Boolean)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ObtenirTousLesTECPourClientAvecAF", vbNullString, 0)
    
    'Set all criteria before calling ObtenirTECduClientAvecAF
    Dim c1 As String
    Dim c2 As Date
    Dim c3 As String, c4 As String, c5 As String
    c1 = wshFAC_Brouillon.Range("B18").Value
    Dim filterDate As Date
    filterDate = dateValue(d)
    c2 = filterDate
    c3 = Fn_Convert_Value_Boolean_To_Text(True)
    If includeBilledTEC Then
        c4 = Fn_Convert_Value_Boolean_To_Text(True)
    Else
        c4 = Fn_Convert_Value_Boolean_To_Text(False)
    End If
    c5 = Fn_Convert_Value_Boolean_To_Text(False)

    Call ObtenirTECduClientAvecAF(c1, c2, c3, c4, c5)
    
    Dim cutOffDateProjet As Date
    cutOffDateProjet = wshFAC_Brouillon.Range("B53").Value
    
    Call CopierTECFiltresVersFACBrouillon(cutOffDateProjet)
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ObtenirTousLesTECPourClientAvecAF", vbNullString, startTime)

End Sub

Sub ObtenirTECNonFacturablePourClient()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ObtenirTECNonFacturablePourClient", vbNullString, 0)
    
    'Mettre en place les critères pour aller chercher le temps NON-FACTURABLE pour le client avec AF#
    Dim c1 As String, c3 As String, c4 As String, c5 As String
    Dim c2 As Date
    c1 = wshFAC_Brouillon.Range("B18").Value
    c2 = #12/31/2099#
    c3 = Fn_Convert_Value_Boolean_To_Text(False)
    c4 = Fn_Convert_Value_Boolean_To_Text(False)
    c5 = Fn_Convert_Value_Boolean_To_Text(False)

    Call ObtenirTECduClientAvecAF(c1, c2, c3, c4, c5)
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ObtenirTECNonFacturablePourClient", vbNullString, startTime)

End Sub

Sub ObtenirTECNonFacturableDansUserForm()

    'Les charges NON FACTURABLES pour ce client sont dans TEC_Local, AF# 2
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ObtenirTECNonFacturableDansUserForm", vbNullString, 0)

    Dim lastUsedRow As Long
    lastUsedRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "AQ").End(xlUp).Row
    If lastUsedRow < 3 Then Exit Sub 'No rows
    
    Application.ScreenUpdating = False
    
    Dim totalHres As Double
    Dim ufNonBillableTime As Object
    Dim fraisDiversMsg As String
    Dim arr() As Variant
    ReDim arr(1 To (lastUsedRow - 2), 1 To 5) As Variant
    
    With wsdTEC_Local
        Dim i As Long
        For i = 3 To lastUsedRow
            arr(i - 2, 1) = .Range("AQ" & i).Value      'TECID
            arr(i - 2, 2) = .Range("AT" & i).Value      'Date
            arr(i - 2, 3) = .Range("AS" & i).Value      'Prof
            arr(i - 2, 4) = .Range("AW" & i).Value      'Description
            If Len(arr(i - 2, 4)) > 105 Then
                arr(i - 2, 4) = Left$(arr(i - 2, 4), 102) & "..."
            End If
            arr(i - 2, 5) = Format$(.Range("AX" & i).Value, "##0.00")     'Heures
            arr(i - 2, 5) = Space(7 - Len(arr(i - 2, 5))) & arr(i - 2, 5)
            totalHres = totalHres + .Range("AX" & i).Value
        Next i
    End With
    
    Set ufNonBillableTime = UserForms.Add("ufNonBillableTime")
    ufNonBillableTime.lstNonBillable.Clear
    
    With ufNonBillableTime.lstNonBillable
        .ColumnCount = UBound(arr, 2)
        .ColumnHeads = False
        .ColumnWidths = "35;60;30;463;50"
        .MultiSelect = fmMultiSelectMulti
        .List = arr
    End With
    
    ufNonBillableTime.shpConvertir.Visible = False

    ufNonBillableTime.show vbModeless
        
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ObtenirTECNonFacturableDansUserForm", vbNullString, startTime)
    
End Sub

Sub ObtenirTECduClientAvecAF(clientID As String, _
                          cutoffDate As Date, _
                          isBillable As String, _
                          isInvoiced As String, _
                          isDeleted As String)
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ObtenirTECduClientAvecAF '", clientID & _
                    "', " & cutoffDate & ", " & isBillable & ", " & isInvoiced & ", " & isDeleted, 0)
    
    Dim ws As Worksheet: Set ws = wsdTEC_Local
    
    'wshTEC_Loal_AF#2
    
    Application.ScreenUpdating = False

    With ws
        'Y a-t-il des données à filtrer ?
        Dim lastSourceRow As Long, lastResultRow As Long
        lastSourceRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row 'Last TEC Entry row
        If lastSourceRow < 3 Then Exit Sub 'Nothing to filter
        
        'Effacer les données de la dernière utilisation
        ws.Range("AM6:AM10").ClearContents
        ws.Range("AM6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:nn:ss")
        
        'Définir le range pour la source des données en utilisant un tableau
        Dim rngData As Range
        Set rngData = ws.Range("l_tbl_TEC_Local[#All]")
        ws.Range("AM7").Value = rngData.Address
        
        'Définir le range des critères
        Dim rngCriteria As Range
        Set rngCriteria = ws.Range("AK2:AO3")
        .Range("AK3").Value = clientID
        .Range("AL3").Value = "'<=" & CLng(cutoffDate)
        .Range("AM3").Value = isBillable
        If isInvoiced = True Or isInvoiced = "VRAI" Then
'        If isInvoiced <> True And isInvoiced <> "VRAI" Then
            .Range("AN3").Value = vbNullString
        Else
            .Range("AN3").Value = "FAUX"
        End If
        .Range("AO3").Value = isDeleted
        .Range("AM8").Value = rngCriteria.Address
        
        'Définir le range des résultats et effacer avant le traitement
        Dim rngResult As Range
        Set rngResult = ws.Range("AQ1").CurrentRegion
        rngResult.offset(2, 0).Clear
        Set rngResult = ws.Range("AQ2:BF2")
        .Range("AM9").Value = rngResult.Address
        
        rngData.AdvancedFilter _
                    action:=xlFilterCopy, _
                    criteriaRange:=rngCriteria, _
                    CopyToRange:=rngResult, _
                    Unique:=True
        
        'Combien avons-nous de lignes en résultat ?
        lastResultRow = .Cells(.Rows.count, "AQ").End(xlUp).Row
        .Range("AM10").Value = lastResultRow - 2 & " lignes"

        'Est-il nécessaire de trier les résultats ?
        If lastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=wsdTEC_Local.Range("AT3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add key:=wsdTEC_Local.Range("AR3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On ProfID
            .SortFields.Add key:=wsdTEC_Local.Range("AQ3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TECID
            .SetRange wsdTEC_Local.Range("AQ3:BE" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
         
No_Sort_Required:
    End With
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:ObtenirTECduClientAvecAF", vbNullString, startTime)

End Sub

Sub CopierTECFiltresVersFACBrouillon(cutOffDateProjet As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:CopierTECFiltresVersFACBrouillon", vbNullString, 0)

    Dim lastUsedRow As Long
    lastUsedRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "AQ").End(xlUp).Row
    If lastUsedRow < 3 Then Exit Sub 'No rows
    
    Application.ScreenUpdating = False
    
    Dim totalHres As Double
    Dim collFraisDivers As Collection: Set collFraisDivers = New Collection
    Dim ufFraisDivers As Object
    Dim fraisDiversMsg As String
    Dim arr() As Variant
    ReDim arr(1 To (lastUsedRow - 2), 1 To 6) As Variant
    
    With wsdTEC_Local
        Dim i As Long
        For i = 3 To lastUsedRow
            arr(i - 2, 1) = .Range("AT" & i).Value 'Date
            arr(i - 2, 2) = .Range("AS" & i).Value 'Prof
            arr(i - 2, 3) = .Range("AW" & i).Value 'Description
            arr(i - 2, 4) = .Range("AX" & i).Value 'Heures
            totalHres = totalHres + .Range("AX" & i).Value
            arr(i - 2, 5) = .Range("BB" & i).Value 'Facturée ou pas
            arr(i - 2, 6) = .Range("AQ" & i).Value 'TECID
            'Commentaires doivent être affichés
            If Trim$(.Range("AY" & i).Value) <> vbNullString Then
                fraisDiversMsg = Trim$(.Range("AY" & i).Value)
                collFraisDivers.Add fraisDiversMsg
            End If
        Next i
        'Copy array to worksheet
        Dim rng As Range
        'Set rng = .Range("D8").Resize(UBound(arr, 1), UBound(arr, 2))
        Set rng = wshFAC_Brouillon.Range("D7").Resize(lastUsedRow - 2, UBound(arr, 2))
        rng.Value = arr
    End With
    
    'Déverrouiller les cellules des descriptions des TEC - 2025-06-30 @ 07:40
    wshFAC_Brouillon.Unprotect
    Dim lastDisplayedRow As Long
    lastDisplayedRow = wshFAC_Brouillon.Cells(wshFAC_Brouillon.Rows.count, 5).End(xlUp).Row
    If lastDisplayedRow >= 7 Then
        wshFAC_Brouillon.Range("F7:F" & lastDisplayedRow).Locked = False
    End If
    wshFAC_Brouillon.Protect UserInterfaceOnly:=True
    
    'Création du userForm s'il y a quelque chose à afficher
    If collFraisDivers.count > 0 Then
        Set ufFraisDivers = UserForms.Add("ufFraisDivers")
        'Nettoyer le userForm avant d'ajouter des éléments
        ufFraisDivers.ListBox1.Clear
        'Ajouter les éléments dans le listBox
        Dim item As Variant
        For Each item In collFraisDivers
            ufFraisDivers.ListBox1.AddItem item
        Next item
        'Afficher le userForm de façon non modale
        ufFraisDivers.show vbModeless
    End If
    
    lastUsedRow = wshFAC_Brouillon.Cells(wshFAC_Brouillon.Rows.count, "D").End(xlUp).Row
    If lastUsedRow < 7 Then Exit Sub 'No rows

    'Section des TEC pour le client à une date données
    With wshFAC_Brouillon
        .Range("D7:H" & lastUsedRow + 2).Font.Color = vbBlack
        .Range("D7:H" & lastUsedRow + 2).Font.Bold = False
        
        Application.EnableEvents = False
        .Range("G" & lastUsedRow + 2).Value = totalHres
        Application.EnableEvents = False
        .Range("G7:G" & lastUsedRow + 2).NumberFormat = "##0.00"
    End With
        
    Call AjouterCasesACocherFACBrouillon(lastUsedRow, cutOffDateProjet) 'Exclude totals row

    'Adjust the formula in the hours summary
    Call AjusterFormulesDansSommaireParIndividu(lastUsedRow)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set collFraisDivers = Nothing
    Set item = Nothing
    Set ufFraisDivers = Nothing
    Set rng = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:CopierTECFiltresVersFACBrouillon", vbNullString, startTime)
    
End Sub
 
Sub shpDeplacerVersFACFinale_Click()

    Call DeplacerVersFeuilleFACFinale
    
End Sub

Sub DeplacerVersFeuilleFACFinale()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:DeplacerVersFeuilleFACFinale", vbNullString, 0)
   
    Application.ScreenUpdating = False
    
    'Vérification des montants reçus en dépôt pour le client
    If wshFAC_Brouillon.Range("B5").Value = "VRAI" Then
        GoTo Depot_Checked
    End If
    
    'Les résultats du AvancedFilter sont dans GL_Trans - Colonnes P @ Y
    Dim lastUsedRowResult As Double
    lastUsedRowResult = wsdGL_Trans.Cells(wsdGL_Trans.Rows.count, "P").End(xlUp).Row
    Dim soldeDepotClient As Double
    Dim i As Long
    For i = 2 To lastUsedRowResult
        If InStr(wsdGL_Trans.Cells(i, 18).Value, "Client:" & wshFAC_Brouillon.Range("B18").Value) <> 0 Then
            soldeDepotClient = soldeDepotClient - wsdGL_Trans.Cells(i, "V").Value + wsdGL_Trans.Cells(i, "W").Value
        End If
    Next i
    
    If soldeDepotClient > 0 Then
        MsgBox "Il y a un dépôt de client de disponible de " & Format$(soldeDepotClient, "###,##0.00 $") & vbNewLine & vbNewLine & _
            "Le total de la facture est de " & Format$(wshFAC_Brouillon.Range("O55").Value, "###,##0.00 $"), vbInformation

        Application.EnableEvents = False
        Application.ScreenUpdating = True
        
        'Cellule en surbrillance
        With wshFAC_Brouillon.Range("O57")
            .Value = WorksheetFunction.Min(soldeDepotClient, wshFAC_Brouillon.Range("O55").Value)
            .Interior.Color = RGB(255, 255, 0)
            .Select
        End With
        
        'Initialise la valeur initiale de la cellule O57 pour la comparaison
        Dim montantInitial As Variant
        montantInitial = wshFAC_Brouillon.Range("O57").Value

        'Boucle pour demander la validation
        Dim reponse As VbMsgBoxResult
        Do
            DoEvents
            'Si le montant a changé, demande confirmation
            If wshFAC_Brouillon.Range("O57").Value <> montantInitial Then
                reponse = MsgBox("Veuillez confirmer le montant du dépôt de client à appliquer" & vbNewLine & _
                                    "sur cette facture." & vbNewLine & vbNewLine & _
                                    "Appuyez sur OK pour accepter le montant suggéré," & vbNewLine & _
                                    "ou Annuler pour modifier le montant du dépôt.", _
                                    vbOKCancel + vbInformation, "Confirmation de l'imputation du dépôt de client")
               'Si l'utilisateur sélectionne Annuler, lui permet de modifier le montant
                If reponse = vbCancel Then
                    montantInitial = wshFAC_Brouillon.Range("O57").Value
                    wshFAC_Brouillon.Range("O57").Select
                End If
            End If
        Loop Until reponse = vbOK 'Continue uniquement lorsque l'utilisateur clique sur OK
        
        wshFAC_Brouillon.Range("O57").Interior.ColorIndex = xlNone
        
        Application.EnableEvents = True

    End If
    
    'Indique que la vérification a bel et bien étét faite déjà
    wshFAC_Brouillon.Range("B5").Value = "VRAI"
    
Depot_Checked:
    
    Application.ScreenUpdating = False
    
    'Copy all services line from FAC_Brouillon to FAC_Finale
    Dim iFacFinale As Long: iFacFinale = 34
    For i = 11 To 45
        With wshFAC_Finale.Range("B" & iFacFinale)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 1
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With

        If wshFAC_Brouillon.Range("L" & i).Value <> vbNullString Or wshFAC_Brouillon.Range("L" & i + 1).Value <> vbNullString Then
            Dim tiret As String
            If InStr(wshFAC_Brouillon.Range("L" & i).Value, " - ") = 0 Then
                tiret = "' - "
            Else
                tiret = "'"
            End If
            wshFAC_Finale.Range("B" & iFacFinale).Value = tiret & wshFAC_Brouillon.Range("L" & i).Value
            If wshFAC_Finale.Range("B" & iFacFinale).Value = " - " Then
                wshFAC_Finale.Range("B" & iFacFinale).Value = "'"
            End If
            iFacFinale = iFacFinale + 1
        End If
    Next i
    
    'On ne pourra plus demander une nouvelle facture, uen fois rendu ici...
    wshFAC_Brouillon.Range("B27").Value = True
    
    Call CacherHeuresParLigne
    Call MontrerSommaireTaux
    
    'Afficher le code et le nom du client, pour faciliter la sauvegarde de la facture (format EXCEL)
    wshFAC_Finale.Range("L79").Value = wshFAC_Brouillon.Range("B18").Value
    wshFAC_Finale.Range("L81").Value = wshFAC_Brouillon.Range("E3").Value
    
    wshFAC_Finale.Visible = xlSheetVisible
    wshFAC_Finale.Activate
    wshFAC_Finale.Range("I50").Select
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:DeplacerVersFeuilleFACFinale", vbNullString, startTime)

End Sub

Sub shpRetournerAuMenu_Click()

    Call RetournerMenuFAC

End Sub

Sub RetournerMenuFAC()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:RetournerMenuFAC", vbNullString, 0)
   
    On Error Resume Next
        DoEvents
        Unload ufFraisDivers
        Unload ufNonBillableTime
    On Error GoTo 0
            
    
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Application.EnableEvents = False
    
    wshFAC_Brouillon.Range("B27").Value = False
    
    'Masquer la forme (détail TEC) si elle est présente
    On Error Resume Next
    Dim shapeTextBox As Shape
    Set shapeTextBox = wshFAC_Brouillon.Shapes("shpTECInfo")
    If Not shapeTextBox Is Nothing Then
        shapeTextBox.Visible = msoFalse
    End If
    On Error GoTo 0
    
    'Libérer la mémoire
    Set shapeTextBox = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:RetournerMenuFAC", vbNullString, startTime)
    
    Call modAppli.QuitterFeuillePourMenu(wshMenuFAC, True)

End Sub

Sub AjouterCasesACocherFACBrouillon(row As Long, dateCutOffProjet As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:AjouterCasesACocherFACBrouillon", vbNullString, 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Brouillon
    
    'Unprotect the worksheet in order to be able to Unlock the cells associated with checkboxes
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    Dim chkBoxRange As Range: Set chkBoxRange = ws.Range("C7:C" & row)
    
    Dim cell As Range
    Dim cbx As checkBox
    Dim newTECapresProjet As Boolean
    newTECapresProjet = False
    
    For Each cell In chkBoxRange
    'Check if the cell is empty and doesn't have a checkbox already
    If ActiveSheet.Cells(cell.row, 8).Value = False Then
        'Create a checkbox linked to the cell
        Set cbx = wshFAC_Brouillon.CheckBoxes.Add(cell.Left + 5, cell.Top, cell.Width, cell.Height)
        With cbx
            .Name = "chkBox - " & cell.row
            .Text = vbNullString
            If dateCutOffProjet = "00:00:00" Then
                .Value = ActiveSheet.Cells(cell.row, 4).Value < wshFAC_Brouillon.Range("O3").Value
            Else
                If ActiveSheet.Cells(cell.row, 4).Value <= dateCutOffProjet Then
                    .Value = True
                Else
                    .Value = False
                    newTECapresProjet = True
                End If
            End If
            .linkedCell = cell.Address
            .Display3DShading = True
        End With
        ws.Range("C" & cell.row).Locked = False
    End If
    Next cell
    
    With ws
        .Range("D7:D" & row).NumberFormat = "dd/mm/yyyy"
        .Range("D7:D" & row).Font.Bold = False
        
        .Range("D" & row + 2).formula = "=SUMIF(C7:C" & row + 5 & ",True,G7:G" & row + 5 & ")"
        .Range("D" & row + 2).NumberFormat = "##0.00"
        .Range("D" & row + 2).Font.Bold = True
        
        .Range("B19").formula = "=SUMIF(C7:C" & row + 5 & ",True,G7:G" & row + 5 & ")"
        
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
'        .EnableSelection = xlUnlockedCells
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    DoEvents
    
    If newTECapresProjet = True Then
        MsgBox "ATTENTION - Des charges se sont ajoutées après le projet de facture" & vbNewLine & vbNewLine & _
                "VOUS DEVEZ EN TENIR COMPTE DANS VOTRE FACTURE", vbInformation + vbExclamation, _
                "Le date limite du projet de facture < Date de la facture"
    End If

    'Libérer la mémoire
    Set cbx = Nothing
    Set cell = Nothing
    Set chkBoxRange = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:AjouterCasesACocherFACBrouillon", vbNullString, startTime)

End Sub

Sub EffacerCasesACocherFACBrouillon(row As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:EffacerCasesACocherFACBrouillon", vbNullString, 0)
    
    Application.EnableEvents = False
    
    Dim cbx As Shape
    For Each cbx In wshFAC_Brouillon.Shapes
        If InStr(cbx.Name, "chkBox - ") Then
            cbx.Delete
        End If
    Next cbx
    
    'Unprotect the worksheet AND Lock the cells associated with checkbox
    Dim ws As Worksheet: Set ws = wshFAC_Brouillon
    
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    'Lock the range
    If row >= 7 Then
        ws.Range("C7:C" & row).Locked = True
    End If
    
    'Protect the worksheet
    With ws
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    wshFAC_Brouillon.Range("C7:C" & row).Value = vbNullString  'Remove text left over
    wshFAC_Brouillon.Range("D" & row + 2).Value = vbNullString 'Remove the TEC selected total formula
    wshFAC_Brouillon.Range("G" & row + 2).Value = vbNullString 'Remove the Grand total formula
    
    Application.EnableEvents = True

    'Libérer la mémoire
    Set cbx = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Brouillon:EffacerCasesACocherFACBrouillon", vbNullString, startTime)

End Sub

Sub PreparerSommaireDesHeures()

    Dim ws As Worksheet: Set ws = wshFAC_Brouillon
    
    Application.EnableEvents = False
    ws.Range("R25:U34").ClearContents
    
    Dim r As Long
    r = 11
    With wsdADMIN
        Do While .Range("D" & r).Value <> vbNullString
            ws.Range("R" & r + 14).Value = .Range("D" & r).Value 'Initiales
            ws.Range("W" & r + 14).Value = .Range("E" & r).Value 'Taux horaire
            r = r + 1
        Loop
        ws.Range("R35").Value = "Totals"
    End With
    
    With ws
        r = 25
        Do While .Range("R" & r).Value <> vbNullString
            .Range("S" & r).formula = "=SUMIFS(G7:G999, C7:C999, TRUE, E7:E999, R" & r & ")"
            .Range("U" & r).formula = "=S" & r & "*T" & r
            r = r + 1
        Loop
        ws.Range("S" & 35).formula = "=sum(S25:S34)"
        ws.Range("U" & 35).formula = "=sum(U25:U34)"
        
    End With
    
    'Cleaning - 2024-07-04 @ 16:15
    Set ws = Nothing
    
End Sub

Sub AjusterFormulesDansSommaireParIndividu(lur As Long)

    Dim i As Long, p As Long
    Application.EnableEvents = False
    For i = 25 To 34
        If wshFAC_Brouillon.Range("R" & i).Value <> vbNullString Then
            Dim f As String
            f = wshFAC_Brouillon.Range("S" & i).formula
            If InStr(1, f, "999") Then
                f = Replace(f, "999", lur)
            Else
                f = "=SUMIFS($G$7:$G$" & lur & ", $C$7:$C$" & lur & ", " & "TRUE, $E$7:$E$" & lur & ", R" & i & ")"
            End If
            wshFAC_Brouillon.Range("S" & i).formula = f
        End If
    Next i
    Application.EnableEvents = True

    'Une fois le sommaire des TEC à facturer rempli, trier en ordre descendant de la valeur
    Dim rngTECSummary As Range
    Set rngTECSummary = wshFAC_Brouillon.Range("R25:U34")
    Call TrierSommaireDesTEC(rngTECSummary)
    
    'Libérer la mémoire
    Set rngTECSummary = Nothing
    
End Sub

Sub TrierSommaireDesTEC(r As Range)

    Dim formules As Object
    Set formules = CreateObject("Scripting.Dictionary")
    
    'Enregistrer les formules et copier leurs valeurs dans les cellules
    Dim cell As Range
    For Each cell In r
        If cell.HasFormula Then
            formules.Add cell.Address, cell.formula 'Utiliser l'adresse comme clé
            cell.Value = cell.Value 'Remplacer la formule par sa valeur temporairement
        End If
    Next cell
    
    'Tri descendant sur la 4ème colonne
    r.Sort Key1:=r.Columns(4), Order1:=xlDescending, Header:=xlNo
    
    'Parcourir chaque ligne pour vider les cellules non utilisées
    Dim i As Long
    For i = 1 To r.Rows.count
        If r.Cells(i, 2).Value = 0 Then
            'Vider toutes les cellules de la ligne si la valeur de la 2ème colonne est 0
            r.Rows(i).ClearContents
        End If
    Next i
    
    'Réinsérer les formules dans les cellules concernées uniquement si la colonne 2 n'est pas zéro
    Dim addr As Variant
    Dim ligne As Integer
    Application.EnableEvents = False
    For Each addr In formules.keys
        ligne = r.Worksheet.Range(addr).Row 'Obtenir le numéro de la ligne de l'adresse
        'Vérifier la valeur de la 2ème colonne dans la ligne correspondante
        If r.Worksheet.Cells(ligne, 19).Value <> 0 Then
            'Vérifier si l'adresse est dans la colonne 2 ou 4
            If r.Worksheet.Range(addr).Column = 19 Or r.Worksheet.Range(addr).Column = 21 Then
                r.Worksheet.Range(addr).formula = formules(addr)
            End If
        End If
        
    Next addr
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set addr = Nothing
    Set cell = Nothing
    Set formules = Nothing

End Sub

Sub ChargerGabaritDescriptionFacture(t As String)

    'Is there a template letter supplied ?
    If t = vbNullString Then
        Exit Sub
    End If
    
    'Confirm use of Template
    Dim userResponse As String
    userResponse = MsgBox("Êtes-vous CERTAIN de vouloir utiliser le gabarit '" & t & "'" & vbNewLine & "pour cette facture ?", vbYesNo + vbQuestion, "Confirmation d'utilisation de gabarit")
    'If user confirms, delete the worksheets
    If userResponse <> vbYes Then
        Exit Sub
    End If
    
    'Clear whatever was there (both Brouillon & Finale)
    wshFAC_Brouillon.Range("L11:M45").ClearContents
    wshFAC_Finale.Range("B34:E63").ClearContents
    
    Dim lastUsedRow As Long
    lastUsedRow = wsdADMIN.Cells(wsdADMIN.Rows.count, "Z").End(xlUp).Row
    
    'Get the services with the appropriate template letter
    Dim strServices As String
    Dim i As Long
    For i = 12 To lastUsedRow
        If InStr(1, wsdADMIN.Range("AA" & i), t) Then
            'Build a string with 2 digits + Service description
            strServices = strServices & Right$(wsdADMIN.Range("AA" & i).Value, 2) & wsdADMIN.Range("Z" & i).Value & "|"
        End If
    Next i
    
    'Is there anything for that template ?
    If strServices = vbNullString Then
        Exit Sub
    End If
    
    'Sort the services based on the two digits in front of the service description
    Dim arr() As String
    arr = Split(strServices, "|")
    Call TrierTableauBubble(arr)

    'Go thru all the services for the template
    Dim facRow As Long
    facRow = 11
    For i = LBound(arr) + 1 To UBound(arr)
        wshFAC_Brouillon.Range("L" & facRow).Value = "'" & Mid$(arr(i), 3)
        wshFAC_Finale.Range("B" & facRow + 23).Value = "' - " & Mid$(arr(i), 3)
        facRow = facRow + 2
    Next i
        
    Application.Goto wshFAC_Brouillon.Range("L" & facRow)
    
End Sub

Public Sub test()

    Dim testPath As String
    testPath = "C:\VBA\GC_FISCALITÉ\Factures_Excel"
    
    If Len(Dir(testPath, vbDirectory)) > 0 Then
        MsgBox "Le dossier existe bien."
    Else
        MsgBox "Le dossier est introuvable."
    End If
    
    If GetAttr(testPath) And vbDirectory Then
    MsgBox "C'est bien un dossier."
    
End If

    
End Sub
