Attribute VB_Name = "modFAC_Brouillon"
Option Explicit

Dim invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Dim lastRow As Long, lastResultRow As Long, resultRow As Long

Sub FAC_Brouillon_New_Invoice() 'Clear contents
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_New_Invoice", 0)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Are we entering a NEW invoice ?
    If wshFAC_Brouillon.Range("B27").value = False Then
        With wshFAC_Brouillon
            .Range("B24").value = True
            .Range("K3:L7,O3,O5").ClearContents 'Clear cells for a new Invoice
            .Range("O6").value = Fn_Get_Next_Invoice_Number
            
            On Error Resume Next
                DoEvents
                Unload ufFraisDivers
            On Error GoTo 0
            
            Call FAC_Brouillon_Clear_All_TEC_Displayed
            
            Call FAC_Brouillon_Setup_All_Cells
            
            Application.EnableEvents = False
            .Range("B20").value = ""
            .Range("B24").value = False
            .Range("B26").value = False
            .Range("B27").value = True
            .Range("B51, B52, B53, B54").value = "" 'Requests for invoice
            .Range("R44:T48").ClearContents 'Hours & Fees summary from request for invoice
            Application.EnableEvents = True
        End With
        
        With wshFAC_Finale
            Application.EnableEvents = False
            .Range("B21,B23:C27,E28").ClearContents
            .Range("A34:F68").ClearContents
            .Range("E28").value = wshFAC_Brouillon.Range("O6").value 'Invoice #
            .Range("B69:F81").ClearContents 'NOT the formulas
            .Range("L79").value = ""
            .Range("L81").value = ""
            Application.EnableEvents = True
            
            Call FAC_Finale_Setup_All_Cells
        
        End With
        
        Application.EnableEvents = False
        wshFAC_Brouillon.Range("B16").value = False 'Does not see billed charges
        Application.EnableEvents = True
        
        Application.ScreenUpdating = True
        
        'Ensure all pending events could be processed
        DoEvents

        'Save button is disabled UNTIL the invoice is saved
        Call FAC_Finale_Disable_Save_Button
        
        flagEtapeFacture = 0
    
        'Ensure all pending events could be processed
        DoEvents

        'Introduce a small delay to ensure the worksheet is fully updated
'        Application.Wait (Now + TimeValue("0:00:01")) '2024-09-03 @ 06:45
        
        'Do we have pending requests to invoice ?
        Dim lastUsedRow As Long, liveOne As Long
        lastUsedRow = wshFAC_Projets_Ent�te.Range("A9999").End(xlUp).Row
        If lastUsedRow > 1 Then
            Dim i As Long
            For i = 2 To lastUsedRow
                If UCase(wshFAC_Projets_Ent�te.Range("Z" & i).value) = "FAUX" Or _
                    wshFAC_Projets_Ent�te.Range("Z" & i).value = 0 Then
                        liveOne = liveOne + 1
                End If
            Next i
        End If
        
        'Bring the visible area to the top
        wshFAC_Brouillon.Range("E3").Select

        If liveOne Then
            ufListeProjetsFacture.show
        End If
        
        Dim projetID As Long
        If wshFAC_Brouillon.Range("B51").value <> "" Then
            Application.EnableEvents = False
            projetID = CLng(wshFAC_Brouillon.Range("B52").value)
            'Get the Ent�te for this projetID
            lastUsedRow = wshFAC_Projets_Ent�te.Range("A9999").End(xlUp).Row
            Dim rngToSearch As Range: Set rngToSearch = wshFAC_Projets_Ent�te.Range("A1:A" & lastUsedRow)
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
                    arr(ii, 1) = wshFAC_Projets_Ent�te.Cells(matchedRow, (ii - 1) * 4 + 6).value
                    arr(ii, 2) = wshFAC_Projets_Ent�te.Cells(matchedRow, (ii - 1) * 4 + 7).value
                    arr(ii, 3) = wshFAC_Projets_Ent�te.Cells(matchedRow, (ii - 1) * 4 + 8).value
                Next ii
                'Update the summary for billing
                'Transfer data to the worksheet
                Application.EnableEvents = False
                Dim r As Long: r = 44
                For ii = 44 To 48
                    If arr(ii - 43, 1) <> "" And arr(ii - 43, 2) <> 0 Then
                        wshFAC_Brouillon.Range("R" & r).value = arr(ii - 43, 1)
                        wshFAC_Brouillon.Range("S" & r).value = arr(ii - 43, 2)
                        wshFAC_Brouillon.Range("S" & r).NumberFormat = "#,##0.00"
                        wshFAC_Brouillon.Range("T" & r).value = arr(ii - 43, 3)
                        wshFAC_Brouillon.Range("T" & r).NumberFormat = "#,##0.00 $"
                        wshFAC_Brouillon.Range("U" & r).formula = "=S" & r & "*T" & r & ""
                        r = r + 1
                   End If
                Next ii
            End If
            
            'Calcul du total des heures & des honoraires
            wshFAC_Brouillon.Range("S49").formula = "=sum(S44:S48)"
            wshFAC_Brouillon.Range("U49").formula = "=sum(U44:U48)"
            
            'The total fees amount id determined by the fees summary
            wshFAC_Brouillon.Range("O47").value = wshFAC_Brouillon.Range("U49").value
            
            wshFAC_Brouillon.Range("E3").value = wshFAC_Brouillon.Range("B51").value
            Call FAC_Brouillon_Client_Change(wshFAC_Brouillon.Range("B51").value)
            
            Application.EnableEvents = False
            
            'Utilisation de la date du projet de facture
'            Debug.Print "FAC_Brouillon_New_Invoice_140   wshFAC_Brouillon.Range(""B53"").value = "; wshFAC_Brouillon.Range("B53").value; "   "; TypeName(wshFAC_Brouillon.Range("B53").value)
'            wshFAC_Brouillon.Range("O3").value = wshFAC_Brouillon.Range("B53").value
            wshFAC_Brouillon.Range("O3").value = Now()
'            Debug.Print "FAC_Brouillon_New_Invoice_142   wshFAC_Brouillon.Range(""O3"").value = "; wshFAC_Brouillon.Range("O3").value; "   "; TypeName(wshFAC_Brouillon.Range("O3").value)
            Call FAC_Brouillon_Date_Change(wshFAC_Brouillon.Range("O3").value)
            
            wshFAC_Brouillon.Range("O9").Select
            
            Application.EnableEvents = True
        Else
            Application.EnableEvents = True
            wshFAC_Brouillon.Select
            wshFAC_Brouillon.Range("E3").Select 'Start inputing values for a NEW invoice
        End If
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Clean up
    Set rngToSearch = Nothing
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_New_Invoice", startTime)

End Sub

Sub FAC_Brouillon_Client_Change(clientName As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Client_Change(" & clientName & ")", 0)
    
    Dim myInfo() As Variant
    Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
    
    myInfo = Fn_Find_Data_In_A_Range(rng, 1, clientName, 3)
    
    If myInfo(1) = "" Then
        MsgBox "M: 101 - Je ne peux retrouver ce client dans ma liste", vbCritical
        GoTo Clean_Exit
    End If
    
    Dim clientNamePurged As String
    clientNamePurged = clientName
    Do While InStr(clientNamePurged, "[") > 0 And InStr(clientNamePurged, "]") > 0
        clientNamePurged = Fn_Strip_Contact_From_Client_Name(clientNamePurged)
    Loop
        
    Application.EnableEvents = False
    wshFAC_Brouillon.Range("B18").value = wshBD_Clients.Cells(myInfo(2), 2)
    Application.EnableEvents = True
    
    With wshFAC_Brouillon
        Application.EnableEvents = False
        .Range("K3").value = wshBD_Clients.Cells(myInfo(2), 3)
        .Range("K4").value = clientNamePurged
        .Range("K5").value = wshBD_Clients.Cells(myInfo(2), 6) 'Adresse1
        If wshBD_Clients.Cells(myInfo(2), 7) <> "" Then
            .Range("K6").value = wshBD_Clients.Cells(myInfo(2), 7) 'Adresse2
            .Range("K7").value = wshBD_Clients.Cells(myInfo(2), 8) & ", " & _
                                 wshBD_Clients.Cells(myInfo(2), 9) & ", " & _
                                 wshBD_Clients.Cells(myInfo(2), 10) 'Ville, Province & Code postal
        Else
            .Range("K6").value = wshBD_Clients.Cells(myInfo(2), 8) & ", " & _
                                 wshBD_Clients.Cells(myInfo(2), 9) & ", " & _
                                 wshBD_Clients.Cells(myInfo(2), 10) 'Ville, Province & Code postal
            .Range("K7").value = ""
        End If
        Application.EnableEvents = True
    End With
    
    With wshFAC_Finale
        Application.EnableEvents = False
        .Range("B23").value = wshBD_Clients.Cells(myInfo(2), 3)
        .Range("B24").value = clientNamePurged
        .Range("B25").value = wshBD_Clients.Cells(myInfo(2), 6) 'Adresse1
        If wshBD_Clients.Cells(myInfo(2), 7) <> "" Then
            .Range("B26").value = wshBD_Clients.Cells(myInfo(2), 7) 'Adresse2
            .Range("B27").value = wshBD_Clients.Cells(myInfo(2), 8) & ", " & _
                                wshBD_Clients.Cells(myInfo(2), 9) & ", " & _
                                wshBD_Clients.Cells(myInfo(2), 10) 'Ville, Province & Code postal
        Else
            .Range("B26").value = wshBD_Clients.Cells(myInfo(2), 8) & ", " & _
                                wshBD_Clients.Cells(myInfo(2), 9) & ", " & _
                                wshBD_Clients.Cells(myInfo(2), 10) 'Ville, Province & Code postal
            .Range("B27").value = ""
        End If
        If Trim(.Range("B26").value) = ", ," Then
            .Range("B26").value = ""
        End If
        If Trim(.Range("B27").value) = ", ," Then
            .Range("B27").value = ""
        End If
        Application.EnableEvents = True
    End With
    
    Call FAC_Brouillon_Clear_All_TEC_Displayed
    
'    wshFAC_Brouillon.Range("O3").Select 'Move on to Invoice Date

Clean_Exit:

    'Cleaning memory - 2024-07-01 @ 09:34
    Set rng = Nothing
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Client_Change - clientCode = '" & wshFAC_Brouillon.Range("B18").value & "'", startTime)
    
End Sub

Sub FAC_Brouillon_Date_Change(D As String)

    Application.EnableEvents = False
    
    If InStr(wshFAC_Brouillon.Range("O6").value, "-") = 0 Then
        Dim Y As String
        Y = Right(year(D), 2)
        wshFAC_Brouillon.Range("O6").value = Y & "-" & wshFAC_Brouillon.Range("O6").value
        wshFAC_Finale.Range("E28").value = wshFAC_Brouillon.Range("O6").value
    End If
    
'    wshFAC_Finale.Range("B21").value = "Le " & Format$(d, "d MMMM yyyy")
    
    'Must Get GST & PST rates and store them in wshFAC_Brouillon 'B' column at that date
    Dim DateTaxRates As Date
    DateTaxRates = D
    wshFAC_Brouillon.Range("B29").value = Fn_Get_Tax_Rate(DateTaxRates, "TPS")
    wshFAC_Brouillon.Range("B30").value = Fn_Get_Tax_Rate(DateTaxRates, "TVQ")
        
    'Adjust hourly rate base on the date
    Dim lastUsedProfInSummary As Long
    lastUsedProfInSummary = wshFAC_Brouillon.Range("W999").End(xlUp).Row
    
    Dim dateTauxHoraire As Date
    dateTauxHoraire = D
    Dim i As Long
    For i = 25 To lastUsedProfInSummary
        Dim profID As Long
        profID = wshFAC_Brouillon.Range("W" & i).value
        Dim hRate As Currency
        hRate = Fn_Get_Hourly_Rate(profID, dateTauxHoraire)
        
'        Dim j As Integer
'        For j = 19 To 26
'            If wshAdmin.Range("D" & j).value = wshFAC_Brouillon.Range("W" & i).value Then
'                If CDate(d) >= CDate(wshAdmin.Range("E" & j).value) Then
'                    hRate = wshAdmin.Range("F" & j).value
'                End If
'            End If
'        Next j
        wshFAC_Brouillon.Range("T" & i).value = hRate
    Next i
    
    'Get all TEC for the client at a certain date
    Dim cutoffDate As Date
    cutoffDate = D
    Call FAC_Brouillon_Get_All_TEC_By_Client(cutoffDate, False)
    
    Dim rng As Range: Set rng = wshFAC_Brouillon.Range("L11")

    On Error Resume Next
    wshFAC_Brouillon.Range("L11").Select 'Move on to Services Entry
    On Error GoTo 0
    
    Application.EnableEvents = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rng = Nothing
    
End Sub

Sub FAC_Brouillon_Inclure_TEC_Factures_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Inclure_TEC_Factures_Click", 0)
    
    Dim cutoffDate As Date
    cutoffDate = wshFAC_Brouillon.Range("O3").value
    
    If wshFAC_Brouillon.Range("B16").value = True Then
        Call FAC_Brouillon_Get_All_TEC_By_Client(cutoffDate, True)
    Else
        Call FAC_Brouillon_Get_All_TEC_By_Client(cutoffDate, False)
    End If
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Inclure_TEC_Factures_Click", startTime)

End Sub

Sub FAC_Brouillon_Setup_All_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Setup_All_Cells", 0)

    Application.EnableEvents = False
    
    With wshFAC_Brouillon
        .Range("B9").value = False
        .Range("O9").value = "" 'Clear the template code
        .Range("L11:O45").ClearContents
        .Range("J47:P60").ClearContents
        
        Call FAC_Brouillon_Set_Labels(.Range("K47"), "FAC_Label_SubTotal_1")
        Call FAC_Brouillon_Set_Labels(.Range("K51"), "FAC_Label_SubTotal_2")
        Call FAC_Brouillon_Set_Labels(.Range("K52"), "FAC_Label_TPS")
        Call FAC_Brouillon_Set_Labels(.Range("K53"), "FAC_Label_TVQ")
        Call FAC_Brouillon_Set_Labels(.Range("K55"), "FAC_Label_GrandTotal")
        Call FAC_Brouillon_Set_Labels(.Range("K57"), "FAC_Label_Deposit")
        Call FAC_Brouillon_Set_Labels(.Range("K59"), "FAC_Label_AmountDue")
        
        'Establish Formulas
'        .Range("M47").formula = "=SUM(M11:M45)"                          'Total hours entered OR TEC selected"
'        .Range("N47").formula = "=T25"                                   'Uses the first professional rate
'        .Range("N47").formula = wshAdmin.Range("TauxHoraireFacturation") 'Rate per hour
        .Range("O47").formula = "=U35"                                   'Fees sub-total from hours summary
        .Range("O47").Font.Bold = True
        
        .Range("M48").value = wshAdmin.Range("FAC_Label_Frais_1").value   'Misc. # 1 - Descr.
        .Range("O48").value = ""                                          'Misc. # 1 - Amount
        .Range("M49").value = wshAdmin.Range("FAC_Label_Frais_2").value   'Misc. # 2 - Descr.
        .Range("O49").value = ""                                          'Misc. # 2 - Amount
        .Range("M50").value = wshAdmin.Range("FAC_Label_Frais_3").value   'Misc. # 3 - Descr.
        .Range("O50").value = ""                                          'Misc. # 3 - Amount
        
        .Range("O51").formula = "=sum(O47:O50)"                           'Sub-total
        .Range("O51").Font.Bold = True
        
        .Range("N52").value = wshFAC_Brouillon.Range("B29").value         'GST Rate
        .Range("N52").NumberFormat = "0.00%"
        .Range("O52").formula = "=round(o51*n52,2)"                     'GST Amnt
        .Range("N53").value = wshFAC_Brouillon.Range("B30").value       'PST Rate
        .Range("N53").NumberFormat = "0.000%"
        .Range("O53").formula = "=round(o51*n53,2)"                     'PST Amnt
        .Range("O55").formula = "=sum(o51:o54)"                         'Grand Total"
        .Range("O57").value = ""
        .Range("O59").formula = "=O55-O57"                              'Deposit Amount
        
        'ON �limine les cellules qui pourraient avoir du vert p�le...
        With .Range("E3:F3,O3,O9,L11:N45,O48:O50,M48:M50").Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        'Setup the hours summary to handle different rates
        Call Setup_Hours_Summary
        
    End With
    
    Application.EnableEvents = True
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Setup_All_Cells", startTime)

End Sub

Sub FAC_Brouillon_Open_Copy_Paste() '2024-07-27 @ 07:46

    'Step 1 - Open the Excel file
    Dim filePath As String
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Fichier Excel � ouvrir")
    If filePath = "False" Then Exit Sub 'User canceled
    
    Dim wbSource As Workbook: Set wbSource = Workbooks.Open(filePath)
    Dim wsSource As Worksheet: Set wsSource = wbSource.Sheets(wbSource.Sheets.count) 'Position to the last worksheet
    
    'Step 2 - Let the user selects the cells to be copied
    MsgBox "SVP, s�lectionnez les cellules � copier," & vbNewLine & vbNewLine _
         & "et par la suite, pesez sur <Enter>.", vbInformation
    On Error Resume Next
    Dim rngSource As Range
    Set rngSource = Application.InputBox("S�lectionnez les cellules � copier", Type:=8)
    On Error GoTo 0
    
    If rngSource Is Nothing Then
        MsgBox "Aucune cellule de s�lectionn�e. L'Op�ration est annul�e.", vbExclamation
        wbSource.Close SaveChanges:=False
        Set wbSource = Nothing
        Exit Sub
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
        .Range("L11:N" & 11 + rngSource.rows.count - 1).value = rngSource.value
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Application.EnableEvents = True
    Application.CutCopyMode = False
    
    'Step 5 - Close and release the Excel file
    wbSource.Close SaveChanges:=False
    
    'Clean up - 2024-07-27 @ 07:39
'    Set rngDestination = Nothing
    Set rngSource = Nothing
'    Set wbDestination = Nothing
    Set wbSource = Nothing
'    Set wsDestination = Nothing
    Set wsSource = Nothing
    
End Sub

Sub FAC_Brouillon_Set_Labels(r As Range, l As String)

    r.value = wshAdmin.Range(l).value
    If wshAdmin.Range(l & "_Bold").value = "OUI" Then r.Font.Bold = True

End Sub

Sub FAC_Brouillon_Goto_Misc_Charges()
    
    ActiveWindow.SmallScroll Down:=6
    wshFAC_Brouillon.Range("O47").Select 'Hours Summary
    
End Sub

Sub FAC_Brouillon_Clear_All_TEC_Displayed()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Clear_All_TEC_Displayed", 0)
    
    Dim lastRow As Long
    lastRow = wshFAC_Brouillon.Range("D9999").End(xlUp).Row 'First line of data is at row 7
    If lastRow > 6 Then
        Application.EnableEvents = False
        wshFAC_Brouillon.Range("D7:I" & lastRow + 2).ClearContents
        Application.EnableEvents = True
        Call FAC_Brouillon_TEC_Remove_Check_Boxes(lastRow - 2)
    End If
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Clear_All_TEC_Displayed", startTime)

End Sub

Sub FAC_Brouillon_Get_All_TEC_By_Client(D As Date, includeBilledTEC As Boolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Get_All_TEC_By_Client", 0)
    
    'Set all criteria before calling FAC_Brouillon_Get_TEC_For_Client_AF
    Dim c1 As String
    Dim c2 As Date
    Dim c3 As String, c4 As String, c5 As String
    c1 = wshFAC_Brouillon.Range("B18").value
    Dim filterDate As Date
    filterDate = dateValue(D)
    c2 = filterDate
    c3 = ConvertValueBooleanToText(True)
    If includeBilledTEC Then
        c4 = ConvertValueBooleanToText(True)
    Else
        c4 = ConvertValueBooleanToText(False)
    End If
    c5 = ConvertValueBooleanToText(False)

    Call FAC_Brouillon_Clear_All_TEC_Displayed
    Call FAC_Brouillon_Filtre_Manuel_TEC(c1, c2, c3, c4, c5)
'    Call FAC_Brouillon_Get_TEC_For_Client_AF(c1, c2, c3, c4, c5)
    Dim cutOffDateProjet As Date
    cutOffDateProjet = wshFAC_Brouillon.Range("B53").value
    Call FAC_Brouillon_TEC_Filtered_Entries_Copy_To_FAC_Brouillon(cutOffDateProjet)
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Get_All_TEC_By_Client", startTime)

End Sub

Sub FAC_Brouillon_Get_TEC_For_Client_AF(clientID As String, _
        cutoffDate As Date, _
        isBillable As String, _
        isInvoiced As String, _
        isDeleted As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Get_TEC_For_Client_AF", 0)
    
    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    Application.ScreenUpdating = False

    With ws
        'Is there anything to filter ?
        Dim lastSourceRow As Long, lastResultRow As Long
        lastSourceRow = .Range("A99999").End(xlUp).Row 'Last TEC Entry row
        If lastSourceRow < 3 Then Exit Sub 'Nothing to filter
        
        'Define the source area Range
        Dim sRng As Range: Set sRng = .Range("A2:P" & lastSourceRow)
        .Range("AM10").value = sRng.Address
        
        'Define and Clear the destination area Range
        Dim dRng As Range
        lastResultRow = .Range("AQ9999").End(xlUp).Row
        If lastResultRow > 2 Then .Range("AQ3:BE" & lastResultRow).ClearContents
        Set dRng = .Range("AQ2:BE2")
        .Range("AM11").value = dRng.Address
        
        'Define the Criteria Range
        Dim cRng As Range
        If clientID <> "" Then
            .Range("AK3").value = clientID
        Else
            .Range("AK3").value = ""
        End If
        Dim filterDate As Date
        filterDate = dateValue(cutoffDate)
        .Range("AL3").value = "'<=" & Format$(cutoffDate, "dd/mm/yyyy")
'        .Range("AL3").NumberFormat = "dd/mm/yyyy"
        
        .Range("AM3").value = isBillable
        If isInvoiced <> True Then
            .Range("AN3").value = isInvoiced
        Else
            .Range("AN3").value = ""
        End If
        .Range("AO3").value = isDeleted
        Set cRng = .Range("AK2:AO3")
        .Range("AM12").value = cRng.Address
        
        'Do the Advanced Filter
        sRng.AdvancedFilter action:=xlFilterCopy, _
                            criteriaRange:=cRng, _
                            CopyToRange:=dRng, _
                            Unique:=True
        
        lastResultRow = .Range("AQ9999").End(xlUp).Row
        .Range("AM13").value = lastResultRow - 2 & " rows returned"
        .Range("AM14").value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")

        If lastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=wshTEC_Local.Range("AT3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add key:=wshTEC_Local.Range("AR3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof_ID
            .SortFields.Add key:=wshTEC_Local.Range("AQ3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshTEC_Local.Range("AQ3:BE" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
No_Sort_Required:
    End With
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set sRng = Nothing
    Set dRng = Nothing
    Set cRng = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Get_TEC_For_Client_AF", startTime)

End Sub

Sub FAC_Brouillon_Filtre_Manuel_TEC(codeClient As String, _
                                        dteCutoff As Date, _
                                        estFacturable As String, _
                                        estFacturee As String, _
                                        estDetruit As String)
    
    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    'On efface ce qui est d�j� l�...
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "AQ").End(xlUp).Row
    If lastUsedRow > 2 Then
        ws.Range("AQ3:BE" & lastUsedRow).ClearContents
    End If
    
    'D�finir la derni�re ligne contenant des donn�es
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    
    Dim rr As Long
    rr = 3
    
    Dim i As Long
    With ws
    'Boucler sur chaque ligne et masquer celles qui ne correspondent pas � tous les crit�res
        For i = 3 To lastRow ' Suppose que les donn�es commencent � la ligne 3
            If ws.Cells(i, "D").value <= dteCutoff And _
                ws.Cells(i, "E").value = codeClient And _
                ws.Cells(i, "J").value = estFacturable And _
                ws.Cells(i, "L").value = estFacturee And _
                ws.Cells(i, "N").value = estDetruit Then
                ws.Cells(rr, "AQ").value = ws.Cells(i, "A").value
                ws.Cells(rr, "AR").value = ws.Cells(i, "B").value
                ws.Cells(rr, "AS").value = ws.Cells(i, "C").value
                ws.Cells(rr, "AT").value = ws.Cells(i, "D").value
                ws.Cells(rr, "AU").value = ws.Cells(i, "E").value
                ws.Cells(rr, "AV").value = ws.Cells(i, "G").value
                ws.Cells(rr, "AW").value = ws.Cells(i, "H").value
                ws.Cells(rr, "AX").value = ws.Cells(i, "I").value
                ws.Cells(rr, "AY").value = ws.Cells(i, "J").value
                ws.Cells(rr, "AZ").value = ws.Cells(i, "K").value
                ws.Cells(rr, "BA").value = ws.Cells(i, "L").value
                ws.Cells(rr, "BB").value = ws.Cells(i, "M").value
                ws.Cells(rr, "BC").value = ws.Cells(i, "N").value
                ws.Cells(rr, "BD").value = ws.Cells(i, "O").value
                ws.Cells(rr, "BE").value = ws.Cells(i, "P").value
                rr = rr + 1
            End If
        Next i
        
        Dim lastResultRow As Long
        lastResultRow = ws.Cells(ws.rows.count, "AQ").End(xlUp).Row
        If lastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=wshTEC_Local.Range("AT3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add key:=wshTEC_Local.Range("AR3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof_ID
            .SortFields.Add key:=wshTEC_Local.Range("AQ3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshTEC_Local.Range("AQ3:BE" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
        End With
    End With
     
No_Sort_Required:
    
    'Clean up
    Set ws = Nothing
    
End Sub

Sub FAC_Brouillon_TEC_Filtered_Entries_Copy_To_FAC_Brouillon(cutOffDateProjet As Date) '2024-03-21 @ 07:10

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_TEC_Filtered_Entries_Copy_To_FAC_Brouillon", 0)

    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Local.Range("AQ9999").End(xlUp).Row
    If lastUsedRow < 3 Then Exit Sub 'No rows
    
    Application.ScreenUpdating = False
    
    Dim totalHres As Double
    Dim collFraisDivers As Collection: Set collFraisDivers = New Collection
    Dim ufFraisDivers As Object
    Dim fraisDiversMsg As String
    Dim arr() As Variant
    ReDim arr(1 To (lastUsedRow - 2), 1 To 6) As Variant
    
    With wshTEC_Local
        Dim i As Long
        For i = 3 To lastUsedRow
            arr(i - 2, 1) = .Range("AT" & i).value 'Date
            arr(i - 2, 2) = .Range("AS" & i).value 'Prof
            arr(i - 2, 3) = .Range("AV" & i).value 'Description
            arr(i - 2, 4) = .Range("AW" & i).value 'Heures
            totalHres = totalHres + .Range("AW" & i).value
            arr(i - 2, 5) = .Range("BA" & i).value 'Factur�e ou pas
            arr(i - 2, 6) = .Range("AQ" & i).value 'TEC_ID
            'Commentaires doivent �tre affich�s
            If Trim(.Range("AX" & i).value) <> "" Then
                fraisDiversMsg = Trim(.Range("AX" & i).value)
                collFraisDivers.Add fraisDiversMsg
            End If
        Next i
        'Copy array to worksheet
        Dim rng As Range
        'Set rng = .Range("D8").Resize(UBound(arr, 1), UBound(arr, 2))
        Set rng = wshFAC_Brouillon.Range("D7").Resize(lastUsedRow - 2, UBound(arr, 2))
        rng.value = arr 'RMV
    End With
    
    'Cr�ation du userForm s'il y a quelque chose � afficher
    If collFraisDivers.count > 0 Then
        Set ufFraisDivers = UserForms.Add("ufFraisDivers")
        'Nettoyer le userForm avant d'ajouter des �l�ments
        ufFraisDivers.ListBox1.Clear
        'Ajouter les �l�ments dans le listBox
        Dim item As Variant
        For Each item In collFraisDivers
            ufFraisDivers.ListBox1.AddItem item
        Next item
        'Afficher le userForm de fa�on non modale
        ufFraisDivers.show vbModeless
    End If
    
    lastUsedRow = wshFAC_Brouillon.Range("D9999").End(xlUp).Row
    If lastUsedRow < 7 Then Exit Sub 'No rows

    With wshFAC_Brouillon
        .Range("D7:H" & lastUsedRow + 2).Font.Color = vbBlack
        .Range("D7:H" & lastUsedRow + 2).Font.Bold = False
        
        Application.EnableEvents = False
        .Range("G" & lastUsedRow + 2).value = totalHres
        Application.EnableEvents = False
        .Range("G7:G" & lastUsedRow + 2).NumberFormat = "##0.00"
    End With
        
    Call FAC_Brouillon_TEC_Add_Check_Boxes(lastUsedRow, cutOffDateProjet) 'Exclude totals row

    'Adjust the formula in the hours summary
    Call Adjust_Formulas_In_The_Summary(lastUsedRow)
    
    Application.ScreenUpdating = True
    
    'Clean up
    Set collFraisDivers = Nothing
    Set item = Nothing
    Set ufFraisDivers = Nothing
    Set rng = Nothing

    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_TEC_Filtered_Entries_Copy_To_FAC_Brouillon", startTime)
    
End Sub
 
Sub FAC_Brouillon_Goto_Onglet_FAC_Finale()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Goto_Onglet_FAC_Finale", 0)
   
    Application.ScreenUpdating = False
    
    'Copy all services line from FAC_Brouillon to FAC_Finale
    Dim i As Long
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

        If wshFAC_Brouillon.Range("L" & i).value <> "" Or wshFAC_Brouillon.Range("L" & i + 1).value <> "" Then
            Dim tiret As String
            If InStr(wshFAC_Brouillon.Range("L" & i).value, " - ") = 0 Then
                tiret = "' - "
            Else
                tiret = "'"
            End If
            wshFAC_Finale.Range("B" & iFacFinale).value = tiret & wshFAC_Brouillon.Range("L" & i).value
            If wshFAC_Finale.Range("B" & iFacFinale).value = " - " Then
                wshFAC_Finale.Range("B" & iFacFinale).value = "'"
            End If
            iFacFinale = iFacFinale + 1
        End If
    Next i
    
    Call FAC_Finale_Cacher_Heures
    Call FAC_Finale_Montrer_Sommaire_Taux
    
    'Afficher le code et le nom du client, pour faciliter la sauvegarde de la facture (format EXCEL)
    wshFAC_Finale.Range("L79").value = wshFAC_Brouillon.Range("B18").value
    wshFAC_Finale.Range("L81").value = wshFAC_Brouillon.Range("E3").value
    
    wshFAC_Finale.Visible = xlSheetVisible
    wshFAC_Finale.Activate
    wshFAC_Finale.Range("I50").Select
    
    Application.ScreenUpdating = True

    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Goto_Onglet_FAC_Finale", startTime)

End Sub

Sub FAC_Brouillon_Back_To_FAC_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Back_To_FAC_Menu", 0)
   
    DoEvents
    
    Application.Wait (Now + TimeValue("0:00:02")) '2024-09-06 @ 13:42
    
    Application.EnableEvents = False
    wshFAC_Brouillon.Range("B27").value = False
    Application.EnableEvents = True
    
    wshFAC_Brouillon.Visible = xlSheetHidden
    
    wshMenuFAC.Activate
    
'    Call SlideIn_PrepFact
'    Call SlideIn_SuiviCC
'    Call SlideIn_Encaissement
'    Call SlideIn_FAC_Historique
    
    wshMenuFAC.Range("A1").Select
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_Back_To_FAC_Menu", startTime)

End Sub

Sub FAC_Brouillon_TEC_Add_Check_Boxes(Row As Long, dateCutOffProjet As Date)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_TEC_Add_Check_Boxes", 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Brouillon
    
    'Unprotect the worksheet in order to be able to Unlock the cells associated with checkboxes
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    Dim chkBoxRange As Range: Set chkBoxRange = ws.Range("C7:C" & Row)
    
    Dim cell As Range
    Dim cbx As checkBox
    Dim newTECapresProjet As Boolean
    newTECapresProjet = False
    
    For Each cell In chkBoxRange
    'Check if the cell is empty and doesn't have a checkbox already
    If Cells(cell.Row, 8).value = False Then
        'Create a checkbox linked to the cell
        Set cbx = wshFAC_Brouillon.CheckBoxes.Add(cell.Left + 5, cell.Top, cell.width, cell.Height)
        With cbx
            .name = "chkBox - " & cell.Row
            .Text = ""
            If Cells(cell.Row, 4).value <= dateCutOffProjet Then
                .value = True
            Else
                .value = False
                newTECapresProjet = True
            End If
            .linkedCell = cell.Address
            .Display3DShading = True
        End With
        ws.Range("C" & cell.Row).Locked = False
    End If
    Next cell
    
    'Unlock the checkbox to view Billed charges
    Call UnprotectCells(ws.Range("B16"))
'    ws.Range("B16").Locked = False
'
'    'Protect the worksheet
'    ws.Protect UserInterfaceOnly:=True
     
    With ws
        .Range("D7:D" & Row).NumberFormat = "dd/mm/yyyy"
        .Range("D7:D" & Row).Font.Bold = False
        
        .Range("D" & Row + 2).formula = "=SUMIF(C7:C" & Row + 5 & ",True,G7:G" & Row + 5 & ")"
        .Range("D" & Row + 2).NumberFormat = "##0.00"
        .Range("D" & Row + 2).Font.Bold = True
        
        .Range("B19").formula = "=SUMIF(C7:C" & Row + 5 & ",True,G7:G" & Row + 5 & ")"
        
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    DoEvents
    
    If newTECapresProjet = True Then
        MsgBox "ATTENTION - Des charges se sont ajout�es apr�s le projet de facture" & vbNewLine & vbNewLine & _
                "VOUS DEVEZ EN TENIR COMPTE DANS VOTRE FACTURE", vbInformation + vbExclamation, _
                "Le date limite du projet de facture < Date de la facture"
    End If

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cbx = Nothing
    Set cell = Nothing
    Set chkBoxRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_TEC_Add_Check_Boxes", startTime)

End Sub

Sub FAC_Brouillon_TEC_Remove_Check_Boxes(Row As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Brouillon:FAC_Brouillon_TEC_Remove_Check_Boxes", 0)
    
    Application.EnableEvents = False
    
    Dim cbx As Shape
    For Each cbx In wshFAC_Brouillon.Shapes
        If InStr(cbx.name, "chkBox - ") Then
            cbx.Delete
        End If
    Next cbx
    
    'Unprotect the worksheet AND Lock the cells associated with checkbox
    Dim ws As Worksheet: Set ws = wshFAC_Brouillon
    
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    'Lock the range
    ws.Range("C7:C" & Row).Locked = True
    
    'Protect the worksheet
    With ws
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    wshFAC_Brouillon.Range("C7:C" & Row).value = ""  'Remove text left over
    wshFAC_Brouillon.Range("D" & Row + 2).value = "" 'Remove the TEC selected total formula
    wshFAC_Brouillon.Range("G" & Row + 2).value = "" 'Remove the Grand total formula
    
    'Unprotect the worksheet to LOCK the cells that were associated with checkbox

    Application.EnableEvents = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cbx = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Brouillon:FAC_Brouillon_TEC_Remove_Check_Boxes", startTime)

End Sub

Sub Setup_Hours_Summary()

    Dim ws As Worksheet: Set ws = wshFAC_Brouillon
    
    Application.EnableEvents = False
    ws.Range("R25:U34").ClearContents
    Application.EnableEvents = False
    
    Dim r As Long
    r = 11
    With wshAdmin
        Do While .Range("D" & r).value <> ""
            ws.Range("R" & r + 14).value = .Range("D" & r).value
            ws.Range("W" & r + 14).value = .Range("E" & r).value
            r = r + 1
        Loop
        ws.Range("R35").value = "Totals"
    End With
    
    With ws
        r = 25
        Do While .Range("R" & r).value <> ""
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

Sub Adjust_Formulas_In_The_Summary(lur As Long)

    Dim i As Long, p As Long
    Application.EnableEvents = False
    For i = 25 To 34
        If wshFAC_Brouillon.Range("R" & i).value <> "" Then
            Dim f As String
            f = wshFAC_Brouillon.Range("S" & i).formula
            If InStr(1, f, "999") Then
                f = Replace(f, "999", lur)
            Else
                f = "=SUMIFS(G7:G" & lur & ", C7:C" & lur & ", " & "TRUE, E7:E" & lur & ", R" & i & ")"
            End If
            wshFAC_Brouillon.Range("S" & i).formula = f
        End If
    Next i
    Application.EnableEvents = True

End Sub

Sub Load_Invoice_Template(t As String)

    'Is there a template letter supplied ?
    If t = "" Then
        Exit Sub
    End If
    
    'Confirm use of Template
    Dim userResponse As String
    userResponse = MsgBox("�tes-vous CERTAIN de vouloir utiliser le gabarit '" & t & "'" & vbNewLine & "pour cette facture ?", vbYesNo + vbQuestion, "Confirmation d'utilisation de gabarit")
    'If user confirms, delete the worksheets
    If userResponse <> vbYes Then
        Exit Sub
    End If
    
    'Clear whatever was there (both Brouillon & Finale)
    wshFAC_Brouillon.Range("L11:M45").ClearContents
    wshFAC_Finale.Range("B34:E63").ClearContents
    
    Dim lastUsedRow As Long
    lastUsedRow = wshAdmin.Range("Z999").End(xlUp).Row
    
    'Get the services with the appropriate template letter
    Dim strServices As String
    Dim i As Long
    For i = 12 To lastUsedRow
        If InStr(1, wshAdmin.Range("AA" & i), t) Then
            'Build a string with 2 digits + Service description
            strServices = strServices & Right(wshAdmin.Range("AA" & i).value, 2) & wshAdmin.Range("Z" & i).value & "|"
        End If
    Next i
    
    'Is there anything for that template ?
    If strServices = "" Then
        Exit Sub
    End If
    
    'Sort the services based on the two digits in front of the service description
    Dim arr() As String
    arr = Split(strServices, "|")
    Call BubbleSort(arr)

    'Go thru all the services for the template
    Dim facRow As Long
    facRow = 11
    For i = LBound(arr) + 1 To UBound(arr)
        wshFAC_Brouillon.Range("L" & facRow).value = "'" & Mid(arr(i), 3)
        wshFAC_Finale.Range("B" & facRow + 23).value = "' - " & Mid(arr(i), 3)
        facRow = facRow + 2
    Next i
        
    Application.Goto wshFAC_Brouillon.Range("L" & facRow)
    
End Sub

Sub test_fn_get_hourly_rate()

    Dim hr As Currency
    hr = Fn_Get_Hourly_Rate(2, "2024-07-21")
    Debug.Print "test_fn_get_hourly_rate() = " & hr

End Sub


