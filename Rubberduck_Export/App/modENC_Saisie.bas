Attribute VB_Name = "modENC_Saisie"
Option Explicit

Public lastRow As Long
Private gNumeroEcritureARenverser As Long

Sub ObtenirFacturesEnSuspens(cc As String) '2024-08-21 @ 15:18
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:ObtenirFacturesEnSuspens", vbNullString, 0)
    
    Dim ws As Worksheet: Set ws = wshENC_Saisie
    
    Application.EnableEvents = False
    ws.Range("E12:K36").ClearContents 'Clear the invoices area before loading it
    Application.EnableEvents = True
    
    Call ObtenirFacturesEnSuspensAvecAF(cc)
    
    'Bring the Result from AF into our List of Oustanding Invoices
    Dim lastResultRow As Long
    lastResultRow = wsdFAC_Comptes_Clients.Cells(ws.Rows.count, "R").End(xlUp).Row
    
    Dim i As Integer
    'Unlock the required area
    With ws '2024-08-21 @ 16:06
        If lastResultRow >= 3 Then
            .Unprotect
            .Range("B12:B" & 11 + lastResultRow - 2).Locked = False
            .Range("E12:E" & 11 + lastResultRow - 2).Locked = False
            .Protect UserInterfaceOnly:=True
            .EnableSelection = xlUnlockedCells
        End If
    End With
    
    'Copy à partir du résultat de AF, dans la feuille de saisie des encaissements
    Dim rr As Integer: rr = 12
    With wsdFAC_Comptes_Clients
        For i = 3 To lastResultRow
'        For i = 3 To WorksheetFunction.Min(27, lastResultRow) 'No space for more O/S invoices
            If .Range("X" & i).Value <> 0 And _
                            Fn_FactureConfirmee(.Range("S" & i).Value) = True Then
                Application.EnableEvents = False
                wshENC_Saisie.Range("F" & rr).Value = .Range("S" & i).Value
                wshENC_Saisie.Range("G" & rr).Value = Format$(.Range("T" & i).Value, wsdADMIN.Range("B1").Value)
                wshENC_Saisie.Range("H" & rr).Value = .Range("U" & i).Value
                wshENC_Saisie.Range("I" & rr).Value = .Range("V" & i).Value + .Range("W" & i).Value
                wshENC_Saisie.Range("J" & rr).Value = .Range("X" & i).Value
                Application.EnableEvents = True
                rr = rr + 1
            End If
        Next i
    End With
    
    Call AjouterCheckBoxesEncaissement(lastResultRow - 2)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:ObtenirFacturesEnSuspens", vbNullString, startTime)

End Sub

Sub ObtenirFacturesEnSuspensAvecAF(cc As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:ObtenirFacturesEnSuspensAvecAF", vbNullString, 0)
    
    Dim ws As Worksheet: Set ws = wsdFAC_Comptes_Clients
    
    'Effacer les données de la dernière utilisation
    ws.Range("O6:O10").ClearContents
    ws.Range("O6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")

    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_FAC_Comptes_Clients[#All]")
    ws.Range("O7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("O2:P3")
    ws.Range("O3").Value = wshENC_Saisie.clientCode
    ws.Range("O8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("R1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("R2:X2")
    ws.Range("O9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
                                        
    'Est-ce que nous avons des résultats ?
'    lastResultRow = ws.Cells(ws.Rows.count, "P").End(xlUp).Row
    Dim lastResultRow As Long
    lastResultRow = ws.Cells(ws.Rows.count, "R").End(xlUp).Row
    ws.Range("O10").Value = lastResultRow - 2 & " lignes"
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 3 Then
        With ws.Sort 'Sort - InvNo
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("S3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange ws.Range("R3:X" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If
    
    'PLUG - Recalculate Column 'W' - Balance after AdvancedFilter
    Dim r As Integer
    For r = 3 To lastResultRow
        ws.Range("X" & r).Value = ws.Range("U" & r).Value - ws.Range("V" & r).Value + ws.Range("W" & r).Value
    Next r

    'libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:ObtenirFacturesEnSuspensAvecAF", vbNullString, startTime)

End Sub

Sub shpMettreAJourEncaissement_Click()

    Call MettreAJourEncaissement

End Sub

Sub MettreAJourEncaissement() '2024-08-22 @ 09:46
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:MettreAJourEncaissement", vbNullString, 0)
    
    With wshENC_Saisie
        'Check for mandatory fields (4)
        If .Range("F5").Value = Empty Or _
           .Range("K5").Value = Empty Or _
           .Range("F7").Value = Empty Or _
           .Range("K7").Value = 0 Then
            MsgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client valide" & vbNewLine & _
                "2. Une date d'encaissement" & vbNewLine & _
                "3. Un type de paiement et" & vbNewLine & _
                "4. Des montants appliqués" & vbNewLine & vbNewLine & _
                "AVANT de sauvegarder la transaction.", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Check to make sure Payment Amount = Applied Amount
        If .Range("K9").Value <> 0 Then
            MsgBox "Assurez-vous que le montant de l'encaissement soit ÉGAL" & vbNewLine & _
                "à la somme des paiements appliqués", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Create records for ENC_Entete
        Call AjouterEncEnteteDansBDMaster
        Call AjouterEncEnteteDansBDLocale
        
        Dim lastOSRow As Integer
        lastOSRow = .Cells(.Rows.count, "F").End(xlUp).Row 'Last applied Item
        
        'Create records for ENC_Details
        If lastOSRow > 11 Then
            Call AjouterEncDetailDansBDMaster(wshENC_Saisie.pmtNo, 12, lastOSRow)
            Call AjouterEncDetailDansBDLocale(wshENC_Saisie.pmtNo, 12, lastOSRow)
        End If
        
        'Update FAC_Comptes_Clients
        If lastOSRow > 11 Then
            Call MettreAJourEncComptesClientsDansBDMaster(12, lastOSRow)
            Call MettreAJourEncComptesClientsDansBDLocale(12, lastOSRow)
        End If
                
        'Mise à jour du bordereau de dépôt
        Dim lastUsedBordereau As Long
        lastUsedBordereau = .Cells(.Rows.count, "P").End(xlUp).Row
        lastUsedBordereau = lastUsedBordereau + 1
        Application.EnableEvents = False
        .Range("O" & lastUsedBordereau & ":Q" & lastUsedBordereau + 1).Clear
        
        .Range("O" & lastUsedBordereau).Value = wshENC_Saisie.pmtNo
        .Range("O" & lastUsedBordereau).HorizontalAlignment = xlCenter
        .Range("P" & lastUsedBordereau).Value = wshENC_Saisie.Range("F5").Value
        .Range("P" & lastUsedBordereau).HorizontalAlignment = xlLeft
        .Range("Q" & lastUsedBordereau).Value = wshENC_Saisie.Range("K7").Value
        .Range("Q" & lastUsedBordereau).NumberFormat = "###,##0.00 $"
        .Range("Q" & lastUsedBordereau).HorizontalAlignment = xlRight
        .Range("Q" & lastUsedBordereau + 2).formula = "=sum(Q6:Q" & lastUsedBordereau & ")"
        .Range("Q" & lastUsedBordereau + 2).NumberFormat = "###,##0.00 $"
        .Range("Q" & lastUsedBordereau + 2).Font.Bold = True
        Application.EnableEvents = True
        
        'Prepare G/L posting
        Dim noEnc As Long
        Dim nomClient As String
        Dim typeEnc As String
        Dim descEnc As String
        Dim dateEnc As Date
        Dim montantEnc As Currency
        
        noEnc = wshENC_Saisie.pmtNo
        dateEnc = wshENC_Saisie.Range("K5").Value
        nomClient = wshENC_Saisie.Range("F5").Value
        typeEnc = wshENC_Saisie.Range("F7").Value
        montantEnc = wshENC_Saisie.Range("K7").Value
        descEnc = wshENC_Saisie.Range("F9").Value
        
        Call ComptabiliserEncaissement(noEnc, dateEnc, nomClient, typeEnc, montantEnc, descEnc) '2025-07-24 @ 12:22

        MsgBox "L'encaissement '" & wshENC_Saisie.pmtNo & "' a été enregistré avec succès", vbOKOnly + vbInformation
        
        Call CreerNouvelEncaissement 'Reset the form
        
        .Range("F5").Select
    End With
    
Clean_Exit:

    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:MettreAJourEncaissement", vbNullString, startTime)

End Sub

Sub shpAnnulerEncaissement_Click()

    Call CreerNouvelEncaissement

End Sub

Sub CreerNouvelEncaissement() '2024-08-21 @ 14:58

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:CreerNouvelEncaissement", vbNullString, 0)

    Call EffacerFeuilleEncaissement
    
    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:CreerNouvelEncaissement", vbNullString, startTime)
    
End Sub

Sub AjouterEncEnteteDansBDMaster() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterEncEnteteDansBDMaster", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "ENC_Entete$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxPmtNo As Long
    strSQL = "SELECT MAX(PayID) AS MaxPmtNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxPmtNo
    recSet.Open strSQL, conn
    
    'Get the last used row
    Dim lr As Long
    If IsNull(recSet.Fields("MaxPmtNo").Value) Then
        'Handle empty table (assign a default value, e.g., 1)
        lr = 0
    Else
        lr = recSet.Fields("MaxPmtNo").Value
    End If
    
    'Calculate the new PmtNo
    wshENC_Saisie.pmtNo = lr + 1

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    recSet.Close
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    recSet.AddNew
        recSet.Fields(fEncEPayID - 1).Value = wshENC_Saisie.pmtNo
        recSet.Fields(fEncEPayDate - 1).Value = wshENC_Saisie.Range("K5").Value
        recSet.Fields(fEncECustomer - 1).Value = wshENC_Saisie.Range("F5").Value
        recSet.Fields(fEncECodeClient - 1).Value = wshENC_Saisie.clientCode
        recSet.Fields(fEncEPayType - 1).Value = wshENC_Saisie.Range("F7").Value
        recSet.Fields(fEncEAmount - 1).Value = CDbl(Format$(wshENC_Saisie.Range("K7").Value, "#,##0.00 $"))
        recSet.Fields(fEncENotes - 1).Value = wshENC_Saisie.Range("F9").Value
        recSet.Fields(fEncETimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    'Update the recordset (create the record)
    recSet.Update
    
    'Close recordset and connection
    recSet.Close
    Set recSet = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterEncEnteteDansBDMaster", vbNullString, startTime)
    
End Sub

Sub AjouterEncEnteteDansBDLocale() '2024-08-22 @ 10:38
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterEncEnteteDansBDLocale", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentPmtNo As Long
    currentPmtNo = wshENC_Saisie.pmtNo
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wsdENC_Entete.Cells(wsdENC_Entete.Rows.count, "A").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    wsdENC_Entete.Cells(rowToBeUsed, fEncEPayID).Value = currentPmtNo
    wsdENC_Entete.Cells(rowToBeUsed, fEncEPayDate).Value = CDate(wshENC_Saisie.Range("K5").Value)
    wsdENC_Entete.Cells(rowToBeUsed, fEncECustomer).Value = wshENC_Saisie.Range("F5").Value
    wsdENC_Entete.Cells(rowToBeUsed, fEncECodeClient).Value = wshENC_Saisie.clientCode
    wsdENC_Entete.Cells(rowToBeUsed, fEncEPayType).Value = wshENC_Saisie.Range("F7").Value
    wsdENC_Entete.Cells(rowToBeUsed, fEncEAmount).Value = CDbl(Format$(wshENC_Saisie.Range("K7").Value, "#,##0.00"))
    wsdENC_Entete.Cells(rowToBeUsed, fEncENotes).Value = wshENC_Saisie.Range("F9").Value
    wsdENC_Entete.Cells(rowToBeUsed, fEncETimeStamp).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterEncEnteteDansBDLocale", vbNullString, startTime)

End Sub

Sub AjouterEncDetailDansBDMaster(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterEncDetailDansBDMaster", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "ENC_Details$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
        
    'Build the recordSet
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).Value = True And _
            wshENC_Saisie.Range("K" & r).Value <> 0 Then
            recSet.AddNew
                recSet.Fields(fEncDPayID - 1).Value = CLng(pmtNo)
                recSet.Fields(fEncDInvNo - 1).Value = wshENC_Saisie.Range("F" & r).Value
                recSet.Fields(fEncDCustomer - 1).Value = wshENC_Saisie.Range("F5").Value
                recSet.Fields(fEncDPayDate - 1).Value = wshENC_Saisie.Range("K5").Value
                recSet.Fields(fEncDPayAmount - 1).Value = CDbl(Format$(wshENC_Saisie.Range("K" & r).Value, "#,##0.00 $"))
                recSet.Fields(fEncDTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            'Update the recordset (create the record)
            recSet.Update
        End If
    Next r
    
    'Close recordset and connection
    recSet.Close
    Set recSet = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterEncDetailDansBDMaster", vbNullString, startTime)
    
End Sub

Sub AjouterEncDetailDansBDLocale(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterEncDetailDansBDLocale", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'What is the last used row in ENC_Details ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wsdENC_Details.Cells(wsdENC_Details.Rows.count, 1).End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).Value = True And _
            wshENC_Saisie.Range("K" & r).Value <> 0 Then
            wsdENC_Details.Range("A" & rowToBeUsed).Value = pmtNo
            wsdENC_Details.Range("B" & rowToBeUsed).Value = wshENC_Saisie.Range("F" & r).Value
            wsdENC_Details.Range("C" & rowToBeUsed).Value = wshENC_Saisie.Range("F5").Value
            wsdENC_Details.Range("D" & rowToBeUsed).Value = CDate(wshENC_Saisie.Range("K5").Value)
            wsdENC_Details.Range("E" & rowToBeUsed).Value = CDbl(Format$(wshENC_Saisie.Range("K" & r).Value, "#,##0.00"))
            wsdENC_Details.Range("F" & rowToBeUsed).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    Next r
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterEncDetailDansBDLocale", vbNullString, startTime)

End Sub

Sub MettreAJourEncComptesClientsDansBDMaster(firstRow As Integer, lastRow As Integer) 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:MettreAJourEncComptesClientsDansBDMaster", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    Dim r As Long
    For r = firstRow To lastRow
        If wshENC_Saisie.Range("B" & r).Value = True And _
            wshENC_Saisie.Range("K" & r).Value <> 0 Then
            'Open the recordset for the specified invoice
            Dim Inv_No As String
            Inv_No = CStr(Trim$(wshENC_Saisie.Range("F" & r).Value))
            
            Dim strSQL As String
            strSQL = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & Inv_No & "'"
            recSet.Open strSQL, conn, 2, 3
            If Not recSet.EOF Then
                'Mettre à jour Amount_Paid
                recSet.Fields(fFacCCTotalPaid - 1).Value = recSet.Fields(fFacCCTotalPaid - 1).Value + CDbl(wshENC_Saisie.Range("K" & r).Value)
                'Mettre à jour Status
                If recSet.Fields(fFacCCTotal - 1).Value - recSet.Fields(fFacCCTotalPaid - 1).Value = 0 Then
                    On Error Resume Next
                    recSet.Fields(fFacCCStatus - 1).Value = "Paid"
                    If Err.Number <> 0 Then
                        MsgBox "Erreur #" & Err.Number & " : " & Err.description
                    End If
                    On Error GoTo 0
                Else
                    recSet.Fields(fFacCCStatus - 1).Value = "Unpaid"
                End If
                'Mettre à jour le solde de la facture
                recSet.Fields(fFacCCBalance - 1).Value = recSet.Fields(fFacCCTotal - 1).Value - recSet.Fields(fFacCCTotalPaid - 1).Value + recSet.Fields(fFacCCTotalRegul - 1).Value
                recSet.Update
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec la facture '" & Inv_No & "' ne peut être retrouvé!", _
                    vbExclamation
                GoTo Clean_Exit
            End If
            'Update the recordset (create the record)
            recSet.Update
            recSet.Close
        End If
    Next r
    
Clean_Exit:
    
    'Close recordset and connection
    Set recSet = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:MettreAJourEncComptesClientsDansBDMaster", vbNullString, startTime)
    
End Sub

Sub MettreAJourEncComptesClientsDansBDLocale(firstRow As Integer, lastRow As Integer) '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:MettreAJourEncComptesClientsDansBDLocale", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wsdFAC_Comptes_Clients
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim r As Integer
    For r = firstRow To lastRow
        Dim Inv_No As String
        Inv_No = CStr(Trim(wshENC_Saisie.Range("F" & r).Value))
        
        Dim foundRange As Range
        Set foundRange = lookupRange.Find(What:=Inv_No, LookIn:=xlValues, LookAt:=xlWhole)
    
        Dim rowToBeUpdated As Long
        If Not foundRange Is Nothing Then
            rowToBeUpdated = foundRange.row
            ws.Cells(rowToBeUpdated, fFacCCTotalPaid).Value = ws.Cells(rowToBeUpdated, fFacCCTotalPaid).Value + wshENC_Saisie.Range("K" & r).Value
            ws.Cells(rowToBeUpdated, fFacCCBalance).Value = ws.Cells(rowToBeUpdated, fFacCCBalance).Value - wshENC_Saisie.Range("K" & r).Value
            'Est-ce que le solde de la facture est à 0,00 $ ?
            If ws.Cells(rowToBeUpdated, fFacCCBalance).Value = 0 Then
                ws.Cells(rowToBeUpdated, fFacCCStatus) = "Paid"
            Else
                ws.Cells(rowToBeUpdated, fFacCCStatus) = "Unpaid"
            End If
        Else
            MsgBox "La facture '" & Inv_No & "' n'existe pas dans FAC_Comptes_Clients.", vbCritical
        End If
    Next r
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:MettreAJourEncComptesClientsDansBDLocale", vbNullString, startTime)

End Sub

Sub ComptabiliserEncaissement(noEnc As Long, dt As Date, nom As String, _
                              typeEnc As String, montant As Currency, desc As String) '2025-07-24 @ 07:02
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:ComptabiliserEncaissement", vbNullString, 0)
    
    Dim ws As Worksheet
    Set ws = wshTEC_Evaluation
    
    Dim glEncaisse As String, descGLEncaisse As String
    Dim glComptesClients As String, descGLComptesClients As String
    Dim glProduitPercuAvance As String, descGLProduitPercuAvance As String
    
    'Comptes de GL et description du poste
    glEncaisse = Fn_NoCompteAPartirIndicateurCompte("Encaisse")
    descGLEncaisse = Fn_DescriptionAPartirNoCompte(glEncaisse)
    glComptesClients = Fn_NoCompteAPartirIndicateurCompte("Comptes Clients")
    descGLComptesClients = Fn_DescriptionAPartirNoCompte(glComptesClients)
    glProduitPercuAvance = Fn_NoCompteAPartirIndicateurCompte("Produit perçu d'avance")
    descGLProduitPercuAvance = Fn_DescriptionAPartirNoCompte(glProduitPercuAvance)
    
    'Déclaration et instanciation d'un objet GL_Entry
    Dim ecr As clsGL_Entry
    Set ecr = New clsGL_Entry

    'Remplissage des propriétés communes
    ecr.DateEcriture = dt
    ecr.description = nom
    ecr.source = "ENCAISSEMENT:" & Format$(noEnc, "00000")

    'Ajoute autant de lignes que nécessaire
    If montant <> 0 Then
        'Portion Débit
        If Not wshENC_Saisie.Range("F7").Value = "Dépôt de client" Then
            'Encaisse
            ecr.AjouterLigne glEncaisse, descGLEncaisse, montant, desc
        Else
            'Produit perçu d'avance
            ecr.description = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            ecr.source = UCase$(wshENC_Saisie.Range("F7").Value) & ":" & Format$(noEnc, "00000")
            ecr.AjouterLigne glProduitPercuAvance, descGLProduitPercuAvance, montant, desc
        End If
        'Crédit Comptes-Clients
        ecr.AjouterLigne glComptesClients, descGLComptesClients, -montant, desc
    End If
    
    'Écriture
    Call modGL_Stuff.AjouterEcritureGLADOPlusLocale(ecr, False)
    
    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:ComptabiliserEncaissement", vbNullString, startTime)

End Sub

Sub AjouterCheckBoxesEncaissement(row As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterCheckBoxesEncaissement", vbNullString, 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshENC_Saisie
    
    Dim chkBoxRange As Range: Set chkBoxRange = ws.Range("E12:E" & 12 + row)
    
    Dim cell As Range
    Dim cbx As checkBox
    For Each cell In chkBoxRange
    'Check if the cell is empty and doesn't have a checkbox already
    If cell.row <= 36 And _
        ActiveSheet.Cells(cell.row, 2).Value = vbNullString And _
        ActiveSheet.Cells(cell.row, 6).Value <> vbNullString Then 'Applied = False
            'Create a checkbox linked to the cell
            Set cbx = wshENC_Saisie.CheckBoxes.Add(cell.Left + 30, cell.Top, cell.Width, cell.Height)
            With cbx
                .Name = "chkBox - " & cell.row
                .Caption = vbNullString
                .Value = False
                .linkedCell = "B" & cell.row
                .Display3DShading = True
                .OnAction = "chkAppliquerEncaissementLigne"
                .Locked = False
            End With
    End If
    Next cell

    'Protect the worksheet
    With ws
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Application.EnableEvents = True

    'Libérer la mémoire
    Set cbx = Nothing
    Set cell = Nothing
    Set chkBoxRange = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:AjouterCheckBoxesEncaissement", vbNullString, startTime)

End Sub

Sub EffacerCasesACocherENC(row As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:EffacerCasesACocherENC", vbNullString, 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Delete all checkboxes whose name are chkBox - ...
    Dim cbx As Shape
    For Each cbx In wshENC_Saisie.Shapes
        If InStr(cbx.Name, "chkBox -") Then
            cbx.Delete
        End If
    Next cbx
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set cbx = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:EffacerCasesACocherENC", vbNullString, startTime)

End Sub

Sub EffacerFeuilleEncaissement()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:EffacerFeuilleEncaissement", vbNullString, 0)
    
    wshENC_Saisie.Unprotect
    
    With wshENC_Saisie
    
        Application.EnableEvents = False
        
        .Range("B5,F5:H5,K7,F9:I9,E12:K36").ClearContents 'Clear Fields
        .Range("B12:B36").ClearContents
        
        .Range("K5").Value = vbNullString
        .Range("F7").Value = "Banque" 'Set Default type
        .Range("F5").Activate
        
    End With
    
    'Note the lastUsedRow for checkBox deletion
    Dim lastUsedRow As Long
    lastUsedRow = wshENC_Saisie.Cells(wshENC_Saisie.Rows.count, "F").End(xlUp).Row
    If lastUsedRow > 36 Then
        lastUsedRow = 36
    End If
    If lastUsedRow > 11 Then
        Call EffacerCasesACocherENC(lastUsedRow)
    End If
        
    With wshENC_Saisie.Range("F5:H5, K5, F7, K7, F9:I9").Interior '2024-08-25 @ 09:21
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    
    wshENC_Saisie.Shapes("shpSauvegarderENC").Visible = False
    wshENC_Saisie.Shapes("shpAnnulerSaisieENC").Visible = False
    
    Application.EnableEvents = True
    
    With wshENC_Saisie
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:EffacerFeuilleEncaissement", vbNullString, startTime)

End Sub

Sub chkAppliquerEncaissementLigne()

    Dim chkBox As checkBox
    Set chkBox = ActiveSheet.CheckBoxes(Application.Caller)
    Dim linkedCell As Range
    Set linkedCell = ActiveSheet.Range(chkBox.linkedCell)
    
    If linkedCell.Value = True Then
        If wshENC_Saisie.Range("K9").Value > 0 Then
            Application.EnableEvents = False
            If wshENC_Saisie.Range("K9").Value > wshENC_Saisie.Range("J" & linkedCell.row).Value Then
                wshENC_Saisie.Range("K" & linkedCell.row).Value = wshENC_Saisie.Range("J" & linkedCell.row).Value
            Else
                wshENC_Saisie.Range("K" & linkedCell.row).Value = wshENC_Saisie.Range("K9").Value
            End If
            Application.EnableEvents = True
        End If
        wshENC_Saisie.Shapes("btnENC_Sauvegarde").Visible = True
        wshENC_Saisie.Shapes("btnENC_Annule").Visible = True
    Else
        ActiveSheet.Range("K" & linkedCell.row).Value = 0
    End If

    'Libérer la mémoire
    Set chkBox = Nothing
    Set linkedCell = Nothing
    
End Sub

Sub shpSortirEncaissement_Click()

    If ActiveSheet.Range("K7").Value <> 0 Then
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox("Voulez-vous vraiment quitter SANS enregistrer" & vbNewLine & vbNewLine & _
                "l'encaissement qui n'a pas été mis à jour ?", _
                vbExclamation + vbYesNo + vbDefaultButton2, "Confirmation avant de quitter SANS enregistrer")
        If reponse = vbYes Then
            'Exemple d'action : retour à la feuille "MenuPrincipal"
            Call RetournerMenuComptabilite
'            Application.GoTo Worksheets("MenuPrincipal").Range("A1")
        End If
    Else
        Call RetournerMenuComptabilite
    End If
    
End Sub

Sub RetournerMenuComptabilite()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:RetournerMenuComptabilite", vbNullString, 0)
   
    If wshENC_Saisie.ProtectContents Then
        wshENC_Saisie.Unprotect
    End If
    
    Application.EnableEvents = False
    
    Call EffacerFeuilleEncaissement
    
    Application.EnableEvents = True
    
    wshENC_Saisie.Visible = xlSheetVeryHidden

    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
    Call modDev_Utils.EnregistrerLogApplication("modENC_Saisie:RetournerMenuComptabilite", vbNullString, startTime)

End Sub

Sub AjusterLibelleDansEncaissement(typeTrans As String)

    Application.EnableEvents = False
    
    If Not typeTrans = "Régularisations" Then
        wshENC_Saisie.Range("J5").Value = "Date encaissement:"
        wshENC_Saisie.Range("J5").Font.Color = vbBlack
        wshENC_Saisie.Range("J7").Value = "Total encaissement:"
        wshENC_Saisie.Range("J7").Font.Color = vbBlack
    Else
        wshENC_Saisie.Range("J5").Value = "Date RÉGULARISATION:"
        wshENC_Saisie.Range("J5").Font.Color = vbRed
        wshENC_Saisie.Range("J7").Value = "Total RÉGULARISATION:"
        wshENC_Saisie.Range("J7").Font.Color = vbRed
    End If

    Application.EnableEvents = True

End Sub

Sub ValiderEtLancerufEncRegularisation()

    Dim ws As Worksheet
    Set ws = wshENC_Saisie
    
    'Vérification des champs obligatoires
    If IsEmpty(ws.Range("F5").Value) Then
        MsgBox "Le client est obligatoire. Veuillez le choisir avant de continuer.", vbExclamation
        Exit Sub
    End If

    If IsEmpty(ws.Range("K5").Value) Then
        MsgBox "La date est obligatoire. Veuillez la saisir avant de continuer.", vbExclamation
        Exit Sub
    End If
    
    If ws.Range("K7").Value = 0 Then
        If Not ws.Range("F7").Value = "Régularisations" Then
            MsgBox _
                Prompt:="Le montant de l'encaissement est obligatoire." & vbNewLine & vbNewLine & _
                        "Veuillez le fournir avant de continuer.", _
                Title:="Un montant est requis", _
                Buttons:=vbExclamation
        Else
            MsgBox _
                Prompt:="Le montant de la régularisation est obligatoire." & vbNewLine & vbNewLine & _
                        "Veuillez le fournir avant de continuer.", _
                Title:="Un montant est requis", _
                Buttons:=vbExclamation
        End If
        Exit Sub
    End If
    
    'Condition pour lancer le UserForm
    If ws.Range("F7").Value = "Régularisations" Then
        ufEncRégularisation.show
    End If
    
End Sub


