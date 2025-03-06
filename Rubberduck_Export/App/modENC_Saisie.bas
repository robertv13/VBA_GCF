Attribute VB_Name = "modENC_Saisie"
Option Explicit

'TODO Is this comment still valid? => Variables globales pour le module
Public lastRow As Long
Private payRow As Long
Private gNumeroEcritureARenverser As Long

Sub ENC_Get_OS_Invoices(cc As String) '2024-08-21 @ 15:18
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices", "", 0)
    
    Dim ws As Worksheet: Set ws = wshENC_Saisie
    
    Application.EnableEvents = False
    ws.Range("E12:K36").ClearContents 'Clear the invoices area before loading it
    Application.EnableEvents = True
    
    Call ENC_Get_OS_Invoices_With_AF(cc)
    
    'Bring the Result from AF into our List of Oustanding Invoices
    Dim lastResultRow As Long
    lastResultRow = wshFAC_Comptes_Clients.Cells(ws.Rows.count, "R").End(xlUp).row
    
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
    With wshFAC_Comptes_Clients
        For i = 3 To WorksheetFunction.Min(27, lastResultRow) 'No space for more O/S invoices
            If .Range("X" & i).value <> 0 And _
                            Fn_Invoice_Is_Confirmed(.Range("S" & i).value) = True Then
                Application.EnableEvents = False
                wshENC_Saisie.Range("F" & rr).value = .Range("S" & i).value
                wshENC_Saisie.Range("G" & rr).value = Format$(.Range("T" & i).value, wshAdmin.Range("B1").value)
                wshENC_Saisie.Range("H" & rr).value = .Range("U" & i).value
                wshENC_Saisie.Range("I" & rr).value = .Range("V" & i).value + .Range("W" & i).value
                wshENC_Saisie.Range("J" & rr).value = .Range("X" & i).value
                Application.EnableEvents = True
                rr = rr + 1
            End If
        Next i
    End With
    
    Call ENC_Add_Check_Boxes(lastResultRow - 2)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices", "", startTime)

End Sub

Sub ENC_Get_OS_Invoices_With_AF(cc As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices_With_AF", "", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    'Effacer les données de la dernière utilisation
    ws.Range("O6:O10").ClearContents
    ws.Range("O6").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")

    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("tblFAC_Comptes_Clients[#All]")
    ws.Range("O7").value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("O2:P3")
    ws.Range("O3").value = wshENC_Saisie.clientCode
    ws.Range("O8").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
'    Set rngResult = ws.Range("P1").CurrentRegion
    Set rngResult = ws.Range("R1").CurrentRegion
    rngResult.offset(2, 0).Clear
'    Set rngResult = ws.Range("P2:U2")
    Set rngResult = ws.Range("R2:X2")
    ws.Range("O9").value = rngResult.Address
    
    rngData.AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
                                        
    'Est-ce que nous avons des résultats ?
'    lastResultRow = ws.Cells(ws.Rows.count, "P").End(xlUp).row
    Dim lastResultRow As Long
    lastResultRow = ws.Cells(ws.Rows.count, "R").End(xlUp).row
    ws.Range("O10").value = lastResultRow - 2 & " lignes"
    
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
    
'    'PLUG - Recalculate Column 'U' - Balance after AdvancedFilter
    'PLUG - Recalculate Column 'W' - Balance after AdvancedFilter
    Dim r As Integer
    For r = 3 To lastResultRow
        ws.Range("X" & r).value = ws.Range("U" & r).value - ws.Range("V" & r).value + ws.Range("W" & r).value
    Next r

    'libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices_With_AF", "", startTime)

End Sub

Sub shp_ENC_Update_Click()

    Call MAJ_Encaissement

End Sub

Sub MAJ_Encaissement() '2024-08-22 @ 09:46
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:MAJ_Encaissement", "", 0)
    
    With wshENC_Saisie
        'Check for mandatory fields (4)
        If .Range("F5").value = Empty Or _
           .Range("K5").value = Empty Or _
           .Range("F7").value = Empty Or _
           .Range("K7").value = 0 Then
            msgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client valide" & vbNewLine & _
                "2. Une date d'encaissement" & vbNewLine & _
                "3. Un type de paiement et" & vbNewLine & _
                "4. Des montants appliqués" & vbNewLine & vbNewLine & _
                "AVANT de sauvegarder la transaction.", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Check to make sure Payment Amount = Applied Amount
        If .Range("K9").value <> 0 Then
            msgBox "Assurez-vous que le montant de l'encaissement soit ÉGAL" & vbNewLine & _
                "à la somme des paiements appliqués", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Create records for ENC_Entête
        Call ENC_Add_DB_Entete
        Call ENC_Add_Locally_Entete
        
        Dim lastOSRow As Integer
        lastOSRow = .Cells(.Rows.count, "F").End(xlUp).row 'Last applied Item
        
        'Create records for ENC_Détails
        If lastOSRow > 11 Then
            Call ENC_Add_DB_Details(wshENC_Saisie.pmtNo, 12, lastOSRow)
            Call ENC_Add_Locally_Details(wshENC_Saisie.pmtNo, 12, lastOSRow)
        End If
        
        'Update FAC_Comptes_Clients
        If lastOSRow > 11 Then
            Call ENC_Update_DB_Comptes_Clients(12, lastOSRow)
            Call ENC_Update_Locally_Comptes_Clients(12, lastOSRow)
        End If
                
        'Mise à jour du bordereau de dépôt
        Dim lastUsedBordereau As Long
        lastUsedBordereau = .Cells(.Rows.count, "P").End(xlUp).row
        lastUsedBordereau = lastUsedBordereau + 1
        Application.EnableEvents = False
        .Range("O" & lastUsedBordereau & ":Q" & lastUsedBordereau + 1).Clear
        
        .Range("O" & lastUsedBordereau).value = wshENC_Saisie.pmtNo
        .Range("O" & lastUsedBordereau).HorizontalAlignment = xlCenter
        .Range("P" & lastUsedBordereau).value = wshENC_Saisie.Range("F5").value
        .Range("P" & lastUsedBordereau).HorizontalAlignment = xlLeft
        .Range("Q" & lastUsedBordereau).value = wshENC_Saisie.Range("K7").value
        .Range("Q" & lastUsedBordereau).NumberFormat = "###,##0.00 $"
        .Range("Q" & lastUsedBordereau).HorizontalAlignment = xlRight
        .Range("Q" & lastUsedBordereau + 2).formula = "=sum(Q6:Q" & lastUsedBordereau & ")"
        .Range("Q" & lastUsedBordereau + 2).NumberFormat = "###,##0.00 $"
        .Range("Q" & lastUsedBordereau + 2).Font.Bold = True
        Application.EnableEvents = True
        
        'Prepare G/L posting
        Dim noEnc As String, nomClient As String, typeEnc As String, descEnc As String
        Dim dateEnc As Date
        Dim montantEnc As Currency
        noEnc = wshENC_Saisie.pmtNo
        dateEnc = wshENC_Saisie.Range("K5").value
        nomClient = wshENC_Saisie.Range("F5").value
        typeEnc = wshENC_Saisie.Range("F7").value
        montantEnc = wshENC_Saisie.Range("K7").value
        descEnc = wshENC_Saisie.Range("F9").value

        Call ENC_GL_Posting_DB(noEnc, dateEnc, nomClient, typeEnc, montantEnc, descEnc)  '2024-08-22 @ 16:08
        Call ENC_GL_Posting_Locally(noEnc, dateEnc, nomClient, typeEnc, montantEnc, descEnc)  '2024-08-22 @ 16:08
        
        msgBox "L'encaissement '" & wshENC_Saisie.pmtNo & "' a été enregistré avec succès", vbOKOnly + vbInformation
        
        Call Encaissement_Add_New 'Reset the form
        
        .Range("F5").Select
    End With
    
Clean_Exit:

    Call Log_Record("modENC_Saisie:MAJ_Encaissement", "", startTime)

End Sub

Sub Encaissement_Add_New() '2024-08-21 @ 14:58

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:Encaissement_Add_New", "", 0)

    Call ENC_Clear_Cells
    
    Call Log_Record("modENC_Saisie:Encaissement_Add_New", "", startTime)
    
End Sub

Sub ENC_Add_DB_Entete() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Entete", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "ENC_Entête$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxPmtNo As Long
    strSQL = "SELECT MAX(PayID) AS MaxPmtNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxPmtNo
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lr As Long
    If IsNull(rs.Fields("MaxPmtNo").value) Then
        'Handle empty table (assign a default value, e.g., 1)
        lr = 0
    Else
        lr = rs.Fields("MaxPmtNo").value
    End If
    
    'Calculate the new PmtNo
    wshENC_Saisie.pmtNo = lr + 1

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
        rs.Fields(fEncEPayID - 1).value = wshENC_Saisie.pmtNo
        rs.Fields(fEncEPayDate - 1).value = wshENC_Saisie.Range("K5").value
        rs.Fields(fEncECustomer - 1).value = wshENC_Saisie.Range("F5").value
        rs.Fields(fEncECodeClient - 1).value = wshENC_Saisie.clientCode
        rs.Fields(fEncEPayType - 1).value = wshENC_Saisie.Range("F7").value
        rs.Fields(fEncEAmount - 1).value = CDbl(Format$(wshENC_Saisie.Range("K7").value, "#,##0.00 $"))
        rs.Fields(fEncENotes - 1).value = wshENC_Saisie.Range("F9").value
        rs.Fields(fEncETimeStamp - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_DB_Entete", "", startTime)
    
End Sub

Sub ENC_Add_Locally_Entete() '2024-08-22 @ 10:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Locally_Entete", "", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentPmtNo As Long
    currentPmtNo = wshENC_Saisie.pmtNo
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_Entête.Cells(wshENC_Entête.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    wshENC_Entête.Cells(rowToBeUsed, fEncEPayID).value = currentPmtNo
    wshENC_Entête.Cells(rowToBeUsed, fEncEPayDate).value = wshENC_Saisie.Range("K5").value
    wshENC_Entête.Cells(rowToBeUsed, fEncECustomer).value = wshENC_Saisie.Range("F5").value
    wshENC_Entête.Cells(rowToBeUsed, fEncECodeClient).value = wshENC_Saisie.clientCode
    wshENC_Entête.Cells(rowToBeUsed, fEncEPayType).value = wshENC_Saisie.Range("F7").value
    wshENC_Entête.Cells(rowToBeUsed, fEncEAmount).value = CDbl(Format$(wshENC_Saisie.Range("K7").value, "#,##0.00"))
    wshENC_Entête.Cells(rowToBeUsed, fEncENotes).value = wshENC_Saisie.Range("F9").value
    wshENC_Entête.Cells(rowToBeUsed, fEncETimeStamp).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Entete", "", startTime)

End Sub

Sub ENC_Add_DB_Details(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "ENC_Détails$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
        
    'Build the recordSet
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).value = True And _
            wshENC_Saisie.Range("K" & r).value <> 0 Then
            rs.AddNew
                rs.Fields(fEncDPayID - 1).value = CLng(pmtNo)
                rs.Fields(fEncDInvNo - 1).value = wshENC_Saisie.Range("F" & r).value
                rs.Fields(fEncDCustomer - 1).value = wshENC_Saisie.Range("F5").value
                rs.Fields(fEncDPayDate - 1).value = wshENC_Saisie.Range("K5").value
                rs.Fields(fEncDPayAmount - 1).value = CDbl(Format$(wshENC_Saisie.Range("K" & r).value, "#,##0.00 $"))
                rs.Fields(fEncDTimeStamp - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            'Update the recordset (create the record)
            rs.Update
        End If
    Next r
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", "", startTime)
    
End Sub

Sub ENC_Add_Locally_Details(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", "", 0)
    
    Application.ScreenUpdating = False
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'What is the last used row in ENC_Détails ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_Détails.Cells(wshENC_Détails.Rows.count, 1).End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).value = True And _
            wshENC_Saisie.Range("K" & r).value <> 0 Then
            wshENC_Détails.Range("A" & rowToBeUsed).value = pmtNo
            wshENC_Détails.Range("B" & rowToBeUsed).value = wshENC_Saisie.Range("F" & r).value
            wshENC_Détails.Range("C" & rowToBeUsed).value = wshENC_Saisie.Range("F5").value
            wshENC_Détails.Range("D" & rowToBeUsed).value = wshENC_Saisie.Range("K5").value
            wshENC_Détails.Range("E" & rowToBeUsed).value = CDbl(Format$(wshENC_Saisie.Range("K" & r).value, "#,##0.00"))
            wshENC_Détails.Range("F" & rowToBeUsed).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    Next r
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", "", startTime)

End Sub

Sub ENC_Update_DB_Comptes_Clients(firstRow As Integer, lastRow As Integer) 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Update_DB_Comptes_Clients", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim r As Long
    For r = firstRow To lastRow
        If wshENC_Saisie.Range("B" & r).value = True And _
            wshENC_Saisie.Range("K" & r).value <> 0 Then
            'Open the recordset for the specified invoice
            Dim Inv_No As String
            Inv_No = CStr(Trim$(wshENC_Saisie.Range("F" & r).value))
            
            Dim strSQL As String
            strSQL = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & Inv_No & "'"
            rs.Open strSQL, conn, 2, 3
            If Not rs.EOF Then
                'Mettre à jour Amount_Paid
                rs.Fields(fFacCCTotalPaid - 1).value = rs.Fields(fFacCCTotalPaid - 1).value + CDbl(wshENC_Saisie.Range("K" & r).value)
                'Mettre à jour Status
                If rs.Fields(fFacCCTotal - 1).value - rs.Fields(fFacCCTotalPaid - 1).value = 0 Then
                    On Error Resume Next
                    rs.Fields(fFacCCStatus - 1).value = "Paid"
                    If Err.Number <> 0 Then
                        msgBox "Erreur #" & Err.Number & " : " & Err.Description
                    End If
                    On Error GoTo 0
                Else
                    rs.Fields(fFacCCStatus - 1).value = "Unpaid"
                End If
                'Mettre à jour le solde de la facture
                rs.Fields(fFacCCBalance - 1).value = rs.Fields(fFacCCTotal - 1).value - rs.Fields(fFacCCTotalPaid - 1).value + rs.Fields(fFacCCTotalRegul - 1).value
                rs.Update
            Else
                'Handle the case where the specified ID is not found
                msgBox "L'enregistrement avec la facture '" & Inv_No & "' ne peut être retrouvé!", _
                    vbExclamation
                GoTo Clean_Exit
            End If
            'Update the recordset (create the record)
            rs.Update
            rs.Close
        End If
    Next r
    
Clean_Exit:
    
    'Close recordset and connection
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Update_DB_Comptes_Clients", "", startTime)
    
End Sub

Sub ENC_Update_Locally_Comptes_Clients(firstRow As Integer, lastRow As Integer) '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Update_Locally_Comptes_Clients", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim r As Integer
    For r = firstRow To lastRow
        Dim Inv_No As String
        Inv_No = CStr(wshENC_Saisie.Range("F" & r).value)
        
        Dim foundRange As Range
        Set foundRange = lookupRange.Find(What:=Inv_No, LookIn:=xlValues, LookAt:=xlWhole)
    
        Dim rowToBeUpdated As Long
        If Not foundRange Is Nothing Then
            rowToBeUpdated = foundRange.row
            ws.Cells(rowToBeUpdated, fFacCCTotalPaid).value = ws.Cells(rowToBeUpdated, fFacCCTotalPaid).value + wshENC_Saisie.Range("K" & r).value
            ws.Cells(rowToBeUpdated, fFacCCBalance).value = ws.Cells(rowToBeUpdated, fFacCCBalance).value - wshENC_Saisie.Range("K" & r).value
            'Est-ce que le solde de la facture est à 0,00 $ ?
            If ws.Cells(rowToBeUpdated, fFacCCBalance).value = 0 Then
                ws.Cells(rowToBeUpdated, fFacCCStatus) = "Paid"
            Else
                ws.Cells(rowToBeUpdated, fFacCCStatus) = "Unpaid"
            End If
        Else
            msgBox "La facture '" & Inv_No & "' n'existe pas dans FAC_Comptes_Clients.", vbCritical
        End If
    Next r
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modENC_Saisie:ENC_Update_Locally_Comptes_Clients", "", startTime)

End Sub

Sub ENC_GL_Posting_DB(no As String, dt As Date, nom As String, typeE As String, montant As Currency, desc As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_GL_Posting_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans$"
    
    'Initialize connection, connection string, open the connection & declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJNo As Long
    strSQL = "SELECT MAX(NoEntrée) AS MaxEJNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new ID
    gNextJENo = lastJE + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim timeStamp As Date
    timeStamp = Now
    
    'Debit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields(fGlTNoEntrée - 1).value = gNextJENo
        rs.Fields(fGlTDate - 1).value = Format$(dt, "yyyy-mm-dd")
        If wshENC_Saisie.Range("F7").value = "Dépôt de client" Then
            rs.Fields(fGlTDescription - 1).value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            rs.Fields(fGlTSource - 1).value = UCase$(wshENC_Saisie.Range("F7").value) & ":" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).value = ObtenirNoGlIndicateur("Produit perçu d'avance")
            rs.Fields(fGlTCompte - 1).value = "Produit perçu d'avance" 'Hardcoded
        Else
            rs.Fields(fGlTDescription - 1).value = nom
            rs.Fields(fGlTSource - 1).value = "ENCAISSEMENT:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).value = ObtenirNoGlIndicateur("Encaisse")
            rs.Fields(fGlTCompte - 1).value = "Encaisse" 'Hardcoded
        End If
        rs.Fields(fGlTDébit - 1).value = montant
        rs.Fields(fGlTAutreRemarque - 1).value = desc
        rs.Fields(fGlTTimeStamp - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    rs.Update
    
    'Credit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields(fGlTNoEntrée - 1).value = gNextJENo
        rs.Fields(fGlTDate - 1).value = Format$(dt, "yyyy-mm-dd")
        If wshENC_Saisie.Range("F7").value = "Dépôt de client" Then
            rs.Fields(fGlTDescription - 1).value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            rs.Fields(fGlTSource - 1).value = UCase$(wshENC_Saisie.Range("F7").value) & ":" & Format$(no, "00000")
        Else
            rs.Fields(fGlTDescription - 1).value = nom
            rs.Fields(fGlTSource - 1).value = "ENCAISSEMENT:" & Format$(no, "00000")
        End If
        rs.Fields(fGlTNoCompte - 1).value = ObtenirNoGlIndicateur("Comptes Clients")
        rs.Fields(fGlTCompte - 1).value = "Comptes clients" 'Hardcoded
        rs.Fields(fGlTCrédit - 1).value = montant
        rs.Fields(fGlTAutreRemarque - 1).value = desc
        rs.Fields(fGlTTimeStamp - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    rs.Update

    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modENC_Saisie:ENC_GL_Posting_DB", "", startTime)

End Sub

Sub ENC_GL_Posting_Locally(no As String, dt As Date, nom As String, typeE As String, montant As Currency, desc As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_GL_Posting_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Cells(wshGL_Trans.Rows.count, 1).End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    gNextJENo = wshENC_Saisie.Range("B10").value
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    With wshGL_Trans
    'Debit side
        .Range("A" & rowToBeUsed).value = gNextJENo
        .Range("B" & rowToBeUsed).value = CDate(dt)
        If wshENC_Saisie.Range("F7").value = "Dépôt de client" Then
            .Range("C" & rowToBeUsed).value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            .Range("D" & rowToBeUsed).value = UCase$(wshENC_Saisie.Range("F7").value) & ":" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = ObtenirNoGlIndicateur("Produit perçu d'avance")
            .Range("F" & rowToBeUsed).value = "Produit perçu d'avance" 'Hardcoded
        Else
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "ENCAISSEMENT:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = ObtenirNoGlIndicateur("Encaisse")
            .Range("F" & rowToBeUsed).value = "Encaisse" 'Hardcoded
        End If
        .Range("G" & rowToBeUsed).value = montant
        .Range("I" & rowToBeUsed).value = desc
        .Range("J" & rowToBeUsed).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    
    'Credit side
        .Range("A" & rowToBeUsed).value = gNextJENo
        .Range("B" & rowToBeUsed).value = CDate(dt)
        If wshENC_Saisie.Range("F7").value = "Dépôt de client" Then
            .Range("C" & rowToBeUsed).value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            .Range("D" & rowToBeUsed).value = UCase$(wshENC_Saisie.Range("F7").value) & ":" & Format$(no, "00000")
        Else
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "ENCAISSEMENT:" & Format$(no, "00000")
        End If
        .Range("E" & rowToBeUsed).value = ObtenirNoGlIndicateur("Comptes Clients")
        .Range("F" & rowToBeUsed).value = "Comptes clients" 'Hardcoded
        .Range("H" & rowToBeUsed).value = montant
        .Range("I" & rowToBeUsed).value = desc
        .Range("J" & rowToBeUsed).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modENC_Saisie:ENC_GL_Posting_Locally", "", startTime)

End Sub

Sub ENC_Add_Check_Boxes(row As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Check_Boxes", "", 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshENC_Saisie
    
    Dim chkBoxRange As Range: Set chkBoxRange = ws.Range("E12:E" & 12 + row)
    
    Dim cell As Range
    Dim cbx As checkBox
    For Each cell In chkBoxRange
    'Check if the cell is empty and doesn't have a checkbox already
    If cell.row <= 36 And _
        ActiveSheet.Cells(cell.row, 2).value = "" And _
        ActiveSheet.Cells(cell.row, 6).value <> "" Then 'Applied = False
            'Create a checkbox linked to the cell
            Set cbx = wshENC_Saisie.CheckBoxes.Add(cell.Left + 30, cell.Top, cell.Width, cell.Height)
            With cbx
                .Name = "chkBox - " & cell.row
                .Caption = ""
                .value = False
                .linkedCell = "B" & cell.row
                .Display3DShading = True
                .OnAction = "chkBox_Apply_Click"
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
    
    Call Log_Record("modENC_Saisie:ENC_Add_Check_Boxes", "", startTime)

End Sub

Sub ENC_Remove_Check_Boxes(row As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Remove_Check_Boxes", "", 0)
    
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
    
    Call Log_Record("modENC_Saisie:ENC_Remove_Check_Boxes", "", startTime)

End Sub

Sub ENC_Clear_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Clear_Cells", "", 0)
    
    wshENC_Saisie.Unprotect
    
    With wshENC_Saisie
    
        Application.EnableEvents = False
        
        .Range("B5,F5:H5,K7,F9:I9,E12:K36").ClearContents 'Clear Fields
        .Range("B12:B36").ClearContents
        
        .Range("K5").value = ""
        .Range("F7").value = "Banque" 'Set Default type
        .Range("F5").Activate
        
    End With
    
    'Note the lastUsedRow for checkBox deletion
    Dim lastUsedRow As Long
    lastUsedRow = wshENC_Saisie.Cells(wshENC_Saisie.Rows.count, "F").End(xlUp).row
    If lastUsedRow > 36 Then
        lastUsedRow = 36
    End If
    If lastUsedRow > 11 Then
        Call ENC_Remove_Check_Boxes(lastUsedRow)
    End If
        
    With wshENC_Saisie.Range("F5:H5, K5, F7, K7, F9:I9").Interior '2024-08-25 @ 09:21
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    
    wshENC_Saisie.Shapes("btnENC_Sauvegarde").Visible = False
    wshENC_Saisie.Shapes("btnENC_Annule").Visible = False
    
    Application.EnableEvents = True
    
    With wshENC_Saisie
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

    Call Log_Record("modENC_Saisie:ENC_Clear_Cells", "", startTime)

End Sub

Sub chkBox_Apply_Click()

    Dim chkBox As checkBox
    Set chkBox = ActiveSheet.CheckBoxes(Application.Caller)
    Dim linkedCell As Range
    Set linkedCell = ActiveSheet.Range(chkBox.linkedCell)
    
    If linkedCell.value = True Then
        If wshENC_Saisie.Range("K9").value > 0 Then
            Application.EnableEvents = False
            If wshENC_Saisie.Range("K9").value > wshENC_Saisie.Range("J" & linkedCell.row).value Then
                wshENC_Saisie.Range("K" & linkedCell.row).value = wshENC_Saisie.Range("J" & linkedCell.row).value
            Else
                wshENC_Saisie.Range("K" & linkedCell.row).value = wshENC_Saisie.Range("K9").value
            End If
            Application.EnableEvents = True
        End If
        wshENC_Saisie.Shapes("btnENC_Sauvegarde").Visible = True
        wshENC_Saisie.Shapes("btnENC_Annule").Visible = True
    Else
        ActiveSheet.Range("K" & linkedCell.row).value = 0
    End If

    'Libérer la mémoire
    Set chkBox = Nothing
    Set linkedCell = Nothing
    
End Sub

Sub shp_ENC_Exit_Click()

    Call ENC_Back_To_GL_Menu

End Sub

Sub ENC_Back_To_GL_Menu()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Back_To_GL_Menu", "", 0)
   
    If wshENC_Saisie.ProtectContents Then
        wshENC_Saisie.Unprotect
    End If
    
    Application.EnableEvents = False
    
    Call ENC_Clear_Cells
    
    Application.EnableEvents = True
    
    wshENC_Saisie.Visible = xlSheetVeryHidden

    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call Log_Record("modENC_Saisie:ENC_Back_To_GL_Menu", "", startTime)

End Sub

Sub AjusteLibelléEncaissement(typeTrans As String)

    Application.EnableEvents = False
    
    If Not typeTrans = "Régularisations" Then
        wshENC_Saisie.Range("J5").value = "Date encaissement:"
        wshENC_Saisie.Range("J5").Font.Color = vbBlack
        wshENC_Saisie.Range("J7").value = "Total encaissement:"
        wshENC_Saisie.Range("J7").Font.Color = vbBlack
    Else
        wshENC_Saisie.Range("J5").value = "Date RÉGULARISATION:"
        wshENC_Saisie.Range("J5").Font.Color = vbRed
        wshENC_Saisie.Range("J7").value = "Total RÉGULARISATION:"
        wshENC_Saisie.Range("J7").Font.Color = vbRed
    End If

    Application.EnableEvents = True

End Sub

Sub ValiderEtLancerufEncRégularisation()

    Dim ws As Worksheet
    Set ws = wshENC_Saisie
    
    'Vérification des champs obligatoires
    If IsEmpty(ws.Range("F5").value) Then
        msgBox "Le client est obligatoire. Veuillez le choisir avant de continuer.", vbExclamation
        Exit Sub
    End If

    If IsEmpty(ws.Range("K5").value) Then
        msgBox "La date est obligatoire. Veuillez la saisir avant de continuer.", vbExclamation
        Exit Sub
    End If
    
    If ws.Range("K7").value = 0 Then
        msgBox "Le montant de la régularisation est obligatoire. Veuillez le fournir avant de continuer.", vbExclamation
        Exit Sub
    End If
    
    'Condition pour lancer le UserForm
    If ws.Range("F7").value = "Régularisations" Then
        ' Lancer le UserForm
        ufEncRégularisation.show
    End If
    
End Sub


