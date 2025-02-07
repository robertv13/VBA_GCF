Attribute VB_Name = "modENC_Saisie"
Option Explicit

'Variables globales pour le module
Dim lastRow As Long, lastResultRow As Long
Dim payRow As Long

Sub ENC_Get_OS_Invoices(cc As String) '2024-08-21 @ 15:18
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices", "", 0)
    
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
'            If .Range("U" & i).Value <> 0 And _
'                            Fn_Invoice_Is_Confirmed(.Range("Q" & i).Value) = True Then
            If .Range("X" & i).Value <> 0 And _
                            Fn_Invoice_Is_Confirmed(.Range("S" & i).Value) = True Then
                Application.EnableEvents = False
                wshENC_Saisie.Range("F" & rr).Value = .Range("S" & i).Value
                wshENC_Saisie.Range("G" & rr).Value = Format$(.Range("T" & i).Value, wshAdmin.Range("B1").Value)
                wshENC_Saisie.Range("H" & rr).Value = .Range("U" & i).Value
                wshENC_Saisie.Range("I" & rr).Value = .Range("V" & i).Value + .Range("W" & i).Value
                wshENC_Saisie.Range("J" & rr).Value = .Range("X" & i).Value
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

    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices_With_AF", "", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    'Effacer les données de la dernière utilisation
    ws.Range("O6:O10").ClearContents
    ws.Range("O6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")

    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("tblFAC_Comptes_Clients[#All]")
    ws.Range("O7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("O2:P3")
    ws.Range("O3").Value = wshENC_Saisie.clientCode
    ws.Range("O8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
'    Set rngResult = ws.Range("P1").CurrentRegion
    Set rngResult = ws.Range("R1").CurrentRegion
    rngResult.offset(2, 0).Clear
'    Set rngResult = ws.Range("P2:U2")
    Set rngResult = ws.Range("R2:X2")
    ws.Range("O9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
                                        
    'Est-ce que nous avons des résultats ?
'    lastResultRow = ws.Cells(ws.Rows.count, "P").End(xlUp).row
    lastResultRow = ws.Cells(ws.Rows.count, "R").End(xlUp).row
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
    
'    'PLUG - Recalculate Column 'U' - Balance after AdvancedFilter
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
    
    Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices_With_AF", "", startTime)

End Sub

Sub shp_ENC_Update_Click()

    Call MAJ_Encaissement

End Sub

Sub MAJ_Encaissement() '2024-08-22 @ 09:46
    
    startTime = Timer: Call Log_Record("modENC_Saisie:MAJ_Encaissement", "", 0)
    
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
        
        'Create records for ENC_Entête
        Call ENC_Add_DB_Entete
        Call ENC_Add_Locally_Entete
        
'        Dim pmtNo As Long
'        pmtNo = wshENC_Saisie.pmtNo
        
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
        Dim noEnc As String, nomClient As String, typeEnc As String, descEnc As String
        Dim dateEnc As Date
        Dim montantEnc As Currency
        noEnc = wshENC_Saisie.pmtNo
        dateEnc = wshENC_Saisie.Range("K5").Value
        nomClient = wshENC_Saisie.Range("F5").Value
        typeEnc = wshENC_Saisie.Range("F7").Value
        montantEnc = wshENC_Saisie.Range("K7").Value
        descEnc = wshENC_Saisie.Range("F9").Value

        Call ENC_GL_Posting_DB(noEnc, dateEnc, nomClient, typeEnc, montantEnc, descEnc)  '2024-08-22 @ 16:08
        Call ENC_GL_Posting_Locally(noEnc, dateEnc, nomClient, typeEnc, montantEnc, descEnc)  '2024-08-22 @ 16:08
        
        MsgBox "L'encaissement '" & wshENC_Saisie.pmtNo & "' a été enregistré avec succès", vbOKOnly + vbInformation
        
        Call Encaissement_Add_New 'Reset the form
        
        .Range("F5").Select
    End With
    
Clean_Exit:

    Call Log_Record("modENC_Saisie:MAJ_Encaissement", "", startTime)

End Sub

Sub Encaissement_Add_New() '2024-08-21 @ 14:58

    startTime = Timer: Call Log_Record("modENC_Saisie:Encaissement_Add_New", "", 0)

    Call ENC_Clear_Cells
    
    Call Log_Record("modENC_Saisie:Encaissement_Add_New", "", startTime)
    
End Sub

Sub ENC_Add_DB_Entete() 'Write to MASTER.xlsx
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Entete", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
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
    If IsNull(rs.Fields("MaxPmtNo").Value) Then
        'Handle empty table (assign a default value, e.g., 1)
        lr = 0
    Else
        lr = rs.Fields("MaxPmtNo").Value
    End If
    
    'Calculate the new PmtNo
    wshENC_Saisie.pmtNo = lr + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
        rs.Fields(fEncEPayID - 1).Value = wshENC_Saisie.pmtNo
        rs.Fields(fEncEPayDate - 1).Value = wshENC_Saisie.Range("K5").Value
        rs.Fields(fEncECustomer - 1).Value = wshENC_Saisie.Range("F5").Value
        rs.Fields(fEncECodeClient - 1).Value = wshENC_Saisie.clientCode
        rs.Fields(fEncEPayType - 1).Value = wshENC_Saisie.Range("F7").Value
        rs.Fields(fEncEAmount - 1).Value = CDbl(Format$(wshENC_Saisie.Range("K7").Value, "#,##0.00 $"))
        rs.Fields(fEncENotes - 1).Value = wshENC_Saisie.Range("F9").Value
        rs.Fields(fEncETimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
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
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Locally_Entete", "", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentPmtNo As Long
    currentPmtNo = wshENC_Saisie.pmtNo
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_Entête.Cells(wshENC_Entête.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    wshENC_Entête.Cells(rowToBeUsed, fEncEPayID).Value = currentPmtNo
    wshENC_Entête.Cells(rowToBeUsed, fEncEPayDate).Value = wshENC_Saisie.Range("K5").Value
    wshENC_Entête.Cells(rowToBeUsed, fEncECustomer).Value = wshENC_Saisie.Range("F5").Value
    wshENC_Entête.Cells(rowToBeUsed, fEncECodeClient).Value = wshENC_Saisie.clientCode
    wshENC_Entête.Cells(rowToBeUsed, fEncEPayType).Value = wshENC_Saisie.Range("F7").Value
    wshENC_Entête.Cells(rowToBeUsed, fEncEAmount).Value = CDbl(Format$(wshENC_Saisie.Range("K7").Value, "#,##0.00"))
    wshENC_Entête.Cells(rowToBeUsed, fEncENotes).Value = wshENC_Saisie.Range("F9").Value
    wshENC_Entête.Cells(rowToBeUsed, fEncETimeStamp).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Entete", "", startTime)

End Sub

Sub ENC_Add_DB_Details(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) 'Write to MASTER.xlsx
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "ENC_Détails$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
        
    'Build the recordSet
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).Value = True And _
            wshENC_Saisie.Range("K" & r).Value <> 0 Then
            rs.AddNew
                rs.Fields(fEncDPayID - 1).Value = CLng(pmtNo)
                rs.Fields(fEncDInvNo - 1).Value = wshENC_Saisie.Range("F" & r).Value
                rs.Fields(fEncDCustomer - 1).Value = wshENC_Saisie.Range("F5").Value
                rs.Fields(fEncDPayDate - 1).Value = wshENC_Saisie.Range("K5").Value
                rs.Fields(fEncDPayAmount - 1).Value = CDbl(Format$(wshENC_Saisie.Range("K" & r).Value, "#,##0.00 $"))
                rs.Fields(fEncDTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
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
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", "", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in ENC_Détails ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_Détails.Cells(wshENC_Détails.Rows.count, 1).End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).Value = True And _
            wshENC_Saisie.Range("K" & r).Value <> 0 Then
            wshENC_Détails.Range("A" & rowToBeUsed).Value = pmtNo
            wshENC_Détails.Range("B" & rowToBeUsed).Value = wshENC_Saisie.Range("F" & r).Value
            wshENC_Détails.Range("C" & rowToBeUsed).Value = wshENC_Saisie.Range("F5").Value
            wshENC_Détails.Range("D" & rowToBeUsed).Value = wshENC_Saisie.Range("K5").Value
            wshENC_Détails.Range("E" & rowToBeUsed).Value = CDbl(Format$(wshENC_Saisie.Range("K" & r).Value, "#,##0.00"))
            wshENC_Détails.Range("F" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            rowToBeUsed = rowToBeUsed + 1
        End If
    Next r
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", "", startTime)

End Sub

Sub ENC_Update_DB_Comptes_Clients(firstRow As Integer, lastRow As Integer) 'Write to MASTER.xlsx
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Update_DB_Comptes_Clients", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim r As Long
    For r = firstRow To lastRow
        If wshENC_Saisie.Range("B" & r).Value = True And _
            wshENC_Saisie.Range("K" & r).Value <> 0 Then
            'Open the recordset for the specified invoice
            Dim Inv_No As String
            Inv_No = CStr(Trim(wshENC_Saisie.Range("F" & r).Value))
            
            Dim strSQL As String
            strSQL = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & Inv_No & "'"
            rs.Open strSQL, conn, 2, 3
            If Not rs.EOF Then
                'Mettre à jour Amount_Paid
                rs.Fields(fFacCCTotalPaid - 1).Value = rs.Fields(fFacCCTotalPaid - 1).Value + CDbl(wshENC_Saisie.Range("K" & r).Value)
                'Mettre à jour Status
                If rs.Fields(fFacCCTotal - 1).Value - rs.Fields(fFacCCTotalPaid - 1).Value = 0 Then
                    On Error Resume Next
                    rs.Fields(fFacCCStatus - 1).Value = "Paid"
                    If Err.Number <> 0 Then
                        MsgBox "Erreur #" & Err.Number & " : " & Err.Description
                    End If
                    On Error GoTo 0
                Else
                    rs.Fields(fFacCCStatus - 1).Value = "Unpaid"
                End If
                'Mettre à jour le solde de la facture
'                rs.Fields(fFacCCBalance - 1).Value = rs.Fields(fFacCCTotal - 1).Value - rs.Fields(fFacCCTotalPaid - 1).Value
                rs.Fields(fFacCCBalance - 1).Value = rs.Fields(fFacCCTotal - 1).Value - rs.Fields(fFacCCTotalPaid - 1).Value + rs.Fields(fFacCCTotalRegul - 1).Value
                rs.Update
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec la facture '" & Inv_No & "' ne peut être retrouvé!", _
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
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Update_Locally_Comptes_Clients", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim r As Integer
    For r = firstRow To lastRow
        Dim Inv_No As String
        Inv_No = CStr(wshENC_Saisie.Range("F" & r).Value)
        
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
    
    Call Log_Record("modENC_Saisie:ENC_Update_Locally_Comptes_Clients", "", startTime)

End Sub

Sub ENC_GL_Posting_DB(no As String, dt As Date, nom As String, typeE As String, montant As Currency, desc As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_GL_Posting_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
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
    If IsNull(rs.Fields("MaxEJNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").Value
    End If
    
    'Calculate the new ID
    Dim nextJENo As Long
    nextJENo = lastJE + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Debit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields(fGlTNoEntrée - 1).Value = nextJENo
        rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
        If wshENC_Saisie.Range("F7").Value = "Dépôt de client" Then
            rs.Fields(fGlTDescription - 1).Value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            rs.Fields(fGlTSource - 1).Value = UCase(wshENC_Saisie.Range("F7").Value) & ":" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).Value = ObtenirNoGlIndicateur("Produit perçu d'avance")
            rs.Fields(fGlTCompte - 1).Value = "Produit perçu d'avance" 'Hardcoded
        Else
            rs.Fields(fGlTDescription - 1).Value = nom
            rs.Fields(fGlTSource - 1).Value = "ENCAISSEMENT:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).Value = ObtenirNoGlIndicateur("Encaisse")
            rs.Fields(fGlTCompte - 1).Value = "Encaisse" 'Hardcoded
        End If
        rs.Fields(fGlTDébit - 1).Value = montant
        rs.Fields(fGlTAutreRemarque - 1).Value = desc
        rs.Fields(fGlTTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    rs.Update
    
    'Credit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields(fGlTNoEntrée - 1).Value = nextJENo
        rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
        If wshENC_Saisie.Range("F7").Value = "Dépôt de client" Then
            rs.Fields(fGlTDescription - 1).Value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            rs.Fields(fGlTSource - 1).Value = UCase(wshENC_Saisie.Range("F7").Value) & ":" & Format$(no, "00000")
        Else
            rs.Fields(fGlTDescription - 1).Value = nom
            rs.Fields(fGlTSource - 1).Value = "ENCAISSEMENT:" & Format$(no, "00000")
        End If
        rs.Fields(fGlTNoCompte - 1).Value = ObtenirNoGlIndicateur("Comptes Clients")
        rs.Fields(fGlTCompte - 1).Value = "Comptes clients" 'Hardcoded
        rs.Fields(fGlTCrédit - 1).Value = montant
        rs.Fields(fGlTAutreRemarque - 1).Value = desc
        rs.Fields(fGlTTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
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
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_GL_Posting_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Cells(wshGL_Trans.Rows.count, 1).End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim nextJENo As Long
    nextJENo = wshENC_Saisie.Range("B10").Value
    
    With wshGL_Trans
    'Debit side
        .Range("A" & rowToBeUsed).Value = nextJENo
        .Range("B" & rowToBeUsed).Value = CDate(dt)
        If wshENC_Saisie.Range("F7").Value = "Dépôt de client" Then
            .Range("C" & rowToBeUsed).Value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            .Range("D" & rowToBeUsed).Value = UCase(wshENC_Saisie.Range("F7").Value) & ":" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = ObtenirNoGlIndicateur("Produit perçu d'avance")
            .Range("F" & rowToBeUsed).Value = "Produit perçu d'avance" 'Hardcoded
        Else
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "ENCAISSEMENT:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = ObtenirNoGlIndicateur("Encaisse")
            .Range("F" & rowToBeUsed).Value = "Encaisse" 'Hardcoded
        End If
        .Range("G" & rowToBeUsed).Value = montant
        .Range("I" & rowToBeUsed).Value = desc
        .Range("J" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    
    'Credit side
        .Range("A" & rowToBeUsed).Value = nextJENo
        .Range("B" & rowToBeUsed).Value = CDate(dt)
        If wshENC_Saisie.Range("F7").Value = "Dépôt de client" Then
            .Range("C" & rowToBeUsed).Value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            .Range("D" & rowToBeUsed).Value = UCase(wshENC_Saisie.Range("F7").Value) & ":" & Format$(no, "00000")
        Else
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "ENCAISSEMENT:" & Format$(no, "00000")
        End If
        .Range("E" & rowToBeUsed).Value = ObtenirNoGlIndicateur("Comptes Clients")
        .Range("F" & rowToBeUsed).Value = "Comptes clients" 'Hardcoded
        .Range("H" & rowToBeUsed).Value = montant
        .Range("I" & rowToBeUsed).Value = desc
        .Range("J" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modENC_Saisie:ENC_GL_Posting_Locally", "", startTime)

End Sub

Sub ENC_Add_Check_Boxes(row As Long)

    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Check_Boxes", "", 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshENC_Saisie
    
    Dim chkBoxRange As Range: Set chkBoxRange = ws.Range("E12:E" & 12 + row)
    
    Dim cell As Range
    Dim cbx As checkBox
    For Each cell In chkBoxRange
    'Check if the cell is empty and doesn't have a checkbox already
    If cell.row <= 36 And _
        Cells(cell.row, 2).Value = "" And _
        Cells(cell.row, 6).Value <> "" Then 'Applied = False
            'Create a checkbox linked to the cell
            Set cbx = wshENC_Saisie.CheckBoxes.Add(cell.Left + 30, cell.Top, cell.Width, cell.Height)
            With cbx
                .Name = "chkBox - " & cell.row
                .Caption = ""
                .Value = False
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

    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Remove_Check_Boxes", "", 0)
    
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

    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Clear_Cells", "", 0)
    
    wshENC_Saisie.Unprotect
    
    With wshENC_Saisie
    
        Application.EnableEvents = False
        
        .Range("B5,F5:H5,K7,F9:I9,E12:K36").ClearContents 'Clear Fields
        .Range("B12:B36").ClearContents
        
        .Range("K5").Value = ""
        .Range("F7").Value = "Banque" 'Set Default type
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
    Set linkedCell = Range(chkBox.linkedCell)
    
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
        Range("K" & linkedCell.row).Value = 0
    End If

    'Libérer la mémoire
    Set chkBox = Nothing
    Set linkedCell = Nothing
    
End Sub

Sub shp_ENC_Exit_Click()

    Call ENC_Back_To_GL_Menu

End Sub

Sub ENC_Back_To_GL_Menu()
    
    startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Back_To_GL_Menu", "", 0)
   
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

Sub ValiderEtLancerufEncRégularisation()

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
        MsgBox "Le montant de la régularisation est obligatoire. Veuillez le fournir avant de continuer.", vbExclamation
        Exit Sub
    End If
    
    'Condition pour lancer le UserForm
    If ws.Range("F7").Value = "Régularisations" Then
        ' Lancer le UserForm
        ufEncRégularisation.show
    End If
    
End Sub

