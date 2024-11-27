Attribute VB_Name = "modENC_Saisie"
Option Explicit

Dim lastRow As Long, lastResultRow As Long
Dim payRow As Long

Sub ENC_Get_OS_Invoices(cc As String) '2024-08-21 @ 15:18
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices", 0)
    
    Dim ws As Worksheet
    Set ws = wshENC_Saisie
    
    Application.EnableEvents = False
    ws.Range("E12:K36").ClearContents 'Clear the invoices area before loading it
    Application.EnableEvents = True
    
    Call ENC_Get_OS_Invoices_With_AF(cc)
    
    'Bring the Result from AF into our List of Oustanding Invoices
    Dim lastResultRow As Long
    lastResultRow = wshFAC_Comptes_Clients.Cells(ws.Rows.count, "P").End(xlUp).row
    
    Dim i As Integer
    'Unlock the required area
    With ws '2024-08-21 @ 16:06
        If lastResultRow >= 3 Then
            .Unprotect
            .Range("B12:B" & 11 + lastResultRow - 2).Locked = False
            .Range("E12:E" & 11 + lastResultRow - 2).Locked = False
            .Protect UserInterfaceOnly:=True
            .EnableSelection = xlNoRestrictions
        End If
    End With
    
    'Copy à partir du résultat de AF, dans la feuille de saisie des encaissements
    Dim rr As Integer: rr = 12
    With wshFAC_Comptes_Clients
        For i = 3 To WorksheetFunction.Min(27, lastResultRow) 'No space for more O/S invoices
            If .Range("U" & i).value <> 0 And _
                            Fn_Invoice_Is_Confirmed(.Range("Q" & i).value) = True Then
                Application.EnableEvents = False
                wshENC_Saisie.Range("F" & rr).value = .Range("Q" & i).value
                wshENC_Saisie.Range("G" & rr).value = Format$(.Range("R" & i).value, wshAdmin.Range("B1").value)
                wshENC_Saisie.Range("H" & rr).value = .Range("S" & i).value
                wshENC_Saisie.Range("I" & rr).value = .Range("T" & i).value
                wshENC_Saisie.Range("J" & rr).value = .Range("U" & i).value
                Application.EnableEvents = True
                rr = rr + 1
            End If
        Next i
    End With
    
    Call ENC_Add_Check_Boxes(lastResultRow - 2)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modFAC_Enc:ENC_Load_OS_Invoices", startTime)

End Sub

Sub ENC_Get_OS_Invoices_With_AF(cc As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices_With_AF", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    'Effacer les données de la dernière utilisation
    ws.Range("M6:M10").ClearContents
    ws.Range("M6").value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")

    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("tblFAC_Comptes_Clients[#All]")
    ws.Range("M7").value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("M2:N3")
    ws.Range("M3").value = wshENC_Saisie.clientCode
    ws.Range("M8").value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("P1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("P2:U2")
    ws.Range("M9").value = rngResult.Address
    
    rngData.AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
                                        
    'Est-ce que nous avons des résultats ?
    lastResultRow = ws.Cells(ws.Rows.count, "P").End(xlUp).row
    ws.Range("M10").value = lastResultRow - 2 & " lignes"
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 3 Then
        With ws.Sort 'Sort - InvNo
            .SortFields.Clear
            'First sort On InvNo
            .SortFields.add key:=ws.Range("Q3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange ws.Range("P3:U" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If
    
    'PLUG - Recalculate Column 'U' - Balance after AdvancedFilter
    Dim r As Integer
    For r = 3 To lastResultRow
        ws.Range("U" & r).value = ws.Range("S" & r).value - ws.Range("T" & r).value
    Next r

    'libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Enc:ENC_Get_OS_Invoices_With_AF", startTime)

End Sub

Sub shp_ENC_Update_Click()

    Call MAJ_Encaissement

End Sub

Sub MAJ_Encaissement() '2024-08-22 @ 09:46
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:MAJ_Encaissement", 0)
    
    With wshENC_Saisie
        'Check for mandatory fields (4)
        If .Range("F5").value = Empty Or _
           .Range("K5").value = Empty Or _
           .Range("F7").value = Empty Or _
           .Range("K7").value = 0 Then
            MsgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client valide" & vbNewLine & _
                "2. Une date d'encaissements" & vbNewLine & _
                "3. Un type de paiement et" & vbNewLine & _
                "4. Des montants appliqués" & vbNewLine & vbNewLine & _
                "AVANT de sauvegarder la transaction.", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Check to make sure Payment Amount = Applied Amount
        If .Range("K9").value <> 0 Then
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
        
        MsgBox "L'encaissement '" & wshENC_Saisie.pmtNo & "' a été enregistré avec succès", vbOKOnly + vbInformation
        
        Call Encaissement_Add_New 'Reset the form
        
        .Range("F5").Select
    End With
    
Clean_Exit:

    Call Log_Record("modENC_Saisie:MAJ_Encaissement", startTime)

End Sub

Sub Encaissement_Add_New() '2024-08-21 @ 14:58

    Dim startTime As Double: startTime = Timer: Call Log_Record("modEnc_Saisie:Encaissement_Add_New", 0)

    Call ENC_Clear_Cells
    
    Call Log_Record("modEnc_Saisie:Encaissement_Add_New", startTime)
    
End Sub

Sub ENC_Add_DB_Entete() 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Entete", 0)
    
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
    strSQL = "SELECT MAX(Pay_ID) AS MaxPmtNo FROM [" & destinationTab & "]"

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

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
        rs.Fields("Pay_ID").value = wshENC_Saisie.pmtNo
        rs.Fields("Pay_Date").value = wshENC_Saisie.Range("K5").value
        rs.Fields("Customer").value = wshENC_Saisie.Range("F5").value
        rs.Fields("codeClient").value = wshENC_Saisie.clientCode
        rs.Fields("Pay_Type").value = wshENC_Saisie.Range("F7").value
        rs.Fields("Amount").value = CDbl(Format$(wshENC_Saisie.Range("K7").value, "#,##0.00 $"))
        rs.Fields("Notes").value = wshENC_Saisie.Range("F9").value
    'Update the recordset (create the record)
    rs.update
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_DB_Entete", startTime)
    
End Sub

Sub ENC_Add_Locally_Entete() '2024-08-22 @ 10:38
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentPmtNo As Long
    currentPmtNo = wshENC_Saisie.pmtNo
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_Entête.Cells(wshENC_Entête.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    wshENC_Entête.Range("A" & rowToBeUsed).value = currentPmtNo
    wshENC_Entête.Range("B" & rowToBeUsed).value = wshENC_Saisie.Range("K5").value
    wshENC_Entête.Range("C" & rowToBeUsed).value = wshENC_Saisie.Range("F5").value
    wshENC_Entête.Range("D" & rowToBeUsed).value = wshENC_Saisie.clientCode
    wshENC_Entête.Range("E" & rowToBeUsed).value = wshENC_Saisie.Range("F7").value
    wshENC_Entête.Range("F" & rowToBeUsed).value = CDbl(Format$(wshENC_Saisie.Range("K7").value, "#,##0.00"))
    wshENC_Entête.Range("G" & rowToBeUsed).value = wshENC_Saisie.Range("F9").value
    
    Application.ScreenUpdating = True

    Call Log_Record("modDEB_Saisie:DEB_Trans_Add_Record_Locally", startTime)

End Sub

Sub ENC_Add_DB_Details(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", 0)
    
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

    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
        
    'Build the recordSet
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).value = True And _
            wshENC_Saisie.Range("K" & r).value <> 0 Then
            rs.AddNew
                rs.Fields("Pay_ID").value = CLng(pmtNo)
                rs.Fields("Inv_No").value = wshENC_Saisie.Range("F" & r).value
                rs.Fields("Customer").value = wshENC_Saisie.Range("F5").value
                rs.Fields("Pay_Date").value = wshENC_Saisie.Range("K5").value
                rs.Fields("Pay_Amount").value = CDbl(Format$(wshENC_Saisie.Range("K" & r).value, "#,##0.00 $"))
            'Update the recordset (create the record)
            rs.update
        End If
    Next r
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", startTime)
    
End Sub

Sub ENC_Add_Locally_Details(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", 0)
    
    Application.ScreenUpdating = False
    
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
            rowToBeUsed = rowToBeUsed + 1
        End If
    Next r
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", startTime)

End Sub

Sub ENC_Update_DB_Comptes_Clients(firstRow As Integer, lastRow As Integer) 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", 0)
    
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
            Inv_No = CStr(Trim(wshENC_Saisie.Range("F" & r).value))
            Dim strSQL As String
            strSQL = "SELECT * FROM [" & destinationTab & "] WHERE Invoice_No = '" & Inv_No & "'"
            rs.Open strSQL, conn, 2, 3
            If Not rs.EOF Then
                'Mettre à jour Amount_Paid
                rs.Fields("Total_Paid").value = rs.Fields("Total_Paid").value + CDbl(wshENC_Saisie.Range("K" & r).value)
                'Mettre à jour Status
                If rs.Fields("Total").value - rs.Fields("Total_Paid").value = 0 Then
                    On Error Resume Next
                    rs.Fields("Status").value = "Paid"
                    If Err.Number <> 0 Then
                        MsgBox "Erreur #" & Err.Number & " : " & Err.Description
                    End If
                    On Error GoTo 0
'                    rs.Fields("Status").value = "Paid"
                Else
                    rs.Fields("Status").value = "Unpaid"
                End If
                'Mettre à jour le solde de la facture
                rs.Fields("Balance").value = rs.Fields("Total").value - rs.Fields("Total_Paid").value
                rs.update
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec la facture '" & Inv_No & "' ne peut être retrouvé!", _
                    vbExclamation
                GoTo Clean_Exit
            End If
            'Update the recordset (create the record)
            rs.update
            rs.Close
        End If
    Next r
    
Clean_Exit:
    
    'Close recordset and connection
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", startTime)
    
End Sub

Sub ENC_Update_Locally_Comptes_Clients(firstRow As Integer, lastRow As Integer) '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", 0)
    
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
            ws.Range("I" & rowToBeUpdated).value = ws.Range("I" & rowToBeUpdated).value + wshENC_Saisie.Range("K" & r).value
            ws.Range("J" & rowToBeUpdated).value = ws.Range("H" & rowToBeUpdated).value - wshENC_Saisie.Range("K" & r).value
            'Est-ce que le solde de la facture est à 0,00 $ ?
            If ws.Range("J" & rowToBeUpdated).value = 0 Then
                ws.Range("E" & rowToBeUpdated) = "Paid"
            Else
                ws.Range("E" & rowToBeUpdated) = "Unpaid"
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
    
    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", startTime)

End Sub

Sub ENC_GL_Posting_DB(no As String, dt As Date, nom As String, typeE As String, montant As Currency, DESC As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_GL_Posting_DB", 0)
    
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
    strSQL = "SELECT MAX(No_Entrée) AS MaxEJNo FROM [" & destinationTab & "]"

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
    Dim nextJENo As Long
    nextJENo = lastJE + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Debit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("No_Entrée").value = nextJENo
        rs.Fields("Date").value = Format$(dt, "yyyy-mm-dd")
        If wshENC_Saisie.Range("F7").value = "Dépôt de client" Then
            rs.Fields("Description").value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            rs.Fields("Source").value = UCase(wshENC_Saisie.Range("F7").value) & ":" & Format$(no, "00000")
            rs.Fields("No_Compte").value = "2400" 'Hardcoded
            rs.Fields("Compte").value = "Produit perçu d'avance" 'Hardcoded
        Else
            rs.Fields("Description").value = nom
            rs.Fields("Source").value = "ENCAISSEMENT:" & Format$(no, "00000")
            rs.Fields("No_Compte").value = "1000" 'Hardcoded
            rs.Fields("Compte").value = "Encaisse" 'Hardcoded
        End If
        rs.Fields("Compte").value = "Encaisse" 'Hardcoded
        rs.Fields("Débit").value = montant
        rs.Fields("AutreRemarque").value = DESC
        rs.Fields("TimeStamp").value = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    rs.update
    
    'Credit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("No_Entrée").value = nextJENo
        rs.Fields("Date").value = Format$(dt, "yyyy-mm-dd")
        If wshENC_Saisie.Range("F7").value = "Dépôt de client" Then
            rs.Fields("Description").value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            rs.Fields("Source").value = UCase(wshENC_Saisie.Range("F7").value) & ":" & Format$(no, "00000")
        Else
            rs.Fields("Description").value = nom
            rs.Fields("Source").value = "ENCAISSEMENT:" & Format$(no, "00000")
        End If
        rs.Fields("No_Compte").value = "1100" 'Hardcoded
        rs.Fields("Compte").value = "Comptes clients" 'Hardcoded
        rs.Fields("Crédit").value = montant
        rs.Fields("AutreRemarque").value = DESC
        rs.Fields("TimeStamp").value = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    rs.update

'    wshENC_Saisie.Range("B10").value = nextJENo
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modENC_Saisie:ENC_GL_Posting_DB", startTime)

End Sub

Sub ENC_GL_Posting_Locally(no As String, dt As Date, nom As String, typeE As String, montant As Currency, DESC As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_GL_Posting_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Cells(wshGL_Trans.Rows.count, 1).End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim nextJENo As Long
    nextJENo = wshENC_Saisie.Range("B10").value
    
    With wshGL_Trans
    'Debit side
        .Range("A" & rowToBeUsed).value = nextJENo
        .Range("B" & rowToBeUsed).value = CDate(dt)
        If wshENC_Saisie.Range("F7").value = "Dépôt de client" Then
            .Range("C" & rowToBeUsed).value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            .Range("D" & rowToBeUsed).value = UCase(wshENC_Saisie.Range("F7").value) & ":" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = "2400" 'Hardcoded
            .Range("F" & rowToBeUsed).value = "Produit perçu d'avance" 'Hardcoded
        Else
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "ENCAISSEMENT:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).value = "1000" 'Hardcoded
            .Range("F" & rowToBeUsed).value = "Encaisse" 'Hardcoded
        End If
        .Range("G" & rowToBeUsed).value = montant
        .Range("I" & rowToBeUsed).value = DESC
        .Range("J" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    
    'Credit side
        .Range("A" & rowToBeUsed).value = nextJENo
        .Range("B" & rowToBeUsed).value = CDate(dt)
        If wshENC_Saisie.Range("F7").value = "Dépôt de client" Then
            .Range("C" & rowToBeUsed).value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            .Range("D" & rowToBeUsed).value = UCase(wshENC_Saisie.Range("F7").value) & ":" & Format$(no, "00000")
        Else
            .Range("C" & rowToBeUsed).value = nom
            .Range("D" & rowToBeUsed).value = "ENCAISSEMENT:" & Format$(no, "00000")
        End If
        .Range("E" & rowToBeUsed).value = "1100" 'Hardcoded
        .Range("F" & rowToBeUsed).value = "Comptes clients" 'Hardcoded
        .Range("H" & rowToBeUsed).value = montant
        .Range("I" & rowToBeUsed).value = DESC
        .Range("J" & rowToBeUsed).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modENC_Saisie:ENC_GL_Posting_Locally", startTime)

End Sub

Sub ENC_Add_Check_Boxes(row As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Check_Boxes", 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshENC_Saisie
    
    Dim chkBoxRange As Range: Set chkBoxRange = ws.Range("E12:E" & 11 + row)
    
    Dim cell As Range
    Dim cbx As checkBox
    For Each cell In chkBoxRange
    'Check if the cell is empty and doesn't have a checkbox already
    If cell.row <= 36 And _
        Cells(cell.row, 2).value = "" And _
        Cells(cell.row, 6).value <> "" Then 'Applied = False
            'Create a checkbox linked to the cell
            Set cbx = wshENC_Saisie.CheckBoxes.add(cell.Left + 30, cell.Top, cell.Width, cell.Height)
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
        .EnableSelection = xlNoRestrictions
    End With
    
    Application.EnableEvents = True

    'Libérer la mémoire
    Set cbx = Nothing
    Set cell = Nothing
    Set chkBoxRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modENC_Saisie:ENC_Add_Check_Boxes", startTime)

End Sub

Sub ENC_Remove_Check_Boxes(row As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Remove_Check_Boxes", 0)
    
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
    
    Call Log_Record("modENC_Saisie:ENC_Remove_Check_Boxes", startTime)

End Sub

Sub ENC_Clear_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Clear_Cells", 0)
    
    wshENC_Saisie.Unprotect
    
    With wshENC_Saisie
    
        Application.EnableEvents = False
        
        'Note the lastUsedRow for checkBox deletion
        Dim lastUsedRow As Long
        lastUsedRow = wshENC_Saisie.Cells(wshENC_Saisie.Rows.count, "F").End(xlUp).row
        If lastUsedRow > 36 Then
            lastUsedRow = 36
        End If
        
        .Range("B5,F5:H5,K5,F7,K7,F9:I9,E12:K36").ClearContents 'Clear Fields
        .Range("B12:B36").ClearContents
        
        If lastUsedRow > 11 Then
            Call ENC_Remove_Check_Boxes(lastUsedRow)
        End If
        
        .Range("K5").value = ""
        .Range("F7").value = "Banque" ' Set Default type
        .Range("F5").Activate
    End With
    
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

    Call Log_Record("modENC_Saisie:ENC_Clear_Cells", startTime)

End Sub

Sub chkBox_Apply_Click()

    Dim chkBox As checkBox
    Set chkBox = ActiveSheet.CheckBoxes(Application.Caller)
    Dim linkedCell As Range
    Set linkedCell = Range(chkBox.linkedCell)
    
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
        Range("K" & linkedCell.row).value = 0
    End If

    'Libérer la mémoire
    Set chkBox = Nothing
    Set linkedCell = Nothing
    
End Sub

Sub shp_ENC_Exit_Click()

    Call ENC_Back_To_FAC_Menu

End Sub

Sub ENC_Back_To_FAC_Menu()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Back_To_FAC_Menu", 0)
   
    wshENC_Saisie.Unprotect
    
    Application.EnableEvents = False
    
    Call ENC_Clear_Cells
    
    Application.EnableEvents = True
    
    wshENC_Saisie.Visible = xlSheetVeryHidden

    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call Log_Record("modENC_Saisie:ENC_Back_To_FAC_Menu", startTime)

End Sub


