Attribute VB_Name = "modENC_Saisie"
Option Explicit

'Variables globales pour le module
Dim LastRow As Long, lastResultRow As Long
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
    
    'Copy � partir du r�sultat de AF, dans la feuille de saisie des encaissements
    Dim rr As Integer: rr = 12
    With wshFAC_Comptes_Clients
        For i = 3 To WorksheetFunction.Min(27, lastResultRow) 'No space for more O/S invoices
            If .Range("U" & i).Value <> 0 And _
                            Fn_Invoice_Is_Confirmed(.Range("Q" & i).Value) = True Then
                Application.EnableEvents = False
                wshENC_Saisie.Range("F" & rr).Value = .Range("Q" & i).Value
                wshENC_Saisie.Range("G" & rr).Value = Format$(.Range("R" & i).Value, wshAdmin.Range("B1").Value)
                wshENC_Saisie.Range("H" & rr).Value = .Range("S" & i).Value
                wshENC_Saisie.Range("I" & rr).Value = .Range("T" & i).Value
                wshENC_Saisie.Range("J" & rr).Value = .Range("U" & i).Value
                Application.EnableEvents = True
                rr = rr + 1
            End If
        Next i
    End With
    
    Call ENC_Add_Check_Boxes(lastResultRow - 2)
    
    'Lib�rer la m�moire
    Set ws = Nothing
    
    Call Log_Record("modFAC_Enc:ENC_Load_OS_Invoices", startTime)

End Sub

Sub ENC_Get_OS_Invoices_With_AF(cc As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Get_OS_Invoices_With_AF", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    'Effacer les donn�es de la derni�re utilisation
    ws.Range("M6:M10").ClearContents
    ws.Range("M6").Value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")

    'D�finir le range pour la source des donn�es en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("tblFAC_Comptes_Clients[#All]")
    ws.Range("M7").Value = rngData.Address
    
    'D�finir le range des crit�res
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("M2:N3")
    ws.Range("M3").Value = wshENC_Saisie.clientCode
    ws.Range("M8").Value = rngCriteria.Address
    
    'D�finir le range des r�sultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("P1").CurrentRegion
    rngResult.offset(2, 0).Clear
    Set rngResult = ws.Range("P2:U2")
    ws.Range("M9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
                                        
    'Est-ce que nous avons des r�sultats ?
    lastResultRow = ws.Cells(ws.Rows.count, "P").End(xlUp).row
    ws.Range("M10").Value = lastResultRow - 2 & " lignes"
    
    'Est-il n�cessaire de trier les r�sultats ?
    If lastResultRow > 3 Then
        With ws.Sort 'Sort - InvNo
            .SortFields.Clear
            'First sort On InvNo
            .SortFields.Add key:=ws.Range("Q3"), _
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
        ws.Range("U" & r).Value = ws.Range("S" & r).Value - ws.Range("T" & r).Value
    Next r

    'lib�rer la m�moire
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
        If .Range("F5").Value = Empty Or _
           .Range("K5").Value = Empty Or _
           .Range("F7").Value = Empty Or _
           .Range("K7").Value = 0 Then
            MsgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client valide" & vbNewLine & _
                "2. Une date d'encaissements" & vbNewLine & _
                "3. Un type de paiement et" & vbNewLine & _
                "4. Des montants appliqu�s" & vbNewLine & vbNewLine & _
                "AVANT de sauvegarder la transaction.", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Check to make sure Payment Amount = Applied Amount
        If .Range("K9").Value <> 0 Then
            MsgBox "Assurez-vous que le montant de l'encaissement soit �GAL" & vbNewLine & _
                "� la somme des paiements appliqu�s", vbExclamation
            GoTo Clean_Exit
        End If
        
        'Create records for ENC_Ent�te
        Call ENC_Add_DB_Entete
        Call ENC_Add_Locally_Entete
        
'        Dim pmtNo As Long
'        pmtNo = wshENC_Saisie.pmtNo
        
        Dim lastOSRow As Integer
        lastOSRow = .Cells(.Rows.count, "F").End(xlUp).row 'Last applied Item
        
        'Create records for ENC_D�tails
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
        Dim noEnc As String, nomCLient As String, typeEnc As String, descEnc As String
        Dim dateEnc As Date
        Dim montantEnc As Currency
        noEnc = wshENC_Saisie.pmtNo
        dateEnc = wshENC_Saisie.Range("K5").Value
        nomCLient = wshENC_Saisie.Range("F5").Value
        typeEnc = wshENC_Saisie.Range("F7").Value
        montantEnc = wshENC_Saisie.Range("K7").Value
        descEnc = wshENC_Saisie.Range("F9").Value

        Call ENC_GL_Posting_DB(noEnc, dateEnc, nomCLient, typeEnc, montantEnc, descEnc)  '2024-08-22 @ 16:08
        Call ENC_GL_Posting_Locally(noEnc, dateEnc, nomCLient, typeEnc, montantEnc, descEnc)  '2024-08-22 @ 16:08
        
        MsgBox "L'encaissement '" & wshENC_Saisie.pmtNo & "' a �t� enregistr� avec succ�s", vbOKOnly + vbInformation
        
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
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "ENC_Ent�te$"
    
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
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Locally_Entete", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentPmtNo As Long
    currentPmtNo = wshENC_Saisie.pmtNo
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_Ent�te.Cells(wshENC_Ent�te.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    wshENC_Ent�te.Range("A" & rowToBeUsed).Value = currentPmtNo
    wshENC_Ent�te.Range("B" & rowToBeUsed).Value = wshENC_Saisie.Range("K5").Value
    wshENC_Ent�te.Range("C" & rowToBeUsed).Value = wshENC_Saisie.Range("F5").Value
    wshENC_Ent�te.Range("D" & rowToBeUsed).Value = wshENC_Saisie.clientCode
    wshENC_Ent�te.Range("E" & rowToBeUsed).Value = wshENC_Saisie.Range("F7").Value
    wshENC_Ent�te.Range("F" & rowToBeUsed).Value = CDbl(Format$(wshENC_Saisie.Range("K7").Value, "#,##0.00"))
    wshENC_Ent�te.Range("G" & rowToBeUsed).Value = wshENC_Saisie.Range("F9").Value
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Entete", startTime)

End Sub

Sub ENC_Add_DB_Details(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "ENC_D�tails$"
    
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
    
    'What is the last used row in ENC_D�tails ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_D�tails.Cells(wshENC_D�tails.Rows.count, 1).End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).Value = True And _
            wshENC_Saisie.Range("K" & r).Value <> 0 Then
            wshENC_D�tails.Range("A" & rowToBeUsed).Value = pmtNo
            wshENC_D�tails.Range("B" & rowToBeUsed).Value = wshENC_Saisie.Range("F" & r).Value
            wshENC_D�tails.Range("C" & rowToBeUsed).Value = wshENC_Saisie.Range("F5").Value
            wshENC_D�tails.Range("D" & rowToBeUsed).Value = wshENC_Saisie.Range("K5").Value
            wshENC_D�tails.Range("E" & rowToBeUsed).Value = CDbl(Format$(wshENC_Saisie.Range("K" & r).Value, "#,##0.00"))
            rowToBeUsed = rowToBeUsed + 1
        End If
    Next r
    
    Application.ScreenUpdating = True

    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", startTime)

End Sub

Sub ENC_Update_DB_Comptes_Clients(firstRow As Integer, LastRow As Integer) 'Write to MASTER.xlsx
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_DB_Details", 0)
    
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
    For r = firstRow To LastRow
        If wshENC_Saisie.Range("B" & r).Value = True And _
            wshENC_Saisie.Range("K" & r).Value <> 0 Then
            'Open the recordset for the specified invoice
            Dim Inv_No As String
            Inv_No = CStr(Trim(wshENC_Saisie.Range("F" & r).Value))
            
            Dim strSQL As String
            strSQL = "SELECT * FROM [" & destinationTab & "] WHERE InvNo = '" & Inv_No & "'"
            rs.Open strSQL, conn, 2, 3
            If Not rs.EOF Then
                'Mettre � jour Amount_Paid
                rs.Fields(fFacCCTotalPaid - 1).Value = rs.Fields(fFacCCTotalPaid - 1).Value + CDbl(wshENC_Saisie.Range("K" & r).Value)
                'Mettre � jour Status
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
                'Mettre � jour le solde de la facture
                rs.Fields(fFacCCBalance - 1).Value = rs.Fields(fFacCCTotal - 1).Value - rs.Fields(fFacCCTotalPaid - 1).Value
                rs.update
            Else
                'Handle the case where the specified ID is not found
                MsgBox "L'enregistrement avec la facture '" & Inv_No & "' ne peut �tre retrouv�!", _
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

Sub ENC_Update_Locally_Comptes_Clients(firstRow As Integer, LastRow As Integer) '2024-08-22 @ 10:55
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim r As Integer
    For r = firstRow To LastRow
        Dim Inv_No As String
        Inv_No = CStr(wshENC_Saisie.Range("F" & r).Value)
        
        Dim foundRange As Range
        Set foundRange = lookupRange.Find(What:=Inv_No, LookIn:=xlValues, LookAt:=xlWhole)
    
        Dim rowToBeUpdated As Long
        If Not foundRange Is Nothing Then
            rowToBeUpdated = foundRange.row
            ws.Range("I" & rowToBeUpdated).Value = ws.Range("I" & rowToBeUpdated).Value + wshENC_Saisie.Range("K" & r).Value
            ws.Range("J" & rowToBeUpdated).Value = ws.Range("H" & rowToBeUpdated).Value - wshENC_Saisie.Range("K" & r).Value
            'Est-ce que le solde de la facture est � 0,00 $ ?
            If ws.Range("J" & rowToBeUpdated).Value = 0 Then
                ws.Range("E" & rowToBeUpdated) = "Paid"
            Else
                ws.Range("E" & rowToBeUpdated) = "Unpaid"
            End If
        Else
            MsgBox "La facture '" & Inv_No & "' n'existe pas dans FAC_Comptes_Clients.", vbCritical
        End If
    Next r
    
    Application.ScreenUpdating = True

    'Lib�rer la m�moire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modENC_Saisie:ENC_Add_Locally_Details", startTime)

End Sub

Sub ENC_GL_Posting_DB(no As String, dt As Date, nom As String, typeE As String, montant As Currency, desc As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_GL_Posting_DB", 0)
    
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
    strSQL = "SELECT MAX(NoEntr�e) AS MaxEJNo FROM [" & destinationTab & "]"

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
        rs.Fields(fGlTNoEntr�e - 1).Value = nextJENo
        rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
        If wshENC_Saisie.Range("F7").Value = "D�p�t de client" Then
            rs.Fields(fGlTDescription - 1).Value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            rs.Fields(fGlTSource - 1).Value = UCase(wshENC_Saisie.Range("F7").Value) & ":" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).Value = "2400" 'Hardcoded
            rs.Fields(fGlTCompte - 1).Value = "Produit per�u d'avance" 'Hardcoded
        Else
            rs.Fields(fGlTDescription - 1).Value = nom
            rs.Fields(fGlTSource - 1).Value = "ENCAISSEMENT:" & Format$(no, "00000")
            rs.Fields(fGlTNoCompte - 1).Value = "1000" 'Hardcoded
            rs.Fields(fGlTCompte - 1).Value = "Encaisse" 'Hardcoded
        End If
        rs.Fields(fGlTD�bit - 1).Value = montant
        rs.Fields(fGlTAutreRemarque - 1).Value = desc
        rs.Fields(fGlTTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    rs.update
    
    'Credit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields(fGlTNoEntr�e - 1).Value = nextJENo
        rs.Fields(fGlTDate - 1).Value = Format$(dt, "yyyy-mm-dd")
        If wshENC_Saisie.Range("F7").Value = "D�p�t de client" Then
            rs.Fields(fGlTDescription - 1).Value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            rs.Fields(fGlTSource - 1).Value = UCase(wshENC_Saisie.Range("F7").Value) & ":" & Format$(no, "00000")
        Else
            rs.Fields(fGlTDescription - 1).Value = nom
            rs.Fields(fGlTSource - 1).Value = "ENCAISSEMENT:" & Format$(no, "00000")
        End If
        rs.Fields(fGlTNoCompte - 1).Value = "1100" 'Hardcoded
        rs.Fields(fGlTCompte - 1).Value = "Comptes clients" 'Hardcoded
        rs.Fields(fGlTCr�dit - 1).Value = montant
        rs.Fields(fGlTAutreRemarque - 1).Value = desc
        rs.Fields(fGlTTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    rs.update

    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modENC_Saisie:ENC_GL_Posting_DB", startTime)

End Sub

Sub ENC_GL_Posting_Locally(no As String, dt As Date, nom As String, typeE As String, montant As Currency, desc As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_GL_Posting_Locally", 0)
    
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
        If wshENC_Saisie.Range("F7").Value = "D�p�t de client" Then
            .Range("C" & rowToBeUsed).Value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            .Range("D" & rowToBeUsed).Value = UCase(wshENC_Saisie.Range("F7").Value) & ":" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = "2400" 'Hardcoded
            .Range("F" & rowToBeUsed).Value = "Produit per�u d'avance" 'Hardcoded
        Else
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "ENCAISSEMENT:" & Format$(no, "00000")
            .Range("E" & rowToBeUsed).Value = "1000" 'Hardcoded
            .Range("F" & rowToBeUsed).Value = "Encaisse" 'Hardcoded
        End If
        .Range("G" & rowToBeUsed).Value = montant
        .Range("I" & rowToBeUsed).Value = desc
        .Range("J" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
    
    'Credit side
        .Range("A" & rowToBeUsed).Value = nextJENo
        .Range("B" & rowToBeUsed).Value = CDate(dt)
        If wshENC_Saisie.Range("F7").Value = "D�p�t de client" Then
            .Range("C" & rowToBeUsed).Value = "Client:" & wshENC_Saisie.clientCode & " - " & nom
            .Range("D" & rowToBeUsed).Value = UCase(wshENC_Saisie.Range("F7").Value) & ":" & Format$(no, "00000")
        Else
            .Range("C" & rowToBeUsed).Value = nom
            .Range("D" & rowToBeUsed).Value = "ENCAISSEMENT:" & Format$(no, "00000")
        End If
        .Range("E" & rowToBeUsed).Value = "1100" 'Hardcoded
        .Range("F" & rowToBeUsed).Value = "Comptes clients" 'Hardcoded
        .Range("H" & rowToBeUsed).Value = montant
        .Range("I" & rowToBeUsed).Value = desc
        .Range("J" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    End With
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modENC_Saisie:ENC_GL_Posting_Locally", startTime)

End Sub

Sub ENC_Add_Check_Boxes(row As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Add_Check_Boxes", 0)
    
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
        .EnableSelection = xlNoRestrictions
    End With
    
    Application.EnableEvents = True

    'Lib�rer la m�moire
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
    
    'Lib�rer la m�moire
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
        
        .Range("K5").Value = ""
        .Range("F7").Value = "Banque" ' Set Default type
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

    'Lib�rer la m�moire
    Set chkBox = Nothing
    Set linkedCell = Nothing
    
End Sub

Sub shp_ENC_Exit_Click()

    Call ENC_Back_To_GL_Menu

End Sub

Sub ENC_Back_To_GL_Menu()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modENC_Saisie:ENC_Back_To_GL_Menu", 0)
   
    If wshENC_Saisie.ProtectContents Then
        wshENC_Saisie.Unprotect
    End If
    
    Application.EnableEvents = False
    
    Call ENC_Clear_Cells
    
    Application.EnableEvents = True
    
    wshENC_Saisie.Visible = xlSheetVeryHidden

    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call Log_Record("modENC_Saisie:ENC_Back_To_GL_Menu", startTime)

End Sub

