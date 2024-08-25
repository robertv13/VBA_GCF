Attribute VB_Name = "modENC_Saisie"
Option Explicit

Dim lastRow As Long, lastResultRow As Long
Dim payRow As Long

Sub ENC_Load_OS_Invoices(clientCode As String) '2024-08-21 @ 15:18
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Load_OS_Invoices()")
    
    wshENC_Saisie.Range("E15:K42").ClearContents 'Clear the invoices area before loading it
    
    With wshFAC_Comptes_Clients
        'Clear previous results
        lastResultRow = .Cells(.rows.count, "P").End(xlUp).Row
        If lastResultRow > 2 Then
            .Range("P3:U" & lastResultRow).ClearContents
        End If
        'Is there anything to work with ?
        lastResultRow = .Cells(.rows.count, "A").End(xlUp).Row
        If lastResultRow < 3 Then Exit Sub
        
        'Setup criteria in wshFAC_Comptes_Clients
        .Range("M3").value = clientCode
        
        .Range("A2:K" & lastResultRow).AdvancedFilter _
                                            xlFilterCopy, _
                                            criteriaRange:=.Range("M2:N3"), _
                                            CopyToRange:=.Range("P2:U2")
                                            
        'Did the AdvancedFilter return ANYTHING ?
        lastResultRow = .Cells(.rows.count, "P").End(xlUp).Row
        If lastResultRow < 3 Then Exit Sub
        
        'PLUG - Recalculate Column 'U' - Balance after AdvancedFilter
        Dim r As Integer
        For r = 3 To lastResultRow
            .Range("U" & r).value = .Range("S" & r).value - .Range("T" & r).value
        Next r
        
        wshENC_Saisie.Range("B4").value = True 'Set PaymentLoad to True
'        .Range("T3:T" & lastResultRow).formula = .Range("T1").formula 'Total Payments Formula

        'Bring the Result data into our List of Oustanding Invoices
        Dim i As Integer
        'Unlock the required area
        With wshENC_Saisie '2024-08-21 @ 16:06
            .Unprotect
            .Range("B12:B" & 11 + lastResultRow - 2).Locked = False
            .Range("E12:J" & 11 + lastResultRow - 2).Locked = False
            .Protect UserInterfaceOnly:=True
            .EnableSelection = xlUnlockedCells
        End With
        
        Dim rr As Integer: rr = 12
        For i = 3 To WorksheetFunction.Min(27, lastResultRow)
            If .Range("U" & i).value <> 0 Then
                wshENC_Saisie.Range("F" & rr).value = .Range("Q" & i).value
                wshENC_Saisie.Range("G" & rr).value = .Range("R" & i).value
                wshENC_Saisie.Range("H" & rr).value = .Range("S" & i).value
                wshENC_Saisie.Range("I" & rr).value = .Range("T" & i).value
                wshENC_Saisie.Range("J" & rr).value = .Range("U" & i).value
                rr = rr + 1
            End If
        Next i
        
        Call ENC_Saisie_Add_Check_Boxes(lastResultRow - 2)
        
'        wshENC_Saisie.Range("E13:I" & lastResultRow + 10).value = .Range("O3:S" & lastResultRow).value

    End With
    
    wshENC_Saisie.Range("B4").value = False 'Set PaymentLoad to False
    
'    Call End_Timer("modFAC_Enc:Encaissement_Load_Open_Invoices()", timerStart)

End Sub

Sub ENC_Update() '2024-08-22 @ 09:46
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Update()")
    
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
        
        Dim pmtNo As Long
        pmtNo = wshENC_Saisie.Range("B9").value
        
        Dim lastOSRow As Integer
        lastOSRow = .Cells(.rows.count, "F").End(xlUp).Row 'Last applied Item
        
        'Create records for ENC_Détails
        If lastOSRow > 11 Then
            Call ENC_Add_DB_Details(pmtNo, 12, lastOSRow)
            Call ENC_Add_Locally_Details(pmtNo, 12, lastOSRow)
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
        noEnc = wshENC_Saisie.Range("B9").value
        dateEnc = wshENC_Saisie.Range("K5").value
        nomClient = wshENC_Saisie.Range("F5").value
        typeEnc = wshENC_Saisie.Range("F7").value
        montantEnc = wshENC_Saisie.Range("K7").value
        descEnc = wshENC_Saisie.Range("F9").value

        Call ENC_GL_Posting_DB(noEnc, dateEnc, nomClient, typeEnc, montantEnc, descEnc)  '2024-08-22 @ 16:08
        Call ENC_GL_Posting_Locally(noEnc, dateEnc, nomClient, typeEnc, montantEnc, descEnc)  '2024-08-22 @ 16:08
        
        MsgBox "L'encaissement '" & pmtNo & "' a été renregistré avec succès", vbInformation
        
        Call Encaissement_Add_New 'Reset the form
        
        .Range("F5").Select
    End With
    
Clean_Exit:

'    Call End_Timer("modENC_Saisie:ENC_Update()", timerStart)

End Sub

Sub Encaissement_Add_New() '2024-08-21 @ 14:58

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modEnc_Saisie:Encaissement_Add_New()")

    Call ENC_Clear_Cells
    
    Call End_Timer("modEnc_Saisie:Encaissement_Add_New()", timerStart)
    
End Sub

'Sub Encaissement_Previous() '2024-02-14 @ 11:04
'
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Encaissement_Previous()")
'
'    Dim MinPayID As Long, PayID As Long
'    With wshENC_Saisie
'        On Error Resume Next
'            MinPayID = Application.WorksheetFunction.Min(wshENC_Entête.Range("Pay_ID"))
'        On Error GoTo 0
'        If MinPayID = 0 Then
'            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
'            Exit Sub
'        End If
'        PayID = .Range("B3").value 'Payment ID
'        If PayID = 0 Or .Range("B4").value = Empty Then 'Load Last Payment Created
'            payRow = wshENC_Entête.Range("A99999").End(xlUp).Row 'Last Row
'        Else
'            payRow = .Range("B4").value - 1 'Pay Row
'        End If
'        If payRow = 3 Or MinPayID = .Range("B3").value Then 'First Payment
'            MsgBox "Vous êtes au premier paiement", vbExclamation
'            Exit Sub
'        End If
'        .Range("B3").value = wshENC_Entête.Range("A" & payRow).value 'Set Payment ID
'        Call Encaissement_Load 'Load Payment
'    End With
'
'    Call End_Timer("modFAC_Enc:Encaissement_Previous()", timerStart)
'
'End Sub
'
'Sub Encaissement_Next() '2024-02-14 @ 11:04
'
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Encaissement_Next()")
'
'    Application.EnableEvents = False
'
'    Dim MaxPayID As Long
'    With wshENC_Saisie
'        On Error Resume Next
'            MaxPayID = Application.WorksheetFunction.Max(wshENC_Entête.Range("Pay_ID"))
'        On Error GoTo 0
'        If MaxPayID = 0 Then
'            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
'            Exit Sub
'        End If
'        Dim PayID As Long
'        PayID = .Range("B3").value 'Payment ID
'        If PayID = 0 Or .Range("B4").value = Empty Then 'Load Last Payment Created
'            payRow = 4 'On new Payment, GOTO first one created
'        Else
'            payRow = .Range("B4").value + 1 'Pay Row
'        End If
'        If MaxPayID = PayID Then 'Last Payment
'            MsgBox "Vous êtes au dernier paiement", vbExclamation
'            Exit Sub
'        End If
'        .Range("B3").value = wshENC_Entête.Range("A" & payRow).value 'Set PayID
'        Call Encaissement_Load 'Load Payment for the PayID
'    End With
'
'    Application.EnableEvents = True
'
'    Call End_Timer("modFAC_Enc:Encaissement_Next()", timerStart)
'
'End Sub
'
'Sub Encaissement_Load() '2024-02-14 @ 11:04
'
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Encaissement_Load()")
'
'    With wshENC_Saisie
'        If .Range("B4").value = Empty Then
'            MsgBox "Assurez vous de choisir un paiement valide", vbExclamation
'            Exit Sub
'        End If
'        payRow = .Range("B4").value 'Payment Row
'        .Range("B4").value = True
'        .Range("F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").ClearContents
'        'Update worksheet fields
'        .Range("J3").value = wshENC_Entête.Cells(payRow, 2).value
'        .Range("F3").value = wshENC_Entête.Cells(payRow, 3).value
'        .Range("F5").value = wshENC_Entête.Cells(payRow, 4).value
'        .Range("J5").value = wshENC_Entête.Cells(payRow, 5).value
'        .Range("F7").value = wshENC_Entête.Cells(payRow, 6).value
'
'        'Load Pay Items
'        With wshENC_Détails
'            .Range("M4:T999999").ClearContents
'            lastRow = .Range("A999999").End(xlUp).Row
'            If lastRow < 4 Then GoTo NoData
'            .Range("A3:G" & lastRow).AdvancedFilter _
'                xlFilterCopy, _
'                criteriaRange:=.Range("J2:J3"), _
'                CopyToRange:=.Range("O3:T3"), _
'                Unique:=True
'            lastResultRow = .Range("O99999").End(xlUp).Row
'            If lastResultRow < 4 Then GoTo NoData
'            'Bring down the formulas into results
'            .Range("M4:N" & lastResultRow).formula = .Range("M1:N1").formula 'Bring Apply and Invoice Date Formulas
'            .Range("P4:R" & lastResultRow).formula = .Range("P1:R1").formula 'Inv. Amount, Prev. payments & Balance formulas
'            wshENC_Saisie.Range("D13:K" & lastResultRow + 9).value = .Range("M4:T" & lastResultRow).value 'Bring over Pay Items
'NoData:
'        End With
'        .Range("B4").value = False 'Payment Load to False
'    End With
'
'    Call End_Timer("modFAC_Enc:Encaissement_Load()", timerStart)
'
'End Sub
'
Sub ENC_Add_DB_Entete() 'Write to MASTER.xlsx
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Add_DB_Entete()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "ENC_Entête"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxPmtNo As Long
    strSQL = "SELECT MAX(Pay_ID) AS MaxPmtNo FROM [" & destinationTab & "$]"

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
    Dim pmtNo As Long
    pmtNo = lr + 1

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
        rs.Fields("Pay_ID").value = pmtNo
        rs.Fields("Pay_Date").value = wshENC_Saisie.Range("K5").value
        rs.Fields("Customer").value = wshENC_Saisie.Range("F5").value
        rs.Fields("codeClient").value = wshENC_Saisie.Range("B8").value
        rs.Fields("Pay_Type").value = wshENC_Saisie.Range("F7").value
        rs.Fields("Amount").value = CDbl(Format$(wshENC_Saisie.Range("K7").value, "#,##0.00 $"))
        rs.Fields("Notes").value = wshENC_Saisie.Range("F9").value
    'Update the recordset (create the record)
    rs.update
    
    Application.EnableEvents = False
    wshENC_Saisie.Range("B9").value = pmtNo
    Application.EnableEvents = True
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

'    Call End_Timer("modENC_Saisie:ENC_Add_DB_Entete()", timerStart)
    
End Sub

Sub ENC_Add_Locally_Entete() '2024-08-22 @ 10:38
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modDEB_Saisie:DEB_Trans_Add_Record_Locally()")
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentPmnNo As Long
    currentPmnNo = CLng(wshENC_Saisie.Range("B9").value)
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_Entête.Range("A99999").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    wshENC_Entête.Range("A" & rowToBeUsed).value = currentPmnNo
    wshENC_Entête.Range("B" & rowToBeUsed).value = CDate(wshENC_Saisie.Range("K5").value)
    wshENC_Entête.Range("C" & rowToBeUsed).value = wshENC_Saisie.Range("F5").value
    wshENC_Entête.Range("D" & rowToBeUsed).value = wshENC_Saisie.Range("B8").value
    wshENC_Entête.Range("E" & rowToBeUsed).value = wshENC_Saisie.Range("F7").value
    wshENC_Entête.Range("F" & rowToBeUsed).value = CDbl(Format$(wshENC_Saisie.Range("K7").value, "#,##0.00"))
    wshENC_Entête.Range("G" & rowToBeUsed).value = wshENC_Saisie.Range("F9").value
    
    Application.ScreenUpdating = True

'    Call End_Timer("modDEB_Saisie:DEB_Trans_Add_Record_Locally()", timerStart)

End Sub

Sub ENC_Add_DB_Details(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) 'Write to MASTER.xlsx
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Add_DB_Details()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "ENC_Détails"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
        
    'Build the recordSet
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).value = True And _
            wshENC_Saisie.Range("K" & r).value <> 0 Then
            rs.AddNew
                rs.Fields("Pay_ID").value = pmtNo
                rs.Fields("Inv_No").value = wshENC_Saisie.Range("F" & r).value
                rs.Fields("Customer").value = wshENC_Saisie.Range("F5").value
                rs.Fields("Pay_Date").value = CDate(wshENC_Saisie.Range("K5").value)
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

'    Call End_Timer("modENC_Saisie:ENC_Add_DB_Details()", timerStart)
    
End Sub

Sub ENC_Add_Locally_Details(pmtNo As Long, firstRow As Integer, lastAppliedRow As Integer) '2024-08-22 @ 10:55
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Add_Locally_Details()")
    
    Application.ScreenUpdating = False
    
    'What is the last used row in ENC_Détails ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshENC_Détails.Cells(wshENC_Détails.rows.count, "A").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    Dim r As Integer
    For r = firstRow To lastAppliedRow
        If wshENC_Saisie.Range("B" & r).value = True And _
            wshENC_Saisie.Range("K" & r).value <> 0 Then
            wshENC_Détails.Range("A" & rowToBeUsed).value = pmtNo
            wshENC_Détails.Range("B" & rowToBeUsed).value = wshENC_Saisie.Range("F" & r).value
            wshENC_Détails.Range("C" & rowToBeUsed).value = wshENC_Saisie.Range("F5").value
            wshENC_Détails.Range("D" & rowToBeUsed).value = CDate(wshENC_Saisie.Range("K5").value)
            wshENC_Détails.Range("E" & rowToBeUsed).value = CDbl(Format$(wshENC_Saisie.Range("K" & r).value, "#,##0.00"))
            rowToBeUsed = rowToBeUsed + 1
        End If
    Next r
    
    Application.ScreenUpdating = True

'    Call End_Timer("modENC_Saisie:ENC_Add_Locally_Details()", timerStart)

End Sub

Sub ENC_Update_DB_Comptes_Clients(firstRow As Integer, lastRow As Integer) 'Write to MASTER.xlsx
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Add_DB_Details()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Comptes_Clients"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim r As Long
    For r = firstRow To lastRow
        If wshENC_Saisie.Range("B" & r).value = True And _
            wshENC_Saisie.Range("K" & r) <> 0 Then
            'Open the recordset for the specified invoice
            Dim Inv_No As String
            Inv_No = CStr(Trim(wshENC_Saisie.Range("F" & r).value))
            Dim strSQL As String
            strSQL = "SELECT * FROM [" & destinationTab & "$] WHERE Invoice_No = '" & Inv_No & "'"
            rs.Open strSQL, conn, 2, 3
            If Not rs.EOF Then
                'Update Amount_Paid
                rs.Fields("Total_Paid").value = rs.Fields("Total_Paid").value + wshENC_Saisie.Range("K" & r).value
                'Update invoice Status
                If rs.Fields("Total").value - rs.Fields("Total_Paid").value = 0 Then
                    rs.Fields("Status").value = "Paid"
                End If
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

'    Call End_Timer("modENC_Saisie:ENC_Add_DB_Details()", timerStart)
    
End Sub

Sub ENC_Update_Locally_Comptes_Clients(firstRow As Integer, lastRow As Integer) '2024-08-22 @ 10:55
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Add_Locally_Details()")
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim r As Integer
    For r = firstRow To lastRow
        Dim Inv_No As String
        Inv_No = CStr(wshENC_Saisie.Range("F" & r).value)
        
        Dim foundRange As Range
        Set foundRange = lookupRange.Find(What:=Inv_No, LookIn:=xlValues, lookAt:=xlWhole)
    
        Dim rowToBeUpdated As Long
        If Not foundRange Is Nothing Then
            rowToBeUpdated = foundRange.Row
            ws.Cells(rowToBeUpdated, 9).value = ws.Cells(rowToBeUpdated, 9).value + wshENC_Saisie.Range("K" & rowToBeUpdated).value
            ws.Cells(rowToBeUpdated, 10).value = ws.Cells(rowToBeUpdated, 10).value - wshENC_Saisie.Range("K" & rowToBeUpdated).value
            If ws.Cells(rowToBeUpdated, 10).value = 0 Then
                ws.Cells(rowToBeUpdated, 5) = "Paid"
            End If
        Else
            MsgBox "La facture '" & Inv_No & "' n'existe pas dans FAC_Comptes_Clients.", vbCritical
        End If
    Next r
    
    Application.ScreenUpdating = True

'    Call End_Timer("modENC_Saisie:ENC_Add_Locally_Details()", timerStart)

End Sub

Sub ENC_GL_Posting_DB(no As String, dt As Date, nom As String, typeE As String, montant As Currency, desc As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_GL_Posting_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans"
    
    'Initialize connection, connection string, open the connection & declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJNo As Long
    strSQL = "SELECT MAX(No_Entrée) AS MaxEJNo FROM [" & destinationTab & "$]"

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
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    'Debit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("No_Entrée").value = nextJENo
        rs.Fields("Date").value = CDate(dt)
        rs.Fields("Description").value = nom
        rs.Fields("Source").value = "ENCAISSEMENT:" & Format$(no, "00000")
        rs.Fields("No_Compte").value = "1000" 'Hardcoded
        rs.Fields("Compte").value = "Encaisse" 'Hardcoded
        rs.Fields("Débit").value = montant
        rs.Fields("AutreRemarque").value = desc
        rs.Fields("TimeStamp").value = Format$(Now(), "dd/mm/yyyy hh:nn:ss")
    rs.update
    
    'Credit side
    rs.AddNew
        'Add fields to the recordset before updating it
        rs.Fields("No_Entrée").value = nextJENo
        rs.Fields("Date").value = CDate(dt)
        rs.Fields("Description").value = nom
        rs.Fields("Source").value = "ENCAISSEMENT:" & Format$(no, "00000")
        rs.Fields("No_Compte").value = "1100" 'Hardcoded
        rs.Fields("Compte").value = "Comptes clients" 'Hardcoded
        rs.Fields("Crédit").value = montant
        rs.Fields("AutreRemarque").value = desc
        rs.Fields("TimeStamp").value = Format$(Now(), "dd/mm/yyyy hh:nn:ss")
    rs.update

    wshENC_Saisie.Range("B10").value = nextJENo
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call End_Timer("modENC_Saisie:ENC_GL_Posting_DB()", timerStart)

End Sub

Sub ENC_GL_Posting_Locally(no As String, dt As Date, nom As String, typeE As String, montant As Currency, desc As String) 'Write/Update to GCF_BD_MASTER / GL_Trans
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_GL_Posting_Locally()")
    
    Application.ScreenUpdating = False
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Cells(wshGL_Trans.rows.count, "A").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    Dim nextJENo As Long
    nextJENo = wshENC_Saisie.Range("B10").value
    
    With wshGL_Trans
    'Debit side
        .Range("A" & rowToBeUsed).value = nextJENo
        .Range("B" & rowToBeUsed).value = CDate(dt)
        .Range("C" & rowToBeUsed).value = nom
        .Range("D" & rowToBeUsed).value = "ENCAISSEMENT:" & Format$(no, "00000")
        .Range("E" & rowToBeUsed).value = "1000" 'Hardcoded
        .Range("F" & rowToBeUsed).value = "Encaisse" 'Hardcoded
        .Range("G" & rowToBeUsed).value = montant
        .Range("I" & rowToBeUsed).value = desc
        .Range("J" & rowToBeUsed).value = Format$(Now(), "dd/mm/yyyy hh:nn:ss")
        rowToBeUsed = rowToBeUsed + 1
    
    'Credit side
        .Range("A" & rowToBeUsed).value = nextJENo
        .Range("B" & rowToBeUsed).value = CDate(dt)
        .Range("C" & rowToBeUsed).value = nom
        .Range("D" & rowToBeUsed).value = "ENCAISSEMENT:" & Format$(no, "00000")
        .Range("E" & rowToBeUsed).value = "1100" 'Hardcoded
        .Range("F" & rowToBeUsed).value = "Comptes clients" 'Hardcoded
        .Range("H" & rowToBeUsed).value = montant
        .Range("I" & rowToBeUsed).value = desc
        .Range("J" & rowToBeUsed).value = Format$(Now(), "dd/mm/yyyy hh:nn:ss")
    End With
    
    Application.ScreenUpdating = True
    
    Call End_Timer("modENC_Saisie:ENC_GL_Posting_Locally()", timerStart)

End Sub

Sub ENC_Saisie_Add_Check_Boxes(Row As Long)

'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Saisie_Add_Check_Boxes()")
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshENC_Saisie
    
    Dim chkBoxRange As Range: Set chkBoxRange = ws.Range("E12:E" & 11 + Row)
    
    Dim cell As Range
    Dim cbx As checkBox
    For Each cell In chkBoxRange
    'Check if the cell is empty and doesn't have a checkbox already
    If cell.Row <= 36 And _
        Cells(cell.Row, 2).value = "" And _
        Cells(cell.Row, 6).value <> "" Then 'Applied = False
            'Create a checkbox linked to the cell
            Set cbx = wshENC_Saisie.CheckBoxes.add(cell.Left + 30, cell.Top, cell.width, cell.Height)
            With cbx
                .name = "chkBox - " & cell.Row
                .Caption = ""
                .value = False
                .linkedCell = "B" & cell.Row
                .Display3DShading = True
                .OnAction = "chkBox_Apply_Click"
                .Locked = False
            End With
    End If
    Next cell

'    'Protect the worksheet
'    ws.Protect UserInterfaceOnly:=True
    
    Application.EnableEvents = True

    'Cleaning memory - 2024-08-21 @ 16:42
    Set cbx = Nothing
    Set cell = Nothing
    Set chkBoxRange = Nothing
    Set ws = Nothing
    
'    Call End_Timer("modENC_Saisie:ENC_Saisie_Add_Check_Boxes()", timerStart)

End Sub

Sub ENC_Remove_Check_Boxes(Row As Long)

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Remove_Check_Boxes()")
    
    Application.EnableEvents = False
    
    'Delete all checkboxes whose name are chkBox - ...
    Dim cbx As Shape
    For Each cbx In wshENC_Saisie.Shapes
        If InStr(cbx.name, "chkBox -") Then
            cbx.delete
        End If
    Next cbx
    
'    wshFAC_Brouillon.Range("B12:B" & row).value = ""
    
    Application.EnableEvents = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cbx = Nothing
    
    Call End_Timer("modENC_Saisie:ENC_Remove_Check_Boxes()", timerStart)

End Sub

Sub ENC_Clear_Cells()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Clear_Cells()")
    
    wshENC_Saisie.Unprotect
    
    With wshENC_Saisie
    
        Application.EnableEvents = False
        
        'Note the lastUsedRow for checkBox deletion
        Dim lastUsedRow As Long
        lastUsedRow = wshENC_Saisie.Cells(wshENC_Saisie.rows.count, "F").End(xlUp).Row
        
        .Range("B4").value = False
        .Range("B5,F5:H5,K5,F7,K7,F9:I9,E12:K36").ClearContents 'Clear Fields
        .Range("B12:B36").ClearContents
        
        If lastUsedRow > 11 Then
            Call ENC_Remove_Check_Boxes(lastUsedRow)
        End If
        
        .Range("K5").value = ""
        .Range("F7").value = "Banque" ' Set Default type
        .Range("F5").Select
    End With
    
    With wshENC_Saisie.Range("F5:H5, K5, F7, K7, F9:I9").Interior '2024-08-25 @ 09:21
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    
    wshENC_Saisie.Shapes("btnENC_Sauvegarde").Visible = False
    wshENC_Saisie.Shapes("btnENC_Annule").Visible = False
    
    Application.EnableEvents = True
    
    wshENC_Saisie.Protect UserInterfaceOnly:=True
    wshENC_Saisie.EnableSelection = xlUnlockedCells

    Call End_Timer("modENC_Saisie:ENC_Clear_Cells()", timerStart)

End Sub

Sub chkBox_Apply_Click()

    Dim chkBox As checkBox
    Set chkBox = ActiveSheet.CheckBoxes(Application.Caller)
    Dim linkedCell As Range
    Set linkedCell = Range(chkBox.linkedCell)
    
    If linkedCell.value = True Then
        If wshENC_Saisie.Range("K9").value > 0 Then
            If wshENC_Saisie.Range("K9").value > wshENC_Saisie.Range("J" & linkedCell.Row).value Then
                wshENC_Saisie.Range("K" & linkedCell.Row).value = wshENC_Saisie.Range("J" & linkedCell.Row).value
            Else
                wshENC_Saisie.Range("K" & linkedCell.Row).value = wshENC_Saisie.Range("K9").value
            End If
        End If
        wshENC_Saisie.Shapes("btnENC_Sauvegarde").Visible = True
        wshENC_Saisie.Shapes("btnENC_Annule").Visible = True
    Else
        Range("K" & linkedCell.Row).value = 0
    End If

    'Clean up - 2024-08-21 @ 20:16
    Set chkBox = Nothing
    Set linkedCell = Nothing
    
End Sub

Sub ENC_Back_To_FAC_Menu()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Back_To_FAC_Menu()")
   
    wshENC_Saisie.Unprotect
    
    Application.EnableEvents = False
    
    Call ENC_Clear_Cells
    
    Application.EnableEvents = True
    
    wshENC_Saisie.Visible = xlSheetHidden

    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call End_Timer("modENC_Saisie:ENC_Back_To_FAC_Menu()", timerStart)

End Sub


