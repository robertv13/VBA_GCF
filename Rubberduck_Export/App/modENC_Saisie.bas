Attribute VB_Name = "modENC_Saisie"
Option Explicit

Dim LastRow As Long, lastResultRow As Long
Dim payRow As Long
Dim resultRow As Long, payItemRow As Long, lastPayItemRow As Long, payitemDBRow As Long

Sub ENC_Load_OS_Invoices(clientCode As String) '2024-08-21 @ 15:18
    
'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Load_OS_Invoices()")
    
    wshENC_Saisie.Range("E15:K42").ClearContents 'Clear the invoices area before loading it
    
    With wshFAC_Comptes_Clients
        'Clear previous results
        lastResultRow = .Cells(.rows.count, "P").End(xlUp).row
        If lastResultRow > 2 Then
            .Range("P3:U" & lastResultRow).ClearContents
        End If
        'Is there anything to work with ?
        lastResultRow = .Cells(.rows.count, "A").End(xlUp).row
        If lastResultRow < 3 Then Exit Sub
        
        'Setup criteria in wshFAC_Comptes_Clients
        .Range("M3").value = clientCode
        
        .Range("A2:K" & lastResultRow).AdvancedFilter _
                                            xlFilterCopy, _
                                            criteriaRange:=.Range("M2:N3"), _
                                            CopyToRange:=.Range("P2:U2")
                                            
        'Did the AdvancedFilter return ANYTHING ?
        lastResultRow = .Cells(.rows.count, "P").End(xlUp).row
        If lastResultRow < 3 Then Exit Sub
        
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
        
        For i = 3 To lastResultRow
            wshENC_Saisie.Range("F" & i + 9).value = .Range("Q" & i).value
            wshENC_Saisie.Range("G" & i + 9).value = .Range("R" & i).value
            wshENC_Saisie.Range("H" & i + 9).value = .Range("S" & i).value
            wshENC_Saisie.Range("I" & i + 9).value = .Range("T" & i).value
            wshENC_Saisie.Range("J" & i + 9).value = .Range("U" & i).value
        Next i
        
        Call ENC_Saisie_Add_Check_Boxes(lastResultRow - 2)
        
'        wshENC_Saisie.Range("E13:I" & lastResultRow + 10).value = .Range("O3:S" & lastResultRow).value

    End With
    
    wshENC_Saisie.Range("B4").value = False 'Set PaymentLoad to False
    
'    Call End_Timer("modFAC_Enc:Encaissement_Load_Open_Invoices()", timerStart)

End Sub

Sub Encaissement_Save_Update() '2024-02-07 @ 12:27
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Encaissement_Save_Update()")
    
    With wshENC_Saisie
        'Check for mandatory fields (4)
        If .Range("F3").value = Empty Or _
           .Range("J3").value = Empty Or _
           .Range("J3").value = Empty Then
            MsgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client" & vbNewLine & _
                "2. Une date de paiement" & vbNewLine & _
                "3. Un type de paiement et" & vbNewLine & _
                "4. Des transactions" & vbNewLine & vbNewLine & _
                "AVANT de sauvegarder la transaction.", vbExclamation
            Exit Sub
        End If
        'Check to make sure Payment Amount = Applied Amount
        If .Range("J5").value <> .Range("J10").value Then
            MsgBox "Assurez-vous que le montant du paiement soit ÉGAL" & vbNewLine & _
                "à la somme des paiements appliqués", vbExclamation
            Exit Sub
        End If
        'New Payment -OR- Existing Payment ?
        If .Range("B4").value = Empty Then 'New Payment
            payRow = wshENC_Entête.Range("A999999").End(xlUp).row + 1 'First Available Row
            .Range("B3").value = .Range("B5").value 'Next payment ID
            'wshENC_Entête.Range("A" & payRow).value = .Range("B3").value 'PayID
            Call Add_Or_Update_Enc_Entete_Record_To_DB(0)
        Else 'Existing Payment
            Call Add_Or_Update_Enc_Entete_Record_To_DB(.Range("B4").value)
        End If
        
        'Save Applied Invoices to Payment Detail
        lastPayItemRow = .Range("E999999").End(xlUp).row 'Last Pay Item
        
        For payItemRow = 13 To lastPayItemRow
            If .Range("D" & payItemRow).value = Chr(252) Then 'The row has been applied
                If .Range("K" & payItemRow).value = Empty Then 'New Pay Item row
                    payitemDBRow = wshENC_Détails.Range("A999999").End(xlUp).row + 1 'First Avail Pay Items Row
                    Call Add_Or_Update_Enc_Detail_Record_To_DB(0, payItemRow)
                    'wshENC_Détails.Range("A" & payitemDBRow).value = .Range("B3").value 'Payment ID
                    'wshENC_Détails.Range("F" & payitemDBRow).value = "=row()"
                    .Range("K" & payItemRow).value = payitemDBRow 'Database Row
                Else 'Existing Pay Item
                    payitemDBRow = .Range("K" & payItemRow).value 'Existing Pay Item Row
                End If
'                wshENC_Détails.Range("B" & payitemDBRow).value = .Range("F" & payItemRow).value 'Invoice ID
'                wshENC_Détails.Range("C" & payitemDBRow).value = .Range("F3").value 'Customer
'                wshENC_Détails.Range("D" & payitemDBRow).value = .Range("J3").value 'Pay Date
'                wshENC_Détails.Range("E" & payitemDBRow).value = .Range("J" & payItemRow).value 'Amount paid
            End If
        Next payItemRow
        
        'Prepare G/L posting
        Dim noEnc As String, nomClient As String, typeEnc As String, descEnc As String
        Dim dateEnc As Date
        Dim montantEnc As Currency
        noEnc = wshENC_Saisie.Range("B5").value
        dateEnc = wshENC_Saisie.Range("J3").value
        nomClient = wshENC_Saisie.Range("F3").value
        typeEnc = wshENC_Saisie.Range("F5").value
        montantEnc = wshENC_Saisie.Range("J5").value
        descEnc = wshENC_Saisie.Range("F7").value

        Call Encaissement_GL_Posting(noEnc, dateEnc, nomClient, typeEnc, montantEnc, descEnc)  '2024-02-09 @ 08:17 - TODO
        
        Call Encaissement_Import_All   'Bring back locally three worksheets
        
        MsgBox "Le paiement a été renregistré avec succès"
        Call Encaissement_Add_New 'Reset the form
        .Range("F3").Select
    End With
    
    Call End_Timer("modFAC_Enc:Encaissement_Save_Update()", timerStart)

End Sub

Sub Encaissement_Add_New() '2024-08-21 @ 14:58

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modEnc_Saisie:Encaissement_Add_New()")

    Call ENC_Clear_Cells
    
    Call End_Timer("modEnc_Saisie:Encaissement_Add_New()", timerStart)
    
End Sub

Sub Encaissement_Previous() '2024-02-14 @ 11:04
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Encaissement_Previous()")

    Dim MinPayID As Long, PayID As Long
    With wshENC_Saisie
        On Error Resume Next
            MinPayID = Application.WorksheetFunction.Min(wshENC_Entête.Range("Pay_ID"))
        On Error GoTo 0
        If MinPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
            Exit Sub
        End If
        PayID = .Range("B3").value 'Payment ID
        If PayID = 0 Or .Range("B4").value = Empty Then 'Load Last Payment Created
            payRow = wshENC_Entête.Range("A99999").End(xlUp).row 'Last Row
        Else
            payRow = .Range("B4").value - 1 'Pay Row
        End If
        If payRow = 3 Or MinPayID = .Range("B3").value Then 'First Payment
            MsgBox "Vous êtes au premier paiement", vbExclamation
            Exit Sub
        End If
        .Range("B3").value = wshENC_Entête.Range("A" & payRow).value 'Set Payment ID
        Call Encaissement_Load 'Load Payment
    End With

    Call End_Timer("modFAC_Enc:Encaissement_Previous()", timerStart)

End Sub

Sub Encaissement_Next() '2024-02-14 @ 11:04
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Encaissement_Next()")

    Application.EnableEvents = False

    Dim MaxPayID As Long
    With wshENC_Saisie
        On Error Resume Next
            MaxPayID = Application.WorksheetFunction.Max(wshENC_Entête.Range("Pay_ID"))
        On Error GoTo 0
        If MaxPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
            Exit Sub
        End If
        Dim PayID As Long
        PayID = .Range("B3").value 'Payment ID
        If PayID = 0 Or .Range("B4").value = Empty Then 'Load Last Payment Created
            payRow = 4 'On new Payment, GOTO first one created
        Else
            payRow = .Range("B4").value + 1 'Pay Row
        End If
        If MaxPayID = PayID Then 'Last Payment
            MsgBox "Vous êtes au dernier paiement", vbExclamation
            Exit Sub
        End If
        .Range("B3").value = wshENC_Entête.Range("A" & payRow).value 'Set PayID
        Call Encaissement_Load 'Load Payment for the PayID
    End With
    
    Application.EnableEvents = True

    Call End_Timer("modFAC_Enc:Encaissement_Next()", timerStart)

End Sub

Sub Encaissement_Load() '2024-02-14 @ 11:04
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Encaissement_Load()")

    With wshENC_Saisie
        If .Range("B4").value = Empty Then
            MsgBox "Assurez vous de choisir un paiement valide", vbExclamation
            Exit Sub
        End If
        payRow = .Range("B4").value 'Payment Row
        .Range("B4").value = True
        .Range("F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").ClearContents
        'Update worksheet fields
        .Range("J3").value = wshENC_Entête.Cells(payRow, 2).value
        .Range("F3").value = wshENC_Entête.Cells(payRow, 3).value
        .Range("F5").value = wshENC_Entête.Cells(payRow, 4).value
        .Range("J5").value = wshENC_Entête.Cells(payRow, 5).value
        .Range("F7").value = wshENC_Entête.Cells(payRow, 6).value
        
        'Load Pay Items
        With wshENC_Détails
            .Range("M4:T999999").ClearContents
            LastRow = .Range("A999999").End(xlUp).row
            If LastRow < 4 Then GoTo NoData
            .Range("A3:G" & LastRow).AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=.Range("J2:J3"), _
                CopyToRange:=.Range("O3:T3"), _
                Unique:=True
            lastResultRow = .Range("O99999").End(xlUp).row
            If lastResultRow < 4 Then GoTo NoData
            'Bring down the formulas into results
            .Range("M4:N" & lastResultRow).formula = .Range("M1:N1").formula 'Bring Apply and Invoice Date Formulas
            .Range("P4:R" & lastResultRow).formula = .Range("P1:R1").formula 'Inv. Amount, Prev. payments & Balance formulas
            wshENC_Saisie.Range("D13:K" & lastResultRow + 9).value = .Range("M4:T" & lastResultRow).value 'Bring over Pay Items
NoData:
        End With
        .Range("B4").value = False 'Payment Load to False
    End With
    
    Call End_Timer("modFAC_Enc:Encaissement_Load()", timerStart)

End Sub

Sub Encaissement_Import_All() '2024-02-14 @ 09:48
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Encaissement_Import_All()")
    
    Application.ScreenUpdating = False
    
    '3 sheets to import
    Call FAC_Comptes_Clients_Import_All
    Call ENC_Entête_Import_All
    Call ENC_Détails_Import_All
    
    Application.ScreenUpdating = True
    
    Call End_Timer("modFAC_Enc:Encaissement_Import_All()", timerStart)
    
End Sub

Sub Add_Or_Update_Enc_Entete_Record_To_DB(r As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Add_Or_Update_Enc_Entete_Record_To_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_ENC_Entête"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'If r is 0, add a new record, otherwise, update an existing record
    If r = 0 Then 'Add a record
        'SQL select command to find the next available ID
        Dim strSQL As String, MaxID As Long
        strSQL = "SELECT MAX(Pay_ID) AS MaxID FROM [" & destinationTab & "$]"
    
        'Open recordset to find out the MaxID
        rs.Open strSQL, conn
        
        'Get the last used row
        Dim LastRow As Long
        If IsNull(rs.Fields("MaxID").value) Then
            ' Handle empty table (assign a default value, e.g., 1)
            LastRow = 1
        Else
            LastRow = rs.Fields("MaxID").value
        End If
        
        'Calculate the new ID
        Dim nextID As Long
        nextID = LastRow + 1
    
        'Close the previous recordset, no longer needed and open an empty recordset
        rs.Close
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
        
        'Add fields to the recordset before updating it
        rs.AddNew
            rs.Fields("Pay_ID").value = nextID
            rs.Fields("Pay_Date").value = CDate(wshENC_Saisie.Range("J3").value)
            rs.Fields("Customer").value = wshENC_Saisie.Range("F3").value
            rs.Fields("Pay_Type").value = wshENC_Saisie.Range("F5").value
            rs.Fields("Amount").value = Format$(wshENC_Saisie.Range("J5").value, "#,##0.00")
            rs.Fields("Notes").value = wshENC_Saisie.Range("F7").value
    Else 'Update an existing record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & r, conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            rs.Fields("Pay_Date").value = CDate(wshENC_Saisie.Range("J3").value)
            rs.Fields("Customer").value = wshENC_Saisie.Range("F3").value
            rs.Fields("Pay_Type").value = wshENC_Saisie.Range("F5").value
            rs.Fields("Amount").value = Format$(wshENC_Saisie.Range("J5").value, "#,##0.00")
            rs.Fields("Notes").value = wshENC_Saisie.Range("F7").value
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le Pay_ID '" & r & "' ne peut être trouvé!", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    End If
    'Update the recordset (create the record)
    rs.update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    Call End_Timer("modFAC_Enc:Add_Or_Update_Enc_Entete_Record_To_DB()", timerStart)
    
End Sub

Sub Add_Or_Update_Enc_Detail_Record_To_DB(r As Long, encRow As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modFAC_Enc:Add_Or_Update_Enc_Detail_Record_To_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_ENC_Détails"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'If r is 0, add a new record, otherwise, update an existing record
    If r = 0 Then 'Add a record
        'SQL select command to find the next available ID
        Dim strSQL As String, MaxID As Long
        strSQL = "SELECT MAX(Pay_ID) AS MaxID FROM [" & destinationTab & "$]"
    
        'Open recordset to find out the MaxID
        rs.Open strSQL, conn
        
        'Get the last used row
        Dim LastRow As Long
        If IsNull(rs.Fields("MaxID").value) Then
            ' Handle empty table (assign a default value, e.g., 1)
            LastRow = 1
        Else
            LastRow = rs.Fields("MaxID").value
        End If
        
        'Calculate the new ID
        Dim nextID As Long
        nextID = LastRow + 1
    
        'Close the previous recordset, no longer needed and open an empty recordset
        rs.Close
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
        
        'Add fields to the recordset before updating it
        rs.AddNew
            rs.Fields("Pay_ID").value = nextID
            rs.Fields("Inv_No").value = wshENC_Saisie.Range("F" & encRow).value
            rs.Fields("Customer").value = wshENC_Saisie.Range("F3").value
            rs.Fields("Pay_Date").value = CDate(wshENC_Saisie.Range("J3").value)
            rs.Fields("Pay_Amount").value = Format$(wshENC_Saisie.Range("J" & encRow).value, "#,##0.00")
    Else 'Update an existing record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & r, conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            rs.Fields("Inv_No").value = wshENC_Saisie.Range("F" & encRow).value
            rs.Fields("Customer").value = wshENC_Saisie.Range("F3").value
            rs.Fields("Pay_Date").value = CDate(wshENC_Saisie.Range("J3").value)
            rs.Fields("Amount").value = Format$(wshENC_Saisie.Range("J5").value, "#,##0.00")
            rs.Fields("Notes").value = wshENC_Saisie.Range("F7").value
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le Pay_ID '" & r & "' ne peut être trouvé!", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    End If
    'Update the recordset (create the record)
    rs.update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call End_Timer("modFAC_Enc:Add_Or_Update_Enc_Detail_Record_To_DB()", timerStart)
    
End Sub

Sub ENC_Saisie_Add_Check_Boxes(row As Long)

'    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Saisie_Add_Check_Boxes()")
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshENC_Saisie
    
    Dim chkBoxRange As Range: Set chkBoxRange = ws.Range("E12:E" & 11 + row)
    
    Dim cell As Range
    Dim cbx As checkBox
    For Each cell In chkBoxRange
    'Check if the cell is empty and doesn't have a checkbox already
    If Cells(cell.row, 2).value = "" Then 'Applied = False
        'Create a checkbox linked to the cell
        Set cbx = wshENC_Saisie.CheckBoxes.add(cell.Left + 30, cell.Top, cell.width, cell.Height)
        With cbx
            .name = "chkBox - " & cell.row
            .Caption = ""
            .value = False
            .linkedCell = "B" & cell.row
            .Display3DShading = True
            .OnAction = "Apply_Click"
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

Sub ENC_Remove_Check_Boxes(row As Long)

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Remove_Check_Boxes()")
    
    Application.EnableEvents = False
    
    'Delete all checkboxes whose name are chkBox - ...
    Dim cbx As Shape
    For Each cbx In wshENC_Saisie.Shapes
        If InStr(cbx.name, "chkBox -") Then
            cbx.delete
        End If
    Next cbx
    
    wshFAC_Brouillon.Range("B12:B" & row).value = ""
    
    Application.EnableEvents = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cbx = Nothing
    
    Call End_Timer("modENC_Saisie:ENC_Remove_Check_Boxes()", timerStart)

End Sub

Sub ENC_Clear_Cells()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modENC_Saisie:ENC_Clear_Cells()")
    
    With wshENC_Saisie
        'Note the lastUsedRow for checkBox deletion
        Dim lastUsedRow As Long
        lastUsedRow = wshENC_Saisie.Cells(wshENC_Saisie.rows.count, "F").End(xlUp).row
        
        .Range("B4").value = False
        .Range("B5,F5:H5,K5,F7,K7,F9:I9,E12:K31").ClearContents 'Clear Fields
        .Range("B12:B31").ClearContents
        
        If lastUsedRow > 11 Then
            Call ENC_Remove_Check_Boxes(lastUsedRow)
        End If
        
        .Range("K5").value = Date 'Set Default Date
        .Range("F7").value = "Banque" ' Set Default type
        .Range("F5").Select
        
    End With
    
    Call End_Timer("modENC_Saisie:ENC_Clear_Cells()", timerStart)

End Sub

Sub Apply_Click()

    
    Dim chkBox As checkBox
    Set chkBox = ActiveSheet.CheckBoxes(Application.Caller)
    Dim linkedCell As Range
    Set linkedCell = Range(chkBox.linkedCell)
    
    If linkedCell.value = True Then
        Dim MAA As Currency
        'Calculate the maximum amount to applied (MAA)
        If wshENC_Saisie.Range("K9").value > 0 Then
            If wshENC_Saisie.Range("K9").value > wshENC_Saisie.Range("J" & linkedCell.row).value Then
                wshENC_Saisie.Range("K" & linkedCell.row).value = wshENC_Saisie.Range("J" & linkedCell.row).value
            Else
                wshENC_Saisie.Range("K" & linkedCell.row).value = wshENC_Saisie.Range("K9").value
            End If
        End If
    Else
        Range("K" & linkedCell.row).value = 0
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


