Attribute VB_Name = "modEncaissement"
Option Explicit
Dim lastRow As Long, lastResultRow As Long
Dim payRow As Long
Dim resultRow As Long, payItemRow As Long, lastPayItemRow As Long, payitemDBRow As Long

Sub Encaissement_Load_Open_Invoices() '2024-02-20 @ 14:09
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Encaissement_Load_Open_Invoices()")
    
    wshENC_Saisie.Range("D13:K42").Clearcontents 'Clear the invoices area before loading it
    With wshCC
        lastResultRow = .Range("A99999").End(xlUp).row 'Last row
        If lastResultRow < 3 Then Exit Sub
        'Cells L3 contains a formula, no need to set it up
        .Range("A2:J" & lastResultRow).AdvancedFilter _
            xlFilterCopy, _
            CriteriaRange:=.Range("L2:M3"), _
            CopyToRange:=.Range("O2:S2")
        lastResultRow = .Range("O9999").End(xlUp).row
        If lastResultRow < 3 Then Exit Sub
        wshENC_Saisie.Range("B2").Value = True 'Set PaymentLoad to True
        .Range("R3:R" & lastResultRow).formula = .Range("R1").formula 'Total Payments Formula
        'Bring the Result data into our Payments List of Invoices
        wshENC_Saisie.Range("E13:I" & lastResultRow + 10).Value = .Range("O3:S" & lastResultRow).Value
    End With
    wshENC_Saisie.Range("B2").Value = False 'Set PaymentLoad to False
    
    Call Output_Timer_Results("modEncaissement:Encaissement_Load_Open_Invoices()", timerStart)

End Sub

Sub Encaissement_Save_Update() '2024-02-07 @ 12:27
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Encaissement_Save_Update()")
    
    With wshENC_Saisie
        'Check for mandatory fields (4)
        If .Range("F3").Value = Empty Or _
           .Range("J3").Value = Empty Or _
           .Range("J3").Value = Empty Then
            MsgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client" & vbNewLine & _
                "2. Une date de paiement" & vbNewLine & _
                "3. Un type de paiement et" & vbNewLine & _
                "4. Des transactions" & vbNewLine & vbNewLine & _
                "AVANT de sauvegarder la transaction.", vbExclamation
            Exit Sub
        End If
        'Check to make sure Payment Amount = Applied Amount
        If .Range("J5").Value <> .Range("J10").Value Then
            MsgBox "Assurez-vous que le montant du paiement soit �GAL" & vbNewLine & _
                "� la somme des paiements appliqu�s", vbExclamation
            Exit Sub
        End If
        'New Payment -OR- Existing Payment ?
        If .Range("B4").Value = Empty Then 'New Payment
            payRow = wshENC_Ent�te.Range("A999999").End(xlUp).row + 1 'First Available Row
            .Range("B3").Value = .Range("B5").Value 'Next payment ID
            'wshENC_Ent�te.Range("A" & payRow).Value = .Range("B3").Value 'PayID
            Call Add_Or_Update_Enc_Entete_Record_To_DB(0)
        Else 'Existing Payment
            Call Add_Or_Update_Enc_Entete_Record_To_DB(.Range("B4").Value)
        End If
        
        'Save Applied Invoices to Payment Detail
        lastPayItemRow = .Range("E999999").End(xlUp).row 'Last Pay Item
        
        For payItemRow = 13 To lastPayItemRow
            If .Range("D" & payItemRow).Value = Chr(252) Then 'The row has been applied
                If .Range("K" & payItemRow).Value = Empty Then 'New Pay Item row
                    payitemDBRow = wshENC_D�tails.Range("A999999").End(xlUp).row + 1 'First Avail Pay Items Row
                    Call Add_Or_Update_Enc_Detail_Record_To_DB(0, payItemRow)
                    'wshENC_D�tails.Range("A" & payitemDBRow).Value = .Range("B3").Value 'Payment ID
                    'wshENC_D�tails.Range("F" & payitemDBRow).Value = "=row()"
                    .Range("K" & payItemRow).Value = payitemDBRow 'Database Row
                Else 'Existing Pay Item
                    payitemDBRow = .Range("K" & payItemRow).Value 'Existing Pay Item Row
                End If
'                wshENC_D�tails.Range("B" & payitemDBRow).Value = .Range("F" & payItemRow).Value 'Invoice ID
'                wshENC_D�tails.Range("C" & payitemDBRow).Value = .Range("F3").Value 'Customer
'                wshENC_D�tails.Range("D" & payitemDBRow).Value = .Range("J3").Value 'Pay Date
'                wshENC_D�tails.Range("E" & payitemDBRow).Value = .Range("J" & payItemRow).Value 'Amount paid
            End If
        Next payItemRow
        
        'Prepare G/L posting
        Dim noEnc As String, nomCLient As String, typeEnc As String, descEnc As String
        Dim dateEnc As Date
        Dim montantEnc As Currency
        noEnc = wshENC_Saisie.Range("B5").Value
        dateEnc = wshENC_Saisie.Range("J3").Value
        nomCLient = wshENC_Saisie.Range("F3").Value
        typeEnc = wshENC_Saisie.Range("F5").Value
        montantEnc = wshENC_Saisie.Range("J5").Value
        descEnc = wshENC_Saisie.Range("F7").Value

        Call Encaissement_GL_Posting(noEnc, dateEnc, nomCLient, typeEnc, montantEnc, descEnc)  '2024-02-09 @ 08:17 - TODO
        
        Call Encaissement_Import_All   'Bring back locally three worksheets
        
        MsgBox "Le paiement a �t� renregistr� avec succ�s"
        Call Encaissement_Add_New 'Reset the form
        .Range("F3").Select
    End With
    
    Call Output_Timer_Results("modEncaissement:Encaissement_Save_Update()", timerStart)

End Sub

Sub Encaissement_Add_New() '2024-02-07 @ 12:39

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Encaissement_Add_New()")

    wshENC_Saisie.Range("B2").Value = False
    wshENC_Saisie.Range("B3,F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").Clearcontents 'Clear Fields
    wshENC_Saisie.Range("J3").Value = Date 'Set Default Date
    wshENC_Saisie.Range("F5").Value = "Banque" ' Set Default type
    
    Call Output_Timer_Results("modEncaissement:Encaissement_Add_New()", timerStart)
    
End Sub

Sub Encaissement_Previous() '2024-02-14 @ 11:04
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Encaissement_Previous()")

    Dim MinPayID As Long, PayID As Long
    With wshENC_Saisie
        On Error Resume Next
            MinPayID = Application.WorksheetFunction.Min(wshENC_Ent�te.Range("Pay_ID"))
        On Error GoTo 0
        If MinPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistr�", vbExclamation
            Exit Sub
        End If
        PayID = .Range("B3").Value 'Payment ID
        If PayID = 0 Or .Range("B4").Value = Empty Then 'Load Last Payment Created
            payRow = wshENC_Ent�te.Range("A99999").End(xlUp).row 'Last Row
        Else
            payRow = .Range("B4").Value - 1 'Pay Row
        End If
        If payRow = 3 Or MinPayID = .Range("B3").Value Then 'First Payment
            MsgBox "Vous �tes au premier paiement", vbExclamation
            Exit Sub
        End If
        .Range("B3").Value = wshENC_Ent�te.Range("A" & payRow).Value 'Set Payment ID
        Call Encaissement_Load 'Load Payment
    End With

    Call Output_Timer_Results("modEncaissement:Encaissement_Previous()", timerStart)

End Sub

Sub Encaissement_Next() '2024-02-14 @ 11:04
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Encaissement_Next()")

    Application.EnableEvents = False

    Dim MaxPayID As Long
    With wshENC_Saisie
        On Error Resume Next
            MaxPayID = Application.WorksheetFunction.Max(wshENC_Ent�te.Range("Pay_ID"))
        On Error GoTo 0
        If MaxPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistr�", vbExclamation
            Exit Sub
        End If
        Dim PayID As Long
        PayID = .Range("B3").Value 'Payment ID
        If PayID = 0 Or .Range("B4").Value = Empty Then 'Load Last Payment Created
            payRow = 4 'On new Payment, GOTO first one created
        Else
            payRow = .Range("B4").Value + 1 'Pay Row
        End If
        If MaxPayID = PayID Then 'Last Payment
            MsgBox "Vous �tes au dernier paiement", vbExclamation
            Exit Sub
        End If
        .Range("B3").Value = wshENC_Ent�te.Range("A" & payRow).Value 'Set PayID
        Call Encaissement_Load 'Load Payment for the PayID
    End With
    
    Application.EnableEvents = True

    Call Output_Timer_Results("modEncaissement:Encaissement_Next()", timerStart)

End Sub

Sub Encaissement_Load() '2024-02-14 @ 11:04
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Encaissement_Load()")

    With wshENC_Saisie
        If .Range("B4").Value = Empty Then
            MsgBox "Assurez vous de choisir un paiement valide", vbExclamation
            Exit Sub
        End If
        payRow = .Range("B4").Value 'Payment Row
        .Range("B2").Value = True
        .Range("F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").Clearcontents
        'Update worksheet fields
        .Range("J3").Value = wshENC_Ent�te.Cells(payRow, 2).Value
        .Range("F3").Value = wshENC_Ent�te.Cells(payRow, 3).Value
        .Range("F5").Value = wshENC_Ent�te.Cells(payRow, 4).Value
        .Range("J5").Value = wshENC_Ent�te.Cells(payRow, 5).Value
        .Range("F7").Value = wshENC_Ent�te.Cells(payRow, 6).Value
        
        'Load Pay Items
        With wshENC_D�tails
            .Range("M4:T999999").Clearcontents
            lastRow = .Range("A999999").End(xlUp).row
            If lastRow < 4 Then GoTo NoData
            .Range("A3:G" & lastRow).AdvancedFilter _
                xlFilterCopy, _
                CriteriaRange:=.Range("J2:J3"), _
                CopyToRange:=.Range("O3:T3"), _
                Unique:=True
            lastResultRow = .Range("O99999").End(xlUp).row
            If lastResultRow < 4 Then GoTo NoData
            'Bring down the formulas into results
            .Range("M4:N" & lastResultRow).formula = .Range("M1:N1").formula 'Bring Apply and Invoice Date Formulas
            .Range("P4:R" & lastResultRow).formula = .Range("P1:R1").formula 'Inv. Amount, Prev. payments & Balance formulas
            wshENC_Saisie.Range("D13:K" & lastResultRow + 9).Value = .Range("M4:T" & lastResultRow).Value 'Bring over Pay Items
NoData:
        End With
        .Range("B2").Value = False 'Payment Load to False
    End With
    
    Call Output_Timer_Results("modEncaissement:Encaissement_Load()", timerStart)

End Sub

Sub Encaissement_Import_All() '2024-02-14 @ 09:48
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Encaissement_Import_All()")
    
    Application.ScreenUpdating = False
    
    '3 sheets to import
    Call FAC_Comptes_Clients_Import_All
    Call FAC_Encaissements_Ent�te_Import_All
    Call FAC_Encaissements_D�tails_Import_All
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("modEncaissement:Encaissement_Import_All()", timerStart)
    
End Sub
'Sub FAC_Comptes_Clients_Import_All() '2024-02-14 @ 09:50
'
'    'Clear all cells, but the headers, in the destination worksheet
'    wshCC.Range("A1").CurrentRegion.Offset(2, 0).ClearContents
'
'    'Import AR_Summary from 'GCF_DB_Sortie.xlsx'
'    Dim sourceWorkbook As String, sourceTab As String
'    sourceWorkbook = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
'                     "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
'    sourceTab = "FAC_Comptes_Clients"
'
'    'Set up source and destination ranges
'    Dim sourceRange As Range, destinationRange As Range
'    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
'    Set destinationRange = wshCC.Range("A2")
'
'    'Copy data, using Range to Range and Autofit all columns
'    sourceRange.Copy destinationRange
'    wshCC.Range("A1").CurrentRegion.EntireColumn.AutoFit
'
'    'Close the source workbook, without saving it
'    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False
'
'    'Insert Formula in column H
'    Dim lastRow As Long
'    lastRow = wshCC.Range("A99999").End(xlUp).row
'    'Check if there is data in column A
'    If lastRow < 3 Then
'        MsgBox "No data found in column A.", vbExclamation
'        Exit Sub
'    End If
'    wshCC.Range("H3:H" & lastRow).formula = "=SUMIFS(pmnt_Amount,pmnt_invNumb, $A3)"
'
''    'Define the named ranges for Pmnt_Amount and Pmnt_invNumb outside of the loop
''    Dim pmnt_Amount_Range As Range
''    Dim Pmnt_invNumb_Range As Range
''    With wshENC_D�tails
''        Set pmnt_Amount_Range = .Range("Pmnt_Amount")
''        Set Pmnt_invNumb_Range = .Range("Pmnt_invNumb")
''    End With
''
''    Dim cell As Range
''    'Loop through each cell in the range H3 to H[lastRow]
''    For Each cell In wshCC.Range("H3:H" & lastRow)
''        'Assign the formula to each cell individually using the Formula property
''        cell.formula = "=SUMIFS('" & wshENC_D�tails.name & "'!" & pmnt_Amount_Range.Address & "," & _
''                               "'" & wshENC_D�tails.name & "'!" & Pmnt_invNumb_Range.Address & "," & _
''                               "'" & wshCC.name & "'!$A" & cell.row & ")"
''        Debug.Print cell.Address
''    Next cell
'
''    With wshCC
''        .Range("A3" & ":F" & lastRow).HorizontalAlignment = xlCenter
''        With .Range("C3:C" & lastRow & ",D3:D" & lastRow & ",E3:E" & lastRow)
''            .HorizontalAlignment = xlLeft
''        End With
''        .Range("G3:G" & lastRow & ",I3:I" & lastRow).HorizontalAlignment = xlRight
''        .Range("G3:I" & lastRow).NumberFormat = "#,##0.00 $"
''        .Range("B3:B" & lastRow & ",F3:F" & lastRow).NumberFormat = "dd/mm/yyyy"
''    End With
'
'End Sub
'
Sub FAC_Encaissements_Ent�te_Import_All() '2024-02-14 @ 10:05
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:FAC_Encaissements_Ent�te_Import_All()")
    
    'Clear all cells, but the headers, in the destination worksheet
    wshENC_Ent�te.Range("A1").CurrentRegion.Offset(3, 0).Clearcontents

    'Import AR_Summary from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
    sourceTab = "FAC_Encaissements_Ent�te"
    
    'Set up source and destination ranges
    Dim sourceRange As Range, destinationRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
    Set destinationRange = wshENC_Ent�te.Range("A3")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshENC_Ent�te.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

'    'Arrange formats on all rows
'    Dim lastRow As Long
'    lastRow = wshENC_Ent�te.Range("A999999").End(xlUp).row
'
'    With wshENC_Ent�te
'        .Range("A4" & ":B" & lastRow).HorizontalAlignment = xlCenter
'        With .Range("C4:C" & lastRow & ",D4:D" & lastRow & ",F4:F" & lastRow)
'            .HorizontalAlignment = xlLeft
'        End With
'        .Range("E4:E" & lastRow).HorizontalAlignment = xlRight
'        .Range("G4:H" & lastRow).NumberFormat = "#,##0.00 $"
'        .Range("B4:B" & lastRow).NumberFormat = "dd/mm/yyyy"
'        .Range("F4:F" & lastRow).NumberFormat = "dd/mm/yyyy"
'    End With
    
    Call Output_Timer_Results("modEncaissement:FAC_Encaissements_Ent�te_Import_All()", timerStart)
  
End Sub

Sub FAC_Encaissements_D�tails_Import_All() '2024-02-14 @ 10:14
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:FAC_Encaissements_D�tails_Import_All()")
    
    'Clear all cells, but the headers, in the destination worksheet
    wshENC_D�tails.Range("A1").CurrentRegion.Offset(3, 0).Clearcontents

    'Import AR_Summary from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
    sourceTab = "FAC_Encaissements_D�tails"
    
    'Set up source and destination ranges
    Dim sourceRange As Range, destinationRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
    Set destinationRange = wshENC_D�tails.Range("A3")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshENC_D�tails.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

'    'Arrange formats on all rows
'    Dim lastRow As Long
'    lastRow = wshENC_D�tails.Range("A999999").End(xlUp).row
'
'    With wshENC_D�tails
'        .Range("A4:B" & lastRow & ",D4:D" & lastRow & ",F4:F" & lastRow).HorizontalAlignment = xlCenter
'        .Range("C4:C" & lastRow).HorizontalAlignment = xlLeft
'        .Range("D3:D" & lastRow).NumberFormat = "dd/mm/yyyy"
'        .Range("E3:E" & lastRow).HorizontalAlignment = xlRight
'        .Range("E3:E" & lastRow).NumberFormat = "#,##0.00 $"
'    End With
    
    Call Output_Timer_Results("modEncaissement:FAC_Encaissements_D�tails_Import_All()", timerStart)
    
End Sub

Sub Add_Or_Update_Enc_Entete_Record_To_DB(r As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Add_Or_Update_Enc_Entete_Record_To_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_Encaissements_Ent�te"
    
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
        Dim lastRow As Long
        If IsNull(rs.Fields("MaxID").Value) Then
            ' Handle empty table (assign a default value, e.g., 1)
            lastRow = 1
        Else
            lastRow = rs.Fields("MaxID").Value
        End If
        
        'Calculate the new ID
        Dim nextID As Long
        nextID = lastRow + 1
    
        'Close the previous recordset, no longer needed and open an empty recordset
        rs.Close
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
        
        'Add fields to the recordset before updating it
        rs.AddNew
            rs.Fields("Pay_ID").Value = nextID
            rs.Fields("Pay_Date").Value = CDate(wshENC_Saisie.Range("J3").Value)
            rs.Fields("Customer").Value = wshENC_Saisie.Range("F3").Value
            rs.Fields("Pay_Type").Value = wshENC_Saisie.Range("F5").Value
            rs.Fields("Amount").Value = Format(wshENC_Saisie.Range("J5").Value, "#,##0.00")
            rs.Fields("Notes").Value = wshENC_Saisie.Range("F7").Value
    Else 'Update an existing record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & r, conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            rs.Fields("Pay_Date").Value = CDate(wshENC_Saisie.Range("J3").Value)
            rs.Fields("Customer").Value = wshENC_Saisie.Range("F3").Value
            rs.Fields("Pay_Type").Value = wshENC_Saisie.Range("F5").Value
            rs.Fields("Amount").Value = Format(wshENC_Saisie.Range("J5").Value, "#,##0.00")
            rs.Fields("Notes").Value = wshENC_Saisie.Range("F7").Value
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le Pay_ID '" & r & "' ne peut �tre trouv�!", vbExclamation
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

    Call Output_Timer_Results("modEncaissement:Add_Or_Update_Enc_Entete_Record_To_DB()", timerStart)
    
End Sub

Sub Add_Or_Update_Enc_Detail_Record_To_DB(r As Long, encRow As Long) 'Write -OR- Update a record to external .xlsx file
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modEncaissement:Add_Or_Update_Enc_Detail_Record_To_DB()")
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_Encaissements_D�tails"
    
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
        Dim lastRow As Long
        If IsNull(rs.Fields("MaxID").Value) Then
            ' Handle empty table (assign a default value, e.g., 1)
            lastRow = 1
        Else
            lastRow = rs.Fields("MaxID").Value
        End If
        
        'Calculate the new ID
        Dim nextID As Long
        nextID = lastRow + 1
    
        'Close the previous recordset, no longer needed and open an empty recordset
        rs.Close
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
        
        'Add fields to the recordset before updating it
        rs.AddNew
            rs.Fields("Pay_ID").Value = nextID
            rs.Fields("Inv_No").Value = wshENC_Saisie.Range("F" & encRow).Value
            rs.Fields("Customer").Value = wshENC_Saisie.Range("F3").Value
            rs.Fields("Pay_Date").Value = CDate(wshENC_Saisie.Range("J3").Value)
            rs.Fields("Pay_Amount").Value = Format(wshENC_Saisie.Range("J" & encRow).Value, "#,##0.00")
    Else 'Update an existing record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & r, conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            rs.Fields("Inv_No").Value = wshENC_Saisie.Range("F" & encRow).Value
            rs.Fields("Customer").Value = wshENC_Saisie.Range("F3").Value
            rs.Fields("Pay_Date").Value = CDate(wshENC_Saisie.Range("J3").Value)
            rs.Fields("Amount").Value = Format(wshENC_Saisie.Range("J5").Value, "#,##0.00")
            rs.Fields("Notes").Value = wshENC_Saisie.Range("F7").Value
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le Pay_ID '" & r & "' ne peut �tre trouv�!", vbExclamation
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

    Call Output_Timer_Results("modEncaissement:Add_Or_Update_Enc_Detail_Record_To_DB()", timerStart)
    
End Sub
