Attribute VB_Name = "modEncaissement"
Option Explicit
Dim lastrow As Long, LastResultRow As Long
Dim PayRow As Long, PayCol As Long
Dim resultRow As Long, PayItemRow As Long, LastPayItemRow As Long, PayItemDBRow As Long

Sub Encaissement_Load_Open_Invoices() '2024-02-20 @ 14:09
    wshEncaissement.Range("D13:K42").ClearContents 'Clear the invoices area before loading it
    With wshAR
        LastResultRow = .Range("A99999").End(xlUp).row 'Last row
        If LastResultRow < 3 Then Exit Sub
        'Cells L3 contains a formula, no need to set it up
        .Range("A2:K" & LastResultRow).AdvancedFilter _
            xlFilterCopy, _
            CriteriaRange:=.Range("L2:M3"), _
            CopyToRange:=.Range("O2:T2"), _
            Unique:=True
        LastResultRow = .Range("O99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        wshEncaissement.Range("B2").value = True 'Set PaymentLoad to True
        .Range("R3:R" & LastResultRow).formula = .Range("R1").formula 'Total Payments Formula
        'Bring the Result data into our Payments List of Invoices
        wshEncaissement.Range("E13:I" & LastResultRow + 10).value = .Range("O3:S" & LastResultRow).value
    End With
    wshEncaissement.Range("B2").value = False 'Set PaymentLoad to False
End Sub

Sub Encaissement_Save_Update() '2024-02-07 @ 12:27
    With wshEncaissement
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
            PayRow = wshEncEntete.Range("A999999").End(xlUp).row + 1 'First Available Row
            .Range("B3").value = .Range("B5").value 'Next payment ID
            'wshEncEntete.Range("A" & PayRow).value = .Range("B3").value 'PayID
            Call Add_Or_Update_Enc_Entete_Record_To_DB(0)
        Else 'Existing Payment
            Call Add_Or_Update_Enc_Entete_Record_To_DB(.Range("B4").value)
        End If
        
        'Save Applied Invoices to Payment Detail
        LastPayItemRow = .Range("E999999").End(xlUp).row 'Last Pay Item
        
        For PayItemRow = 13 To LastPayItemRow
            If .Range("D" & PayItemRow).value = Chr(252) Then 'The row has been applied
                If .Range("K" & PayItemRow).value = Empty Then 'New Pay Item row
                    PayItemDBRow = wshEncDetail.Range("A999999").End(xlUp).row + 1 'First Avail Pay Items Row
                    Call Add_Or_Update_Enc_Detail_Record_To_DB(0, PayItemRow)
                    'wshEncDetail.Range("A" & PayItemDBRow).value = .Range("B3").value 'Payment ID
                    'wshEncDetail.Range("F" & PayItemDBRow).value = "=row()"
                    .Range("K" & PayItemRow).value = PayItemDBRow 'Database Row
                Else 'Existing Pay Item
                    PayItemDBRow = .Range("K" & PayItemRow).value 'Existing Pay Item Row
                End If
'                wshEncDetail.Range("B" & PayItemDBRow).value = .Range("F" & PayItemRow).value 'Invoice ID
'                wshEncDetail.Range("C" & PayItemDBRow).value = .Range("F3").value 'Customer
'                wshEncDetail.Range("D" & PayItemDBRow).value = .Range("J3").value 'Pay Date
'                wshEncDetail.Range("E" & PayItemDBRow).value = .Range("J" & PayItemRow).value 'Amount paid
            End If
        Next PayItemRow
        
        'Prepare G/L posting
        Dim noEnc As String, nomCLient As String, typeEnc As String, descEnc As String
        Dim dateEnc As Date
        Dim montantEnc As Currency
        noEnc = wshEncaissement.Range("B5").value
        dateEnc = wshEncaissement.Range("J3").value
        nomCLient = wshEncaissement.Range("F3").value
        typeEnc = wshEncaissement.Range("F5").value
        montantEnc = wshEncaissement.Range("J5").value
        descEnc = wshEncaissement.Range("F7").value

        Call Encaissement_GL_Posting(noEnc, dateEnc, nomCLient, typeEnc, montantEnc, descEnc)  '2024-02-09 @ 08:17 - TODO
        
        Call Encaissement_Import_All   'Bring back locally three worksheets
        
        MsgBox "Le paiement a été renregistré avec succès"
        Call Encaissement_Add_New 'Reset the form
        .Range("F3").Select
    End With
End Sub

Sub Encaissement_Add_New() '2024-02-07 @ 12:39
    wshEncaissement.Range("B2").value = False
    wshEncaissement.Range("B3,F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").ClearContents 'Clear Fields
    wshEncaissement.Range("J3").value = Date 'Set Default Date
    wshEncaissement.Range("F5").value = "Banque" ' Set Default type
End Sub

Sub Encaissement_Previous() '2024-02-14 @ 11:04
    Dim MinPayID As Long, PayID As Long
    With wshEncaissement
        On Error Resume Next
            MinPayID = Application.WorksheetFunction.Min(wshEncEntete.Range("Pay_ID"))
        On Error GoTo 0
        If MinPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
            Exit Sub
        End If
        PayID = .Range("B3").value 'Payment ID
        If PayID = 0 Or .Range("B4").value = Empty Then 'Load Last Payment Created
            PayRow = wshEncEntete.Range("A99999").End(xlUp).row 'Last Row
        Else
            PayRow = .Range("B4").value - 1 'Pay Row
        End If
        If PayRow = 3 Or MinPayID = .Range("B3").value Then 'First Payment
            MsgBox "Vous êtes au premier paiement", vbExclamation
            Exit Sub
        End If
        .Range("B3").value = wshEncEntete.Range("A" & PayRow).value 'Set Payment ID
        Call Encaissement_Load 'Load Payment
    End With
End Sub

Sub Encaissement_Next() '2024-02-14 @ 11:04
    
    Application.EnableEvents = False

    Dim MaxPayID As Long
    With wshEncaissement
        On Error Resume Next
            MaxPayID = Application.WorksheetFunction.Max(wshEncEntete.Range("Pay_ID"))
        On Error GoTo 0
        If MaxPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
            Exit Sub
        End If
        Dim PayID As Long
        PayID = .Range("B3").value 'Payment ID
        If PayID = 0 Or .Range("B4").value = Empty Then 'Load Last Payment Created
            PayRow = 4 'On new Payment, GOTO first one created
        Else
            PayRow = .Range("B4").value + 1 'Pay Row
        End If
        If MaxPayID = PayID Then 'Last Payment
            MsgBox "Vous êtes au dernier paiement", vbExclamation
            Exit Sub
        End If
        .Range("B3").value = wshEncEntete.Range("A" & PayRow).value 'Set PayID
        Call Encaissement_Load 'Load Payment for the PayID
    End With
    
    Application.EnableEvents = True

End Sub

Sub Encaissement_Load() '2024-02-14 @ 11:04
    With wshEncaissement
        If .Range("B4").value = Empty Then
            MsgBox "Assurez vous de choisir un paiement valide", vbExclamation
            Exit Sub
        End If
        PayRow = .Range("B4").value 'Payment Row
        .Range("B2").value = True
        .Range("F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").ClearContents
        'Update worksheet fields
        .Range("J3").value = wshEncEntete.Cells(PayRow, 2).value
        .Range("F3").value = wshEncEntete.Cells(PayRow, 3).value
        .Range("F5").value = wshEncEntete.Cells(PayRow, 4).value
        .Range("J5").value = wshEncEntete.Cells(PayRow, 5).value
        .Range("F7").value = wshEncEntete.Cells(PayRow, 6).value
        
        'Load Pay Items
        With wshEncDetail
            .Range("M4:T999999").ClearContents
            lastrow = .Range("A999999").End(xlUp).row
            If lastrow < 4 Then GoTo NoData
            .Range("A3:G" & lastrow).AdvancedFilter _
                xlFilterCopy, _
                CriteriaRange:=.Range("J2:J3"), _
                CopyToRange:=.Range("O3:T3"), _
                Unique:=True
            LastResultRow = .Range("O99999").End(xlUp).row
            If LastResultRow < 4 Then GoTo NoData
            'Bring down the formulas into results
            .Range("M4:N" & LastResultRow).formula = .Range("M1:N1").formula 'Bring Apply and Invoice Date Formulas
            .Range("P4:R" & LastResultRow).formula = .Range("P1:R1").formula 'Inv. Amount, Prev. payments & Balance formulas
            wshEncaissement.Range("D13:K" & LastResultRow + 9).value = .Range("M4:T" & LastResultRow).value 'Bring over Pay Items
NoData:
        End With
        .Range("B2").value = False 'Payment Load to False
    End With
End Sub

Sub Encaissement_Import_All() '2024-02-14 @ 09:48
    
    '3 sheets to import
    Application.ScreenUpdating = False
    
    Call AR_Summary_Import_All
    Call Enc_Entete_Import_All
    Call Enc_Detail_Import_All
    
    Application.ScreenUpdating = True
    
End Sub
Sub AR_Summary_Import_All() '2024-02-14 @ 09:50
    
    'Clear all cells, but the headers, in the destination worksheet
    wshAR.Range("A1").CurrentRegion.Offset(2, 0).ClearContents

    'Import AR_Summary from 'GCF_DB_Sortie.xlsx'
    Dim fileName As String, sourceWorkbook As String, sourceTab As String
    fileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
    sourceWorkbook = fileName
    sourceTab = "Comptes_Clients"
    
    'Set up source and destination ranges
    Dim sourceRange As Range, destinationRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
    Set destinationRange = wshAR.Range("A2")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshAR.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Insert Formula in column H
    Dim lastrow As Long
    lastrow = wshAR.Range("A99999").End(xlUp).row
    'Check if there is data in column A
    If lastrow < 3 Then
        MsgBox "No data found in column A.", vbExclamation
        Exit Sub
    End If
    wshAR.Range("H3:H" & lastrow).formula = "=SUMIFS(pmnt_Amount,pmnt_InvNumb, $A3)"

'    'Define the named ranges for Pmnt_Amount and Pmnt_InvNumb outside of the loop
'    Dim pmnt_Amount_Range As Range
'    Dim Pmnt_InvNumb_Range As Range
'    With wshEncDetail
'        Set pmnt_Amount_Range = .Range("Pmnt_Amount")
'        Set Pmnt_InvNumb_Range = .Range("Pmnt_InvNumb")
'    End With
'
'    Dim cell As Range
'    'Loop through each cell in the range H3 to H[lastrow]
'    For Each cell In wshAR.Range("H3:H" & lastrow)
'        'Assign the formula to each cell individually using the Formula property
'        cell.formula = "=SUMIFS('" & wshEncDetail.name & "'!" & pmnt_Amount_Range.Address & "," & _
'                               "'" & wshEncDetail.name & "'!" & Pmnt_InvNumb_Range.Address & "," & _
'                               "'" & wshAR.name & "'!$A" & cell.row & ")"
'        Debug.Print cell.Address
'    Next cell

'    With wshAR
'        .Range("A3" & ":F" & lastrow).HorizontalAlignment = xlCenter
'        With .Range("C3:C" & lastrow & ",D3:D" & lastrow & ",E3:E" & lastrow)
'            .HorizontalAlignment = xlLeft
'        End With
'        .Range("G3:G" & lastrow & ",I3:I" & lastrow).HorizontalAlignment = xlRight
'        .Range("G3:I" & lastrow).NumberFormat = "#,##0.00 $"
'        .Range("B3:B" & lastrow & ",F3:F" & lastrow).NumberFormat = "dd/mm/yyyy"
'    End With
    
End Sub

Sub Enc_Entete_Import_All() '2024-02-14 @ 10:05
    
    'Clear all cells, but the headers, in the destination worksheet
    wshEncEntete.Range("A1").CurrentRegion.Offset(3, 0).ClearContents

    'Import AR_Summary from 'GCF_DB_Sortie.xlsx'
    Dim fileName As String, sourceWorkbook As String, sourceTab As String
    fileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
    sourceWorkbook = fileName
    sourceTab = "Encaissements_Entête"
    
    'Set up source and destination ranges
    Dim sourceRange As Range, destinationRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
    Set destinationRange = wshEncEntete.Range("A3")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshEncEntete.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

'    'Arrange formats on all rows
'    Dim lastrow As Long
'    lastrow = wshEncEntete.Range("A999999").End(xlUp).row
'
'    With wshEncEntete
'        .Range("A4" & ":B" & lastrow).HorizontalAlignment = xlCenter
'        With .Range("C4:C" & lastrow & ",D4:D" & lastrow & ",F4:F" & lastrow)
'            .HorizontalAlignment = xlLeft
'        End With
'        .Range("E4:E" & lastrow).HorizontalAlignment = xlRight
'        .Range("G4:H" & lastrow).NumberFormat = "#,##0.00 $"
'        .Range("B4:B" & lastrow).NumberFormat = "dd/mm/yyyy"
'        .Range("F4:F" & lastrow).NumberFormat = "dd/mm/yyyy"
'    End With
    
End Sub

Sub Enc_Detail_Import_All() '2024-02-14 @ 10:14
    
    'Clear all cells, but the headers, in the destination worksheet
    wshEncDetail.Range("A1").CurrentRegion.Offset(3, 0).ClearContents

    'Import AR_Summary from 'GCF_DB_Sortie.xlsx'
    Dim fileName As String, sourceWorkbook As String, sourceTab As String
    fileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                "GCF_BD_Sortie.xlsx" '2024-02-14 @ 06:22
    sourceWorkbook = fileName
    sourceTab = "Encaissements_Détail"
    
    'Set up source and destination ranges
    Dim sourceRange As Range, destinationRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).usedRange
    Set destinationRange = wshEncDetail.Range("A3")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshEncDetail.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

'    'Arrange formats on all rows
'    Dim lastrow As Long
'    lastrow = wshEncDetail.Range("A999999").End(xlUp).row
'
'    With wshEncDetail
'        .Range("A4:B" & lastrow & ",D4:D" & lastrow & ",F4:F" & lastrow).HorizontalAlignment = xlCenter
'        .Range("C4:C" & lastrow).HorizontalAlignment = xlLeft
'        .Range("D3:D" & lastrow).NumberFormat = "dd/mm/yyyy"
'        .Range("E3:E" & lastrow).HorizontalAlignment = xlRight
'        .Range("E3:E" & lastrow).NumberFormat = "#,##0.00 $"
'    End With
    
End Sub

Sub Add_Or_Update_Enc_Entete_Record_To_DB(r As Long) 'Write -OR- Update a record to external .xlsx file
    
    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    sheetName = "Encaissements_Entête"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'If r is 0, add a new record, otherwise, update an existing record
    If r = 0 Then 'Add a record
        'SQL select command to find the next available ID
        Dim strSQL As String, MaxID As Long
        strSQL = "SELECT MAX(Pay_ID) AS MaxID FROM [" & sheetName & "$]"
    
        'Open recordset to find out the MaxID
        rs.Open strSQL, conn
        
        'Get the last used row
        Dim lastrow As Long
        If IsNull(rs.Fields("MaxID").value) Then
            ' Handle empty table (assign a default value, e.g., 1)
            lastrow = 1
        Else
            lastrow = rs.Fields("MaxID").value
        End If
        
        'Calculate the new ID
        Dim nextID As Long
        nextID = lastrow + 1
    
        'Close the previous recordset, no longer needed and open an empty recordset
        rs.Close
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
        
        'Add fields to the recordset before updating it
        rs.AddNew
            rs.Fields("Pay_ID").value = nextID
            rs.Fields("Pay_Date").value = CDate(wshEncaissement.Range("J3").value)
            rs.Fields("Customer").value = wshEncaissement.Range("F3").value
            rs.Fields("Pay_Type").value = wshEncaissement.Range("F5").value
            rs.Fields("Amount").value = Format(wshEncaissement.Range("J5").value, "#,##0.00")
            rs.Fields("Notes").value = wshEncaissement.Range("F7").value
    Else 'Update an existing record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE TEC_ID=" & r, conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            rs.Fields("Pay_Date").value = CDate(wshEncaissement.Range("J3").value)
            rs.Fields("Customer").value = wshEncaissement.Range("F3").value
            rs.Fields("Pay_Type").value = wshEncaissement.Range("F5").value
            rs.Fields("Amount").value = Format(wshEncaissement.Range("J5").value, "#,##0.00")
            rs.Fields("Notes").value = wshEncaissement.Range("F7").value
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le Pay_ID '" & r & "' ne peut être trouvé!", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    End If
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

Sub Add_Or_Update_Enc_Detail_Record_To_DB(r As Long, encRow As Long) 'Write -OR- Update a record to external .xlsx file
    
    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    sheetName = "Encaissements_Détail"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'If r is 0, add a new record, otherwise, update an existing record
    If r = 0 Then 'Add a record
        'SQL select command to find the next available ID
        Dim strSQL As String, MaxID As Long
        strSQL = "SELECT MAX(Pay_ID) AS MaxID FROM [" & sheetName & "$]"
    
        'Open recordset to find out the MaxID
        rs.Open strSQL, conn
        
        'Get the last used row
        Dim lastrow As Long
        If IsNull(rs.Fields("MaxID").value) Then
            ' Handle empty table (assign a default value, e.g., 1)
            lastrow = 1
        Else
            lastrow = rs.Fields("MaxID").value
        End If
        
        'Calculate the new ID
        Dim nextID As Long
        nextID = lastrow + 1
    
        'Close the previous recordset, no longer needed and open an empty recordset
        rs.Close
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
        
        'Add fields to the recordset before updating it
        rs.AddNew
            rs.Fields("Pay_ID").value = nextID
            rs.Fields("Inv_No").value = wshEncaissement.Range("F" & encRow).value
            rs.Fields("Customer").value = wshEncaissement.Range("F3").value
            rs.Fields("Pay_Date").value = CDate(wshEncaissement.Range("J3").value)
            rs.Fields("Pay_Amount").value = Format(wshEncaissement.Range("J" & encRow).value, "#,##0.00")
    Else 'Update an existing record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE TEC_ID=" & r, conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            rs.Fields("Inv_No").value = wshEncaissement.Range("F" & encRow).value
            rs.Fields("Customer").value = wshEncaissement.Range("F3").value
            rs.Fields("Pay_Date").value = CDate(wshEncaissement.Range("J3").value)
            rs.Fields("Amount").value = Format(wshEncaissement.Range("J5").value, "#,##0.00")
            rs.Fields("Notes").value = wshEncaissement.Range("F7").value
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le Pay_ID '" & r & "' ne peut être trouvé!", vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
    End If
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

