Attribute VB_Name = "modEncaissement"
Option Explicit
Dim lastRow As Long, LastResultRow As Long
Dim PayRow As Long, PayCol As Long
Dim ResultRow As Long, PayItemRow As Long, LastPayItemRow As Long, PayItemDBRow As Long

Sub Encaissements_Load_Open_Invoices() '2024-02-14 @ 10:53
    wshEncaissement.Range("D13:K42").ClearContents 'Clear the invoices area before loading it
    With wshAR
        LastResultRow = .Range("A999999").End(xlUp).row 'Last row
        If LastResultRow < 3 Then Exit Sub
        'Cells L3 contains a formula, no need to set it up
        .Range("A2:I" & LastResultRow).AdvancedFilter _
            xlFilterCopy, _
            criteriaRange:=.Range("J2:K3"), _
            copytorange:=.Range("M2:Q2"), _
            Unique:=True
        LastResultRow = .Range("M999999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        wshEncaissement.Range("B2").value = True 'Set PaymentLoad to True
        .Range("P3:P" & LastResultRow).formula = .Range("P1").formula 'Total Payments Formula
        'Bring the Result data into our Payments List of Invoices
        wshEncaissement.Range("E13:I" & LastResultRow + 10).value = .Range("M3:Q" & LastResultRow).value
    End With
    wshEncaissement.Range("B2").value = False 'Set PaymentLoad to False
End Sub

Sub Encaissements_Save_Update() '2024-02-07 @ 12:27
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
            wshEncEntete.Range("A" & PayRow).value = .Range("B3").value 'PayID
        Else 'Existing Payment
            PayRow = .Range("B4").value
        End If
        'Using mapping (first row of the Payment List)
        For PayCol = 2 To 6
            wshEncEntete.Cells(PayRow, PayCol).value = .Range(wshEncEntete.Cells(1, PayCol).value).value
        Next PayCol
        
        'Save Pay Items to Payment Items
        LastPayItemRow = .Range("E999999").End(xlUp).row 'Last Pay Item
        For PayItemRow = 13 To LastPayItemRow
            If .Range("D" & PayItemRow).value = Chr(252) Then 'The row has been applied
                If .Range("K" & PayItemRow).value = Empty Then 'New Pay Item row
                    PayItemDBRow = wshEncDetail.Range("A999999").End(xlUp).row + 1 'First Avail Pay Items Row
                    wshEncDetail.Range("A" & PayItemDBRow).value = .Range("B3").value 'Payment ID
                    wshEncDetail.Range("F" & PayItemDBRow).value = "=row()"
                    .Range("K" & PayItemRow).value = PayItemDBRow 'Database Row
                Else 'Existing Pay Item
                    PayItemDBRow = .Range("K" & PayItemRow).value 'Existing Pay Item Row
                End If
                wshEncDetail.Range("B" & PayItemDBRow).value = .Range("F" & PayItemRow).value 'Invoice ID
                wshEncDetail.Range("C" & PayItemDBRow).value = .Range("F3").value 'Customer
                wshEncDetail.Range("D" & PayItemDBRow).value = .Range("J3").value 'Pay Date
                wshEncDetail.Range("E" & PayItemDBRow).value = .Range("J" & PayItemRow).value 'Amount paid
            End If
        Next PayItemRow
        
        'Prepare G/L posting
        Dim noEnc As String, nomCLient As String, typeEnc As String, descEnc As String
        Dim dateEnc As Date
        Dim montantEnc As Currency
        noEnc = wshEncEntete.Cells(PayRow, 1).value
        dateEnc = wshEncEntete.Cells(PayRow, 2).value
        nomCLient = wshEncEntete.Cells(PayRow, 3).value
        typeEnc = wshEncEntete.Cells(PayRow, 4).value
        montantEnc = wshEncEntete.Cells(PayRow, 5).value
        descEnc = wshEncEntete.Cells(PayRow, 6).value

        Call Encaissement_GL_Posting(noEnc, dateEnc, nomCLient, typeEnc, montantEnc, descEnc)  '2024-02-09 @ 08:17 - TODO
        
        MsgBox "Le paiement a été renregistré avec succès"
        Call Encaissements_Add_New 'Reset the form
        .Range("F3").Select
    End With
End Sub

Sub Encaissements_Add_New() '2024-02-07 @ 12:39
    wshEncaissement.Range("B2").value = False
    wshEncaissement.Range("B3,F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").ClearContents 'Clear Fields
    wshEncaissement.Range("J3").value = Date 'Set Default Date
    wshEncaissement.Range("F5").value = "Banque" ' Set Default type
End Sub

Sub Encaissements_Load() '2024-02-14 @ 11:04
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
            lastRow = .Range("A999999").End(xlUp).row
            If lastRow < 4 Then GoTo NoData
            .Range("A3:G" & lastRow).AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=.Range("J2:J3"), _
                copytorange:=.Range("O3:T3"), _
                Unique:=True
            LastResultRow = .Range("O99999").End(xlUp).row
            If LastResultRow < 4 Then GoTo NoData
            'Bring down the formulas into results
            .Range("M4:N" & LastResultRow).formula = .Range("M1:N1").formula 'Bring Apply and Invoice Date Formulas
            .Range("P4:R" & LastResultRow).formula = .Range("P1:R1").formula 'Inv. Amount, Prev. payments & Balance formulas
            wshEncaissement.Range("D13:K" & LastResultRow + 9).value = .Range("M4:T" & LastResultRow).value 'Bring over Pay Items
NoData:
        .Range("B2").value = False 'Payment Load to False
        End With
    End With
End Sub

Sub Encaissements_Previous() '2024-02-14 @ 11:04
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
        Call Encaissements_Load 'Load Payment
    End With
End Sub

Sub Encaissements_Next() '2024-02-14 @ 11:04
    Dim MaxPayID As Long, PayID As Long
    With wshEncaissement
        On Error Resume Next
            MaxPayID = Application.WorksheetFunction.Max(wshEncEntete.Range("Pay_ID"))
        On Error GoTo 0
        If MaxPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
            Exit Sub
        End If
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
        .Range("B3").value = wshEncEntete.Range("A" & PayRow).value 'Set Payment ID
        Call Encaissements_Load 'Load Payment
    End With
End Sub

'Sub Payments_Delete() '2024-02-07 @ 15:41
'    If MsgBox("Êtes-vous certain de vouloir DÉTRUIRE ce paiement ? ", vbYesNo, _
'        "Delete Payment") = vbNo Then Exit Sub
'    With wshEncaissement
'        If .Range("B4").value = Empty Then GoTo NotSaved
'        PayRow = .Range("B4").value 'Pay Row
'        wshEncEntete.Range(PayRow & ":" & PayRow).EntireRow.Delete 'Delete Payment Row
'
'        With wshEncDetail
'            lastRow = .Range("A99999").End(xlUp).row
'            If lastRow < 4 Then GoTo NotSaved
'            .Range("A3:G" & lastRow).AdvancedFilter _
'                xlFilterCopy, _
'                criteriaRange:=.Range("J2:J3"), _
'                copytorange:=.Range("O3:T3"), _
'                Unique:=True
'            LastResultRow = .Range("O99999").End(xlUp).row
'            If LastResultRow < 4 Then GoTo NotSaved
'            If LastResultRow < 5 Then GoTo SkipSort
'            With .Sort
'                .SortFields.Clear
'                .SortFields.Add _
'                    Key:=wshEncDetail.Range("T4"), _
'                    SortOn:=xlSortOnValues, _
'                    Order:=xlDescending, _
'                    DataOption:=xlSortNormal 'Sort
'                .SetRange wshEncDetail.Range("O4:T" & LastResultRow) 'Set Range
'                .Apply 'Apply Sort
'            End With
'SkipSort:
'            On Error Resume Next
'            For ResultRow = 4 To LastResultRow
'                PayItemDBRow = .Range("T" & ResultRow).value 'Pay Item DB Row
'                .Range(PayItemDBRow & ":" & PayItemDBRow).EntireRow.Delete 'Delete Pay Item DB Row
'            Next ResultRow
'        End With
'NotSaved:
'        Call Encaissements_Add_New
'    End With
'End Sub

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
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).UsedRange
    Set destinationRange = wshAR.Range("A2")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshAR.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Arrange formats on all rows
    Dim lastRow As Long
    lastRow = wshAR.Range("A999999").End(xlUp).row
    
    With wshAR
        .Range("A3" & ":F" & lastRow).HorizontalAlignment = xlCenter
        With .Range("C3:C" & lastRow & ",D3:D" & lastRow & ",E3:E" & lastRow)
            .HorizontalAlignment = xlLeft
        End With
        .Range("G3:G" & lastRow & ",H3:H" & lastRow).HorizontalAlignment = xlRight
        .Range("G3:H" & lastRow).NumberFormat = "#,##0.00 $"
        .Range("B3:B" & lastRow & ",F3:F" & lastRow).NumberFormat = "dd/mm/yyyy"
    End With
    
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
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).UsedRange
    Set destinationRange = wshEncEntete.Range("A2")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshEncEntete.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Arrange formats on all rows
    Dim lastRow As Long
    lastRow = wshEncEntete.Range("A999999").End(xlUp).row
    
    With wshEncEntete
        .Range("A4" & ":B" & lastRow).HorizontalAlignment = xlCenter
        With .Range("C4:C" & lastRow & ",D4:D" & lastRow & ",F4:F" & lastRow)
            .HorizontalAlignment = xlLeft
        End With
        .Range("E3:E" & lastRow).HorizontalAlignment = xlRight
        .Range("G3:H" & lastRow).NumberFormat = "#,##0.00 $"
        .Range("B3:B" & lastRow).NumberFormat = "dd/mm/yyyy"
        .Range("F3:F" & lastRow).NumberFormat = "dd/mm/yyyy"
    End With
    
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
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).UsedRange
    Set destinationRange = wshEncDetail.Range("A3")

    'Copy data, using Range to Range and Autofit all columns
    sourceRange.Copy destinationRange
    wshEncDetail.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Close the source workbook, without saving it
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    'Arrange formats on all rows
    Dim lastRow As Long
    lastRow = wshEncDetail.Range("A999999").End(xlUp).row
    
    With wshEncDetail
        .Range("A4:B" & lastRow & ",D4:D" & lastRow & ",F4:F" & lastRow).HorizontalAlignment = xlCenter
        .Range("C4:C" & lastRow).HorizontalAlignment = xlLeft
        .Range("D3:D" & lastRow).NumberFormat = "dd/mm/yyyy"
        .Range("E3:E" & lastRow).HorizontalAlignment = xlRight
        .Range("E3:E" & lastRow).NumberFormat = "#,##0.00 $"
    End With
    
End Sub


