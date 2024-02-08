Attribute VB_Name = "Payments_Macros"
Option Explicit
Dim LastRow As Long, LastResultRow As Long
Dim PayRow As Long, PayCol As Long
Dim ResultRow As Long, PayItemRow As Long, LastPayItemRow As Long, PayItemDBRow As Long

Sub Payments_LoadOpenInvoices() '2024-02-07 @ 11:50
    wshEncaissement.Range("D13:K42").ClearContents 'Clear the invoices area before loading it
    With wshAR
        LastResultRow = .Range("A999999").End(xlUp).Row 'Last row
        If LastResultRow < 3 Then Exit Sub
        'Cells L3 contains a formula, no need to set it up
        .Range("A2:K" & LastResultRow).AdvancedFilter _
            xlFilterCopy, _
            criteriaRange:=.Range("L2:M3"), _
            copytorange:=.Range("P2:T2"), _
            Unique:=True
        LastResultRow = .Range("P999999").End(xlUp).Row
        If LastResultRow < 3 Then Exit Sub
        wshEncaissement.Range("B2").Value = True 'Set PaymentLoad to True
        .Range("S3:S" & LastResultRow).Formula = .Range("S1").Formula 'Total Payments Formula
        'Bring the Result data into our Payments List of Invoices
        wshEncaissement.Range("E13:I" & LastResultRow + 10).Value = .Range("P3:T" & LastResultRow).Value
    End With
    wshEncaissement.Range("B2").Value = False 'Set PaymentLoad to False
End Sub

Sub Payments_SaveUpdate() '2024-02-07 @ 12:27
    With wshEncaissement
        'Check for mandatory fields (4)
        If .Range("F3").Value = Empty Or _
           .Range("J3").Value = Empty Or _
           .Range("J3").Value = Empty Then
            MsgBox "Assurez-vous d'avoir..." & vbNewLine & vbNewLine & _
                "1. Un client" & vbNewLine & _
                "2. Une date de paiement" & vbNewLine & _
                "3. Un type de paiement et" & vbNewLine & _
                "4. Un montant de paiement" & vbNewLine & vbNewLine & _
                "AVANT de sauvegarder la transaction.", vbExclamation
            Exit Sub
        End If
        'Check to make sure Payment Amount = Applied Amount
        If .Range("J5").Value <> .Range("J10").Value Then
            MsgBox "Assurez-vous que le montant du paiement soit ÉGAL" & vbNewLine & _
                "à la somme des paiements appliqués", vbExclamation
            Exit Sub
        End If
        'New Payment -OR- Existing Payment ?
        If .Range("B4").Value = Empty Then 'New Payment
            PayRow = wshEncEntete.Range("A999999").End(xlUp).Row + 1 'First Available Row
            .Range("B3").Value = .Range("B5").Value 'Next payment ID
            wshEncEntete.Range("A" & PayRow).Value = .Range("B3").Value 'PayID
        Else 'Existing Payment
            PayRow = .Range("B4").Value
        End If
        'Using mapping (first row of the Payment List)
        For PayCol = 2 To 6
            wshEncEntete.Cells(PayRow, PayCol).Value = .Range(wshEncEntete.Cells(1, PayCol).Value).Value
        Next PayCol
        'Save Pay Items to Payment Items
        LastPayItemRow = .Range("E999999").End(xlUp).Row 'Last Pay Item
        For PayItemRow = 13 To LastPayItemRow
            If .Range("D" & PayItemRow).Value = Chr(252) Then 'The row has been applied
                If .Range("K" & PayItemRow).Value = Empty Then 'New Pay Item row
                    PayItemDBRow = wshEncDetail.Range("A999999").End(xlUp).Row + 1 'First Avail Pay Items Row
                    wshEncDetail.Range("A" & PayItemDBRow).Value = .Range("B3").Value 'Payment ID
                    wshEncDetail.Range("F" & PayItemDBRow).Value = "=row()"
                    .Range("K" & PayItemRow).Value = PayItemDBRow 'Database Row
                Else 'Existing Pay Item
                    PayItemDBRow = .Range("K" & PayItemRow).Value 'Existing Pay Item Row
                End If
                wshEncDetail.Range("B" & PayItemDBRow).Value = .Range("F" & PayItemRow).Value 'Invoice ID
                wshEncDetail.Range("C" & PayItemDBRow).Value = .Range("F3").Value 'Customer
                wshEncDetail.Range("D" & PayItemDBRow).Value = .Range("J3").Value 'Pay Date
                wshEncDetail.Range("E" & PayItemDBRow).Value = .Range("J" & PayItemRow).Value 'Amount paid
            End If
        Next PayItemRow
        MsgBox "Le paiement a été renregistré avec succès"
        Call Payments_AddNew 'Reset the form
        .Range("F3").Select
    End With
End Sub

Sub Payments_AddNew() '2024-02-07 @ 12:39
    wshEncaissement.Range("B3,F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").ClearContents 'Clear Fields
    wshEncaissement.Range("J3").Value = Date 'Set Default Date
    wshEncaissement.Range("F5").Value = "Banque" ' Set Default type
End Sub

Sub Payments_Load() '2024-02-07 @ 15:09
    With wshEncaissement
        If .Range("B4").Value = Empty Then
            MsgBox "Assurez vous de choisir un paiement valide", vbExclamation
            Exit Sub
        End If
        PayRow = .Range("B4").Value 'Payment Row
        .Range("B2").Value = True
        .Range("F3:G3,J3,F5:G5,J5,F7:J8,D13:K42").ClearContents
        'Using mapping (first row of the Payment List)
        For PayCol = 2 To 6
            .Range(wshEncEntete.Cells(1, PayCol).Value).Value = wshEncEntete.Cells(PayRow, PayCol).Value
        Next PayCol
        'Load Pay Items
        With wshEncDetail
            .Range("M4:T999999").ClearContents
            LastRow = .Range("A99999").End(xlUp).Row
            If LastRow < 4 Then GoTo NoData
            .Range("A3:G" & LastRow).AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=.Range("J2:J3"), _
                copytorange:=.Range("O3:T3"), _
                Unique:=True
            LastResultRow = .Range("O99999").End(xlUp).Row
            If LastResultRow < 4 Then GoTo NoData
            'Bring down the formulas into results
            .Range("M4:N" & LastResultRow).Formula = .Range("M1:N1").Formula 'Bring Apply and Invoice Date Formulas
            .Range("P4:R" & LastResultRow).Formula = .Range("P1:R1").Formula 'Inv. Amount, Prev. payments & Balance formulas
            wshEncaissement.Range("D13:K" & LastResultRow + 9).Value = .Range("M4:T" & LastResultRow).Value 'Bring over Pay Items
NoData:
        .Range("B2").Value = False 'Payment Load to False
        End With
    End With
End Sub

Sub Payments_Previous() '2024-02-07 @ 15:23
    Dim MinPayID As Long, PayID As Long
    With wshEncaissement
        On Error Resume Next
            MinPayID = Application.WorksheetFunction.Min(wshEncEntete.Range("Pay_ID"))
        On Error GoTo 0
        If MinPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
            Exit Sub
        End If
        PayID = .Range("B3").Value 'Payment ID
        If PayID = 0 Or .Range("B4").Value = Empty Then 'Load Last Payment Created
            PayRow = wshEncEntete.Range("A99999").End(xlUp).Row 'Last Row
        Else
            PayRow = .Range("B4").Value - 1 'Pay Row
        End If
        If PayRow = 3 Or MinPayID = .Range("B3").Value Then 'First Payment
            MsgBox "Vous êtes au premier paiement", vbExclamation
            Exit Sub
        End If
        .Range("B3").Value = wshEncEntete.Range("A" & PayRow).Value 'Set Payment ID
        Call Payments_Load 'Load Payment
    End With
End Sub

Sub Payments_Next() '2024-02-07 @ 15:30
    Dim MaxPayID As Long, PayID As Long
    With wshEncaissement
        On Error Resume Next
            MaxPayID = Application.WorksheetFunction.Max(wshEncEntete.Range("Pay_ID"))
        On Error GoTo 0
        If MaxPayID = 0 Then
            MsgBox "Vous devez avoir au minimum 1 paiement d'enregistré", vbExclamation
            Exit Sub
        End If
        PayID = .Range("B3").Value 'Payment ID
        If PayID = 0 Or .Range("B4").Value = Empty Then 'Load Last Payment Created
            PayRow = 4 'On new Payment, GOTO first one created
        Else
            PayRow = .Range("B4").Value + 1 'Pay Row
        End If
        If MaxPayID = PayID Then 'Last Payment
            MsgBox "Vous êtes au dernier paiement", vbExclamation
            Exit Sub
        End If
        .Range("B3").Value = wshEncEntete.Range("A" & PayRow).Value 'Set Payment ID
        Call Payments_Load 'Load Payment
    End With
End Sub

Sub Payments_Delete() '2024-02-07 @ 15:41
    If MsgBox("Êtes-vous certain de vouloir DÉTRUIRE ce paiement ? ", vbYesNo, _
        "Delete Payment") = vbNo Then Exit Sub
    With wshEncaissement
        If .Range("B4").Value = Empty Then GoTo NotSaved
        PayRow = .Range("B4").Value 'Pay Row
        wshEncEntete.Range(PayRow & ":" & PayRow).EntireRow.Delete 'Delete Payment Row

        With wshEncDetail
            LastRow = .Range("A99999").End(xlUp).Row
            If LastRow < 4 Then GoTo NotSaved
            .Range("A3:G" & LastRow).AdvancedFilter _
                xlFilterCopy, _
                criteriaRange:=.Range("J2:J3"), _
                copytorange:=.Range("O3:T3"), _
                Unique:=True
            LastResultRow = .Range("O99999").End(xlUp).Row
            If LastResultRow < 4 Then GoTo NotSaved
            If LastResultRow < 5 Then GoTo SkipSort
            With .Sort
                .SortFields.Clear
                .SortFields.Add _
                    Key:=wshEncDetail.Range("T4"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlDescending, _
                    DataOption:=xlSortNormal 'Sort
                .SetRange wshEncDetail.Range("O4:T" & LastResultRow) 'Set Range
                .Apply 'Apply Sort
            End With
SkipSort:
            On Error Resume Next
            For ResultRow = 4 To LastResultRow
                PayItemDBRow = .Range("T" & ResultRow).Value 'Pay Item DB Row
                .Range(PayItemDBRow & ":" & PayItemDBRow).EntireRow.Delete 'Delete Pay Item DB Row
            Next ResultRow
        End With
NotSaved:
        Call Payments_AddNew
    End With
End Sub


