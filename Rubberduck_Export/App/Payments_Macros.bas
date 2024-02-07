Attribute VB_Name = "Payments_Macros"
Option Explicit
Dim LastRow As Long, LastResultRow As Long, PayRow As Long, PayCol As Long
Dim ResultRow As Long, LastPayItemRow As Long, PayItemRow As Long, PayItemDBRow As Long

Sub Payments_LoadOpenInvoices() '2024-02-07 @ 11:50
    Payments.Range("D11:K35").ClearContents 'Clear the invoices before loading it
    With InvoiceList
        LastResultRow = .Range("A999999").End(xlUp).Row 'Last row
        If LastResultRow < 3 Then Exit Sub
        .Range("A2:K" & LastResultRow).AdvancedFilter _
        xlFilterCopy, _
        criteriaRange:=.Range("L2:M3"), _
        copytorange:=.Range("P2:T2"), _
        Unique:=True
        LastResultRow = .Range("P999999").End(xlUp).Row
        If LastResultRow < 3 Then Exit Sub
        Payments.Range("B2").Value = True 'Set PaymentLoad to True
        .Range("S3:S" & LastResultRow).Formula = .Range("S1").Formula 'Total Payments Formula
        'Bring the Result data into our Payments List of Invoices
        Payments.Range("E11:I" & LastResultRow + 8).Value = .Range("P3:T" & LastResultRow).Value
    End With
    Payments.Range("B2").Value = False 'Set PaymentLoad to False
End Sub

Sub Payments_SaveUpdate() '2024-02-07 @ 12:27
    With Payments
        'Check for required fields
        If .Range("F3").Value = Empty Or _
           .Range("J3").Value = Empty Or _
           .Range("J3").Value = Empty Then
            MsgBox "Please make sure to add in a Customer, payment date and Payment Amount before saving"
            Exit Sub
        End If
        'Check to make sure Payment Amount = Applied Amount
        If .Range("J5").Value <> .Range("J9").Value Then
            MsgBox "Please make sure Payment Amount is equal to Applied Amount"
            Exit Sub
        End If
        'New Payment -OR- Existing Payment ?
        If .Range("B4").Value = Empty Then 'New Payment
            PayRow = PaymentList.Range("A999999").End(xlUp).Row + 1 'First Available Row
            .Range("B3").Value = .Range("B5").Value 'Next payment ID
            PaymentList.Range("A" & PayRow).Value = .Range("B3").Value 'PayID
        Else 'Existing Payment
            PayRow = .Range("B4").Value
        End If
        'Using mapping (first row of the Payment List)
        For PayCol = 2 To 6
            PaymentList.Cells(PayRow, PayCol).Value = .Range(PaymentList.Cells(1, PayCol).Value).Value
        Next PayCol
        'Save Pay Items to Payment Items
        LastPayItemRow = .Range("E999999").End(xlUp).Row 'Last Pay Item
        For PayItemRow = 11 To LastPayItemRow
            If .Range("D" & PayItemRow).Value = Chr(252) Then 'The row has been applied
                If .Range("K" & PayItemRow).Value = Empty Then 'New Pay Item row
                    PayItemDBRow = PayItems.Range("A999999").End(xlUp).Row + 1 'First Avail Pay Items Row
                    PayItems.Range("A" & PayItemDBRow).Value = .Range("B3").Value 'Payment ID
                    PayItems.Range("F" & PayItemDBRow).Value = "=row()"
                    .Range("K" & PayItemRow).Value = PayItemDBRow 'Database Row
                Else 'Existing Pay Item
                    PayItemDBRow = .Range("K" & PayItemRow).Value 'Existing Pay Item Row
                End If
                PayItems.Range("B" & PayItemDBRow).Value = .Range("F" & PayItemRow).Value 'Invoice ID
                PayItems.Range("C" & PayItemDBRow).Value = .Range("F3").Value 'Customer
                PayItems.Range("D" & PayItemDBRow).Value = .Range("J3").Value 'Pay Date
                PayItems.Range("E" & PayItemDBRow).Value = .Range("J" & PayItemRow).Value 'Amount paid
            End If
        Next PayItemRow
    End With
End Sub

Sub Payments_AddNew() '2024-02-07 @ 12:39
    Payments.Range("B3,F3:G3,J3,F5:G5,J5,F7:J8,D11:K35").ClearContents 'Clear Fields
    Payments.Range("J3").Value = Date 'Set Default Date
End Sub

Sub Payments_Load() '2024-02-07 @ 15:09
    With Payments
        If .Range("B4").Value = Empty Then
            MsgBox "Please make sure to select a correct payment"
            Exit Sub
        End If
        PayRow = .Range("B4").Value 'Payment Row
        .Range("B2").Value = True
        .Range("F3:G3,J3,F5:G5,J5,F7:J8,D11:K35").ClearContents
        'Using mapping (first row of the Payment List)
        For PayCol = 2 To 6
            .Range(PaymentList.Cells(1, PayCol).Value).Value = PaymentList.Cells(PayRow, PayCol).Value
        Next PayCol
        'Load Pay Items
        With PayItems
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
            Payments.Range("D11:K" & LastResultRow + 7).Value = .Range("M4:T" & LastResultRow).Value 'Bring over Pay Items
NoData:
        .Range("B2").Value = False 'Payment Load to False
        End With
    End With
End Sub

Sub Payments_Previous() '2024-02-07 @ 15:23
    Dim MinPayID As Long, PayID As Long
    With Payments
        On Error Resume Next
            MinPayID = Application.WorksheetFunction.Min(PaymentList.Range("Pay_ID"))
        On Error GoTo 0
        If MinPayID = 0 Then
            MsgBox "Please create a payment first"
            Exit Sub
        End If
        PayID = .Range("B3").Value 'Payment ID
        If PayID = 0 Or .Range("B4").Value = Empty Then 'Load Last Payment Created
            PayRow = PaymentList.Range("A99999").End(xlUp).Row 'Last Row
        Else
            PayRow = .Range("B4").Value - 1 'Pay Row
        End If
        If PayRow = 3 Or MinPayID = .Range("B3").Value Then 'First Payment
            MsgBox "You are on the first payment"
            Exit Sub
        End If
        .Range("B3").Value = PaymentList.Range("A" & PayRow).Value 'Set Payment ID
        Call Payments_Load 'Load Payment
    End With
End Sub

Sub Payments_Next() '2024-02-07 @ 15:30
    Dim MaxPayID As Long, PayID As Long
    With Payments
        On Error Resume Next
            MaxPayID = Application.WorksheetFunction.Max(PaymentList.Range("Pay_ID"))
        On Error GoTo 0
        If MaxPayID = 0 Then
            MsgBox "Please create a payment first"
            Exit Sub
        End If
        PayID = .Range("B3").Value 'Payment ID
        If PayID = 0 Or .Range("B4").Value = Empty Then 'Load Last Payment Created
            PayRow = 4 'On new Payment, GOTO first one created
        Else
            PayRow = .Range("B4").Value + 1 'Pay Row
        End If
        If MaxPayID = PayID Then 'Last Payment
            MsgBox "You are on the last payment"
            Exit Sub
        End If
        .Range("B3").Value = PaymentList.Range("A" & PayRow).Value 'Set Payment ID
        Call Payments_Load 'Load Payment
    End With
End Sub

Sub Payments_Delete() '2024-02-07 @ 15:41
    If MsgBox("Are you sure you want to delete this payment ? ", vbYesNo, _
        "Delete Payment") = vbNo Then Exit Sub
    With Payments
        If .Range("B4").Value = Empty Then GoTo NotSaved
        PayRow = .Range("B4").Value 'Pay Row
        PaymentList.Range(PayRow & ":" & PayRow).EntireRow.Delete 'Delete Payment Row

        With PayItems
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
                    Key:=PayItems.Range("T4"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlDescending, _
                    DataOption:=xlSortNormal 'Sort
                .SetRange PayItems.Range("O4:T" & LastResultRow) 'Set Range
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


