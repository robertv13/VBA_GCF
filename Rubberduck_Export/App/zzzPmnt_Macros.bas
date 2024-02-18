Attribute VB_Name = "zzzPmnt_Macros"
Option Explicit
Dim PmntRow As Long, PmntID As Long, LastRow As Long, LastResultRow As Long, InvRow As Long

Sub Payment_CustomerPmntsRefresh()
    Payments.Range("D13:H999").ClearContents     'Clear existing Data
    With PmntsDB
        LastRow = .Range("A99999").End(xlUp).row
        If LastRow < 3 Then Exit Sub
        .Range("A2:F" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("K1:K2"), CopyToRange:=.Range("M2:P2"), Unique:=True
        LastResultRow = .Range("M99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        If LastResultRow < 4 Then GoTo SkipSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=PmntsDB.Range("M3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange PmntsDB.Range("M3:P" & LastResultRow)
            .Apply
        End With
SkipSort:
        Payments.Range("D13:G" & LastResultRow + 10).value = .Range("M3:P" & LastResultRow).value
    End With
End Sub

Sub Payment_SaveUpdate()
    With Payments
        If .Range("B5").value = Empty Then       'Incorrect / Missing Invoice #
            MsgBox "Please select a correct Invoice #"
            Exit Sub
        End If
        InvRow = .Range("B5").value              'Invoice Row
        If .Range("H3").value = Empty Or .Range("H7").value = Empty Then 'Empty Fields
            MsgBox "Please make sure to add in a Payment Date and Payment Amount"
            Exit Sub
        End If
    
        If .Range("H8").value < 0 Then
            If MsgBox("The Payment Amount is above the Invoice Balance. Are you sure you want to continue?", vbYesNo, "Payment Amount Issue") = vbNo Then Exit Sub
        End If
    
        If .Range("B3").value = Empty Then       'New Payment
            PmntRow = PmntsDB.Range("A99999").End(xlUp).row + 1
            .Range("B2").value = .Range("B4").value 'Next Pment ID
            PmntsDB.Range("A" & PmntRow).value = .Range("B4").value 'Next Payment ID
        Else                                     'Existing Payment
            PmntRow = .Range("B3").value         'Payment Row
        End If
        PmntsDB.Range("B" & PmntRow).value = .Range("H3").value 'Date
        PmntsDB.Range("C" & PmntRow).value = .Range("D5").value 'Customer
        PmntsDB.Range("D" & PmntRow).value = .Range("E3").value 'Invoice #
        PmntsDB.Range("E" & PmntRow).value = .Range("H7").value 'Amount
        PmntsDB.Range("F" & PmntRow).value = .Range("E9").value 'Notes
    
        'Update invoice Paid Status
        If .Range("H8").value = 0 Then           'Paid
            wshInvoiceList.Range("D" & InvRow).value = Admin.Range("C10").value 'Fully Paid
        Else                                     ' Partial Paid
            wshInvoiceList.Range("D" & InvRow).value = Admin.Range("C9").value 'Partially Paid
        End If
    
    End With
    Payment_CustomerPmntsRefresh                 'Refresh Customer Payments
    Payment_SavedMsg                             'Run Fade out message
End Sub

Sub Payment_SavedMsg()
    With Payments.Shapes("PmntSavedMsg")
        Dim i As Long, Delay As Double, StartTime As Double
        .Visible = msoCTrue
        For i = 1 To 150
            .Fill.Transparency = i / 150
            Delay = 0.009
            StartTime = Timer
            Do
                DoEvents
            Loop While Timer - StartTime < Delay
        Next i
        .Visible = msoFalse
    End With
End Sub

Sub Payment_New()
    Payments.Range("B2,E3,H7,D5:D7,E9:H9,D13:H999").ClearContents
    Payments.Range("E3").Select
End Sub

Sub Payment_Load()
    With Payments
        .Range("E3,H7,D5:D7,E9:H9,D13:H999").ClearContents
        If .Range("B3").value = Empty Then
            MsgBox "Please select a correct payment"
            Exit Sub
        End If
        PmntRow = .Range("B3").value             'Payment Row
        .Range("H3").value = PmntsDB.Range("B" & PmntRow).value 'Date
        .Range("D5").value = PmntsDB.Range("C" & PmntRow).value 'Customer
        .Range("E3").value = PmntsDB.Range("D" & PmntRow).value 'Invoice #
        .Range("H7").value = PmntsDB.Range("E" & PmntRow).value 'Amount
        .Range("E9").value = PmntsDB.Range("F" & PmntRow).value 'Notes
        If .Range("B6").value <> "" Then Payment_CustomerPmntsRefresh 'Load Previous Customer Payments
    End With
End Sub

Sub Payment_Prev()
    With Payments
        PmntID = .Range("B2").value              'Payment ID
        If PmntID = 0 Then                       'No Current ID
            If .Range("B4").value = 1 Then       'No Saved Payments
                MsgBox "Please save any Payments first before navigating to previously saved"
                Exit Sub
            End If
            .Range("B2").value = .Range("B4").value - 1 'Set Pmnt. ID to the last one created
            Payment_Load
            Exit Sub
        End If
        If PmntID = 1 Then
            MsgBox "You are are already at the first Payment created"
            Exit Sub
        End If
        .Range("B2").value = .Range("B2").value - 1 'Set Previous Payment ID
        Payment_Load
    End With
End Sub

Sub Payment_Next()
    With Payments
        PmntID = .Range("B2").value              'Payment ID
        If PmntID = 0 Then                       'No Current ID
            If .Range("B4").value = 1 Then       'No Saved Payments
                MsgBox "Please save any Payments first before navigating to previously saved"
                Exit Sub
            End If
            .Range("B2").value = 1               'Set Pmnt. ID to the first one created
            Payment_Load
            Exit Sub
        End If
        If PmntID = .Range("B4").value - 1 Then
            MsgBox "You are are already at the last Payment created"
            Exit Sub
        End If
        .Range("B2").value = .Range("B2").value + 1 'Set next Payment ID
        Payment_Load
    End With
End Sub

Sub Payment_Delete()
    If MsgBox("Are you sure you want to delete this Payment?", vbYesNo, "Delete Payment") = vbNo Then Exit Sub
    With Payments
        If .Range("B3").value = Empty Then GoTo NotSaved
        PmntRow = .Range("B3").value             'Payment Row
        PmntsDB.Range(PmntRow & ":" & PmntRow).EntireRow.Delete 'Delete Payment Row
NotSaved:
        Payment_New
    End With
End Sub

