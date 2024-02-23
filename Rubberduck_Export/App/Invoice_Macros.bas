Attribute VB_Name = "Invoice_Macros"
Option Explicit
Dim InvRow As Long, InvItemRow As Long, lastrow As Long, LastItemRow As Long
Dim resultRow As Long, LastResultRow As Long, ItemRow As Long, TermRow As Long, StatusRow As Long

Sub Dashboard_Invoice_SaveUpdate()
    With Invoice
        If .Range("G5").value = Empty Then
            MsgBox "Please make sure to add a customer before saving invoice"
            Exit Sub
        End If
        
        'Deterermine New Invoice/Existing Invoice
        If .Range("B3").value = Empty Then       'new Invoice
            InvRow = wshInvoiceList.Range("A99999").End(xlUp).row + 1 'First Avail Row
            .Range("J1").value = .Range("B5").value 'Next Inv. #
            wshInvoiceList.Range("A" & InvRow).value = .Range("B5").value 'Next Inv. #
        Else                                     'Existing Invoice
            InvRow = .Range("B3").value          'Invoice Row
        End If
        wshInvoiceList.Range("B" & InvRow).value = .Range("I3").value 'Date
        wshInvoiceList.Range("C" & InvRow).value = .Range("G5").value 'Customer
        wshInvoiceList.Range("D" & InvRow).value = .Range("I4").value 'Status
        wshInvoiceList.Range("E" & InvRow).value = .Range("I5").value 'Terms
        wshInvoiceList.Range("F" & InvRow).value = .Range("I6").value 'Due Date
        wshInvoiceList.Range("G" & InvRow).value = .Range("J34").value 'Invoice Total
            
            
        'Add/Update Invoice Items
        LastItemRow = .Range("C31").End(xlUp).row 'Last Invoice Row
        If LastItemRow < 9 Then GoTo NoItems
        For ItemRow = 9 To LastItemRow
            If .Range("B" & ItemRow).value <> Empty Then 'DB Row Exists
                InvItemRow = .Range("B" & ItemRow).value 'Inv. Item DB Row
            Else                                 'New DB ItemRow
                InvItemRow = InvoiceItems.Range("A99999").End(xlUp).row + 1 'First avail Row
                InvoiceItems.Range("A" & InvItemRow).value = .Range("J1").value 'Inv ID
                .Range("B" & ItemRow).value = InvItemRow 'Add DB Row
                InvoiceItems.Range("K" & InvItemRow).value = "=Row()" 'Add Row Formula
            End If
            InvoiceItems.Range("B" & InvItemRow & ":H" & InvItemRow).value = .Range("C" & ItemRow & ":I" & ItemRow).value ' Bring over item Values
            InvoiceItems.Range("I" & InvItemRow).value = .Range("K" & ItemRow).value 'Item Cost
            InvoiceItems.Range("J" & InvItemRow).value = ItemRow 'Invoice Row
        Next ItemRow
    
NoItems:
    End With
    Invoice_SavedMsg                             'Run Invoice Fade Out Message
End Sub

Sub Invoice_SavedMsg()
    With Invoice.Shapes("InvSavedMsg")
        Dim i As Long, Delay As Double, StartTime As Double
        .Visible = msoCTrue
        For i = 1 To 150
            .fill.Transparency = i / 150
            Delay = 0.009
            StartTime = Timer
            Do
                DoEvents
            Loop While Timer - StartTime < Delay
        Next i
        .Visible = msoFalse
    End With
End Sub

Sub Customer_AddNew()
    Unload AddCustForm
    AddCustForm.show
End Sub

Sub Dashboard_Invoice_New()
    With Invoice
        .Range("B6").value = True                'Set Inv. Load to true
        .Range("I3:J6,G5:G7,B9:I31,K9:K31").ClearContents
        .Range("J1").value = .Range("B5").value  'Set Next invoice #
        .Range("I3").value = Date                'Set current Date
        .Range("B6").value = False               'Set inv. Load to false
        On Error Resume Next
        TermRow = Admin.Range("H6:H23").Find(Chr(252), , xlValues, xlWhole).row
        On Error GoTo 0
        If TermRow <> 0 Then .Range("I5").value = Admin.Range("F" & TermRow).value 'Set Default Term
        On Error Resume Next
        StatusRow = Admin.Range("D6:D12").Find(Chr(252), , xlValues, xlWhole).row
        On Error GoTo 0
        If StatusRow <> 0 Then .Range("I4").value = Admin.Range("C" & StatusRow).value 'Set Default Status
    End With
End Sub

Sub Dashboard_Invoice_Load()
    With Invoice
        If .Range("B3").value = Empty Then
            MsgBox "Please entere a correct invoice #"
            Exit Sub
        End If
        InvRow = .Range("B3").value              'Invoice Row
        .Range("B6").value = True                'Set Inv. Load to true
        .Range("I3:J6,G5:G7,B9:I31,K9:K31").ClearContents
        .Range("I3").value = wshInvoiceList.Range("B" & InvRow).value 'Inv. Date
        .Range("G5").value = wshInvoiceList.Range("C" & InvRow).value 'Customer
        .Range("I4").value = wshInvoiceList.Range("D" & InvRow).value 'Inv. Status
        .Range("I5").value = wshInvoiceList.Range("E" & InvRow).value 'Terms
        .Range("I6").value = wshInvoiceList.Range("F" & InvRow).value 'Due Date
    
        With InvoiceItems
            lastrow = .Range("A99999").End(xlUp).row
            If lastrow < 3 Then GoTo NoItems
            .Range("A2:K" & lastrow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("M2:M3"), CopyToRange:=.Range("P2:Y2"), Unique:=True
            LastResultRow = .Range("P99999").End(xlUp).row
            If LastResultRow < 3 Then GoTo NoItems
            For resultRow = 3 To LastResultRow
                InvItemRow = .Range("Y" & resultRow).value 'Get Invoice Row
                Invoice.Range("B" & InvItemRow & ":I" & InvItemRow).value = .Range("P" & resultRow & ":W" & resultRow).value 'Item Details
                Invoice.Range("K" & InvItemRow).value = InvoiceItems.Range("X" & resultRow).value 'Item Cost
            Next resultRow
        End With
NoItems:
        .Range("B6").value = False               'Set inv. Load to false
    End With
End Sub

Sub Invoice_Print()
    Invoice.PrintOut , , , , True, , , , False
End Sub

Sub Invoice_SaveAsPDF()
    Dim FilePath As String
    Dashboard_Invoice_SaveUpdate                           'Save invoice
    FilePath = ThisWorkbook.Path & "\" & Invoice.Range("G5").value & "_" & Invoice.Range("J1").value 'File Path
    If Dir(FilePath, vbDirectory) <> "" Then Kill (FilePath)
    Invoice.ExportAsFixedFormat xlTypePDF, FilePath, , , False, , , True
End Sub

Sub Invoice_Delete()
    If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Appointment") = vbNo Then Exit Sub
    With Invoice
        If .Range("B3").value = Empty Then GoTo NotSaved
        InvRow = .Range("B3").value              'Order Row
        wshInvoiceList.Range(InvRow & ":" & InvRow).EntireRow.Delete
        'Remove Invoice Items
        With InvoiceItems
            lastrow = .Range("A99999").End(xlUp).row
            If lastrow < 3 Then GoTo NotSaved
            .Range("A2:K" & lastrow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("M2:M3"), CopyToRange:=.Range("P2:Y2"), Unique:=True
            LastResultRow = .Range("P99999").End(xlUp).row
            If LastResultRow < 3 Then GoTo NotSaved
            If LastResultRow < 4 Then GoTo SingleRow
            'Sort based on descending rows
            With .Sort
                .SortFields.Clear
                .SortFields.Add Key:=InvoiceItems.Range("P3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange InvoiceItems.Range("P3:Y" & LastResultRow)
                .Apply
            End With
SingleRow:
            For resultRow = 3 To LastResultRow
                InvItemRow = .Range("P" & resultRow).value 'Invoice Item DB Row
                If InvItemRow > 3 Then .Range(InvItemRow & ":" & InvItemRow).EntireRow.Delete 'Don't remove 1st Row
            Next resultRow
        End With
NotSaved:
        Dashboard_Invoice_New                              'Clear Out All Invoice Fields
    End With
End Sub

