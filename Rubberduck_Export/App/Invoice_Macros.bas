Attribute VB_Name = "Invoice_Macros"
Option Explicit
Dim invRow As Long, invitemRow As Long, lastRow As Long, lastItemRow As Long
Dim resultRow As Long, lastResultRow As Long, itemRow As Long, termRow As Long, statusRow As Long

Sub Dashboard_Invoice_SaveUpdate()
    With Invoice
        If .Range("G5").value = Empty Then
            MsgBox "Please make sure to add a customer before saving invoice"
            Exit Sub
        End If
        
        'Deterermine New Invoice/Existing Invoice
        If .Range("B3").value = Empty Then       'new Invoice
            invRow = wshCC_Invoice_List.Range("A99999").End(xlUp).row + 1 'First Avail Row
            .Range("J1").value = .Range("B5").value 'Next Inv. #
            wshCC_Invoice_List.Range("A" & invRow).value = .Range("B5").value 'Next Inv. #
        Else                                     'Existing Invoice
            invRow = .Range("B3").value          'Invoice Row
        End If
        wshCC_Invoice_List.Range("B" & invRow).value = .Range("I3").value 'Date
        wshCC_Invoice_List.Range("C" & invRow).value = .Range("G5").value 'Customer
        wshCC_Invoice_List.Range("D" & invRow).value = .Range("I4").value 'Status
        wshCC_Invoice_List.Range("E" & invRow).value = .Range("I5").value 'Terms
        wshCC_Invoice_List.Range("F" & invRow).value = .Range("I6").value 'Due Date
        wshCC_Invoice_List.Range("G" & invRow).value = .Range("J34").value 'Invoice Total
            
            
        'Add/Update Invoice Items
        lastItemRow = .Range("C31").End(xlUp).row 'Last Invoice Row
        If lastItemRow < 9 Then GoTo NoItems
        For itemRow = 9 To lastItemRow
            If .Range("B" & itemRow).value <> Empty Then 'DB Row Exists
                invitemRow = .Range("B" & itemRow).value 'Inv. Item DB Row
            Else                                 'New DB itemRow
                invitemRow = InvoiceItems.Range("A99999").End(xlUp).row + 1 'First avail Row
                InvoiceItems.Range("A" & invitemRow).value = .Range("J1").value 'Inv ID
                .Range("B" & itemRow).value = invitemRow 'Add DB Row
                InvoiceItems.Range("K" & invitemRow).value = "=Row()" 'Add Row Formula
            End If
            InvoiceItems.Range("B" & invitemRow & ":H" & invitemRow).value = .Range("C" & itemRow & ":I" & itemRow).value ' Bring over item Values
            InvoiceItems.Range("I" & invitemRow).value = .Range("K" & itemRow).value 'Item Cost
            InvoiceItems.Range("J" & invitemRow).value = itemRow 'Invoice Row
        Next itemRow
    
NoItems:
    End With
    Invoice_SavedMsg                             'Run Invoice Fade Out Message
End Sub

Sub Invoice_SavedMsg()
    With Invoice.Shapes("InvSavedMsg")
        Dim i As Long, Delay As Double, startTime As Double
        .Visible = msoCTrue
        For i = 1 To 150
            .fill.Transparency = i / 150
            Delay = 0.009
            startTime = Timer
            Do
                DoEvents
            Loop While Timer - startTime < Delay
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
        .Range("I3:J6,G5:G7,B9:I31,K9:K31").Clearcontents
        .Range("J1").value = .Range("B5").value  'Set Next invoice #
        .Range("I3").value = Date                'Set current Date
        .Range("B6").value = False               'Set inv. Load to false
        On Error Resume Next
        termRow = Admin.Range("H6:H23").Find(Chr(252), , xlValues, xlWhole).row
        On Error GoTo 0
        If termRow <> 0 Then .Range("I5").value = Admin.Range("F" & termRow).value 'Set Default Term
        On Error Resume Next
        statusRow = Admin.Range("D6:D12").Find(Chr(252), , xlValues, xlWhole).row
        On Error GoTo 0
        If statusRow <> 0 Then .Range("I4").value = Admin.Range("C" & statusRow).value 'Set Default Status
    End With
End Sub

Sub Dashboard_Invoice_Load()
    With Invoice
        If .Range("B3").value = Empty Then
            MsgBox "Please entere a correct invoice #"
            Exit Sub
        End If
        invRow = .Range("B3").value              'Invoice Row
        .Range("B6").value = True                'Set Inv. Load to true
        .Range("I3:J6,G5:G7,B9:I31,K9:K31").Clearcontents
        .Range("I3").value = wshCC_Invoice_List.Range("B" & invRow).value 'Inv. Date
        .Range("G5").value = wshCC_Invoice_List.Range("C" & invRow).value 'Customer
        .Range("I4").value = wshCC_Invoice_List.Range("D" & invRow).value 'Inv. Status
        .Range("I5").value = wshCC_Invoice_List.Range("E" & invRow).value 'Terms
        .Range("I6").value = wshCC_Invoice_List.Range("F" & invRow).value 'Due Date
    
        With InvoiceItems
            lastRow = .Range("A99999").End(xlUp).row
            If lastRow < 3 Then GoTo NoItems
            .Range("A2:K" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("M2:M3"), CopyToRange:=.Range("P2:Y2"), Unique:=True
            lastResultRow = .Range("P99999").End(xlUp).row
            If lastResultRow < 3 Then GoTo NoItems
            For resultRow = 3 To lastResultRow
                invitemRow = .Range("Y" & resultRow).value 'Get Invoice Row
                Invoice.Range("B" & invitemRow & ":I" & invitemRow).value = .Range("P" & resultRow & ":W" & resultRow).value 'Item Details
                Invoice.Range("K" & invitemRow).value = InvoiceItems.Range("X" & resultRow).value 'Item Cost
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
        invRow = .Range("B3").value              'Order Row
        wshCC_Invoice_List.Range(invRow & ":" & invRow).EntireRow.delete
        'Remove Invoice Items
        With InvoiceItems
            lastRow = .Range("A99999").End(xlUp).row
            If lastRow < 3 Then GoTo NotSaved
            .Range("A2:K" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("M2:M3"), CopyToRange:=.Range("P2:Y2"), Unique:=True
            lastResultRow = .Range("P99999").End(xlUp).row
            If lastResultRow < 3 Then GoTo NotSaved
            If lastResultRow < 4 Then GoTo SingleRow
            'Sort based on descending rows
            With .Sort
                .SortFields.clear
                .SortFields.add Key:=InvoiceItems.Range("P3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange InvoiceItems.Range("P3:Y" & lastResultRow)
                .Apply
            End With
SingleRow:
            For resultRow = 3 To lastResultRow
                invitemRow = .Range("P" & resultRow).value 'Invoice Item DB Row
                If invitemRow > 3 Then .Range(invitemRow & ":" & invitemRow).EntireRow.delete 'Don't remove 1st Row
            Next resultRow
        End With
NotSaved:
        Dashboard_Invoice_New                              'Clear Out All Invoice Fields
    End With
End Sub

