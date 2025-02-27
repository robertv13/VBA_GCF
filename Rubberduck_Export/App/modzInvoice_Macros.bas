Attribute VB_Name = "modzInvoice_Macros"
Option Explicit
Dim invRow As Long, invitemRow As Long, lastRow As Long, lastItemRow As Long
Dim resultRow As Long, lastResultRow As Long, itemRow As Long, termRow As Long, statusRow As Long

Sub Dashboard_Invoice_SaveUpdate()
    With Invoice
        If .Range("G5").Value = Empty Then
            MsgBox "Please make sure to add a customer before saving invoice"
            Exit Sub
        End If
        
        'Deterermine New Invoice/Existing Invoice
        If .Range("B3").Value = Empty Then       'new Invoice
            invRow = wshFAC_Invoice_List.Range("A99999").End(xlUp).row + 1 'First Avail Row
            .Range("J1").Value = .Range("B5").Value 'Next Inv. #
            wshFAC_Invoice_List.Range("A" & invRow).Value = .Range("B5").Value 'Next Inv. #
        Else                                     'Existing Invoice
            invRow = .Range("B3").Value          'Invoice Row
        End If
        wshFAC_Invoice_List.Range("B" & invRow).Value = .Range("I3").Value 'Date
        wshFAC_Invoice_List.Range("C" & invRow).Value = .Range("G5").Value 'Customer
        wshFAC_Invoice_List.Range("D" & invRow).Value = .Range("I4").Value 'Status
        wshFAC_Invoice_List.Range("E" & invRow).Value = .Range("I5").Value 'Terms
        wshFAC_Invoice_List.Range("F" & invRow).Value = .Range("I6").Value 'Due Date
        wshFAC_Invoice_List.Range("G" & invRow).Value = .Range("J34").Value 'Invoice Total
            
            
        'Add/Update Invoice Items
        lastItemRow = .Range("C31").End(xlUp).row 'Last Invoice Row
        If lastItemRow < 9 Then GoTo NoItems
        For itemRow = 9 To lastItemRow
            If .Range("B" & itemRow).Value <> Empty Then 'DB Row Exists
                invitemRow = .Range("B" & itemRow).Value 'Inv. Item DB Row
            Else                                 'New DB itemRow
                invitemRow = InvoiceItems.Range("A99999").End(xlUp).row + 1 'First avail Row
                InvoiceItems.Range("A" & invitemRow).Value = .Range("J1").Value 'Inv ID
                .Range("B" & itemRow).Value = invitemRow 'Add DB Row
                InvoiceItems.Range("K" & invitemRow).Value = "=Row()" 'Add Row Formula
            End If
            InvoiceItems.Range("B" & invitemRow & ":H" & invitemRow).Value = .Range("C" & itemRow & ":I" & itemRow).Value ' Bring over item Values
            InvoiceItems.Range("I" & invitemRow).Value = .Range("K" & itemRow).Value 'Item Cost
            InvoiceItems.Range("J" & invitemRow).Value = itemRow 'Invoice Row
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
        .Range("B6").Value = True                'Set Inv. Load to true
        .Range("I3:J6,G5:G7,B9:I31,K9:K31").Clearcontents
        .Range("J1").Value = .Range("B5").Value  'Set Next invoice #
        .Range("I3").Value = Date                'Set current Date
        .Range("B6").Value = False               'Set inv. Load to false
        On Error Resume Next
        termRow = Admin.Range("H6:H23").Find(Chr(252), , xlValues, xlWhole).row
        On Error GoTo 0
        If termRow <> 0 Then .Range("I5").Value = Admin.Range("F" & termRow).Value 'Set Default Term
        On Error Resume Next
        statusRow = Admin.Range("D6:D12").Find(Chr(252), , xlValues, xlWhole).row
        On Error GoTo 0
        If statusRow <> 0 Then .Range("I4").Value = Admin.Range("C" & statusRow).Value 'Set Default Status
    End With
End Sub

Sub Dashboard_Invoice_Load()
    With Invoice
        If .Range("B3").Value = Empty Then
            MsgBox "Please entere a correct invoice #"
            Exit Sub
        End If
        invRow = .Range("B3").Value              'Invoice Row
        .Range("B6").Value = True                'Set Inv. Load to true
        .Range("I3:J6,G5:G7,B9:I31,K9:K31").Clearcontents
        .Range("I3").Value = wshFAC_Invoice_List.Range("B" & invRow).Value 'Inv. Date
        .Range("G5").Value = wshFAC_Invoice_List.Range("C" & invRow).Value 'Customer
        .Range("I4").Value = wshFAC_Invoice_List.Range("D" & invRow).Value 'Inv. Status
        .Range("I5").Value = wshFAC_Invoice_List.Range("E" & invRow).Value 'Terms
        .Range("I6").Value = wshFAC_Invoice_List.Range("F" & invRow).Value 'Due Date
    
        With InvoiceItems
            lastRow = .Range("A99999").End(xlUp).row
            If lastRow < 3 Then GoTo NoItems
            .Range("A2:K" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("M2:M3"), CopyToRange:=.Range("P2:Y2"), Unique:=True
            lastResultRow = .Range("P99999").End(xlUp).row
            If lastResultRow < 3 Then GoTo NoItems
            For resultRow = 3 To lastResultRow
                invitemRow = .Range("Y" & resultRow).Value 'Get Invoice Row
                Invoice.Range("B" & invitemRow & ":I" & invitemRow).Value = .Range("P" & resultRow & ":W" & resultRow).Value 'Item Details
                Invoice.Range("K" & invitemRow).Value = InvoiceItems.Range("X" & resultRow).Value 'Item Cost
            Next resultRow
        End With
NoItems:
        .Range("B6").Value = False               'Set inv. Load to false
    End With
End Sub

Sub Invoice_Print()
    Invoice.PrintOut , , , , True, , , , False
End Sub

Sub Invoice_SaveAsPDF()
    Dim FilePath As String
    Dashboard_Invoice_SaveUpdate                           'Save invoice
    FilePath = ThisWorkbook.Path & "\" & Invoice.Range("G5").Value & "_" & Invoice.Range("J1").Value 'File Path
    If Dir(FilePath, vbDirectory) <> "" Then Kill (FilePath)
    Invoice.ExportAsFixedFormat xlTypePDF, FilePath, , , False, , , True
End Sub

Sub Invoice_Delete()
    If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Appointment") = vbNo Then Exit Sub
    With Invoice
        If .Range("B3").Value = Empty Then GoTo NotSaved
        invRow = .Range("B3").Value              'Order Row
        wshFAC_Invoice_List.Range(invRow & ":" & invRow).EntireRow.delete
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
                .SortFields.add key:=InvoiceItems.Range("P3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange InvoiceItems.Range("P3:Y" & lastResultRow)
                .Apply
            End With
SingleRow:
            For resultRow = 3 To lastResultRow
                invitemRow = .Range("P" & resultRow).Value 'Invoice Item DB Row
                If invitemRow > 3 Then .Range(invitemRow & ":" & invitemRow).EntireRow.delete 'Don't remove 1st Row
            Next resultRow
        End With
NotSaved:
        Dashboard_Invoice_New                              'Clear Out All Invoice Fields
    End With
End Sub

