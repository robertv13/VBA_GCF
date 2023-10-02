Attribute VB_Name = "Invoice_Macros"
Option Explicit
Dim InvRow As Long, InvCol As Long, ItemDBRow As Long, InvItemRow As Long, InvNumb As Long
Dim LastRow As Long, LastItemRow As Long, LastResultRow As Long, ResultRow As Long
Sub Invoice_New()
Invoice.Range("K4:K6,N3:N5,N6:O6,I9:N35,P9:Q35,O37").ClearContents
Invoice.Range("N3").Value = Invoice.Range("B21").Value 'Set Next Invoice ID
Invoice.Range("N4").Value = Date 'Set Current Date Default
End Sub

Sub Invoice_SaveUpdate()
With Invoice
    'Check For Required Fields
    If .Range("B18").Value = Empty Then
        MsgBox "Please make sure to add a customer before saving invoice"
        Exit Sub
    End If
    If .Range("B20").Value = Empty Then 'New Invoice
       InvRow = InvList.Range("A99999").End(xlUp).Row + 1
       InvList.Range("A" & InvRow).Value = .Range("B21").Value ' Next Invoice #
    Else 'Existing Invoice
        InvRow = .Range("B20").Value 'Set Existing Invoice Row
    End If
        For InvCol = 2 To 8
            InvList.Cells(InvRow, InvCol).Value = .Range(InvList.Cells(1, InvCol).Value).Value 'Save Invoice List Data
        Next InvCol
'Save/Update Invoice Items
    LastItemRow = .Range("K35").End(xlUp).Row
    If LastItemRow < 9 Then GoTo NoItems
    For InvItemRow = 9 To LastItemRow
            If .Range("P" & InvItemRow).Value = "" Then
                ItemDBRow = InvItems.Range("A99999").End(xlUp).Row + 1
                .Range("P" & InvItemRow).Value = ItemDBRow 'Set Item DB Row
                InvItems.Range("A" & ItemDBRow).Value = .Range("N3").Value 'Invoice #
                InvItems.Range("H" & ItemDBRow).Value = InvItemRow 'Set Invoice Row
                InvItems.Range("I" & ItemDBRow).Value = "=Row()"
            Else 'Existing Item
                ItemDBRow = .Range("P" & InvItemRow).Value  'Invoice Item Row
            End If
            InvItems.Range("B" & ItemDBRow & ":G" & ItemDBRow).Value = .Range("J" & InvItemRow & ":O" & InvItemRow).Value 'Save Invoice Item Details
    Next InvItemRow
NoItems:
MsgBox "Invoice Saved"
End With
End Sub

Sub Invoice_Load()
With Invoice
    If .Range("B20").Value = Empty Then
        MsgBox "Please enter a correct Invoice #"
        Exit Sub
    End If
     .Range("B24").Value = True 'Set Invoice Load to true
     .Range("R2,K4:K6,N4:N5,N6:O6,J9:N35,P9:P35,O37").ClearContents
    InvRow = .Range("B20").Value
   
    For InvCol = 2 To 7
           If InvCol <> 3 Then .Range(InvList.Cells(1, InvCol).Value).Value = InvList.Cells(InvRow, InvCol).Value 'Load Invoice List Data
    Next InvCol
    'Load Invoice Items
    With InvItems
        LastRow = .Range("A99999").End(xlUp).Row
        If LastRow < 4 Then Exit Sub
        .Range("A3:J" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("N2:N3"), CopyToRange:=.Range("P2:V2"), Unique:=True
        LastResultRow = .Range("V99999").End(xlUp).Row
        If LastResultRow < 3 Then GoTo NoItems
        For ResultRow = 3 To LastResultRow
            InvItemRow = .Range("U" & ResultRow).Value 'Set Invoice Row
            Invoice.Range("J" & InvItemRow & ":N" & InvItemRow).Value = .Range("P" & ResultRow & ":T" & ResultRow).Value 'Item details
            Invoice.Range("P" & InvItemRow).Value = .Range("V" & ResultRow).Value  'Set Item DB Row
        Next ResultRow
NoItems:
    End With
    .Range("B24").Value = False 'Set Invoice Load To false
End With
End Sub

Sub Invoice_Delete()
With Invoice
    If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Invoice") = vbNo Then Exit Sub
    If .Range("B20").Value = Empty Then GoTo NotSaved
    InvRow = .Range("B20").Value 'Set Invoice Row
    InvList.Range(InvRow & ":" & InvRow).EntireRow.Delete
    With InvItems
        LastRow = .Range("A99999").End(xlUp).Row
        If LastRow < 4 Then Exit Sub
        .Range("A3:J" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("N2:N3"), CopyToRange:=.Range("P2:W2"), Unique:=True
        LastResultRow = .Range("V99999").End(xlUp).Row
        If LastResultRow < 3 Then GoTo NoItems
'        If LastResultRow < 4 Then GoTo SkipSort
'        'Sort Rows Descending
'         With .Sort
'         .SortFields.Clear
'         .SortFields.Add Key:=InvItems.Range("W3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal  'Sort
'         .SetRange InvItems.Range("P3:W" & LastResultRow) 'Set Range
'         .Apply 'Apply Sort
'         End With
SkipSort:
        For ResultRow = 3 To LastResultRow
            ItemDBRow = .Range("V" & ResultRow).Value 'Set Invoice Database Row
            .Range("A" & ItemDBRow & ":J" & ItemDBRow).ClearContents 'Clear Fields (deleting creates issues with results
        Next ResultRow
        'Resort DB to remove spaces
         With .Sort
         .SortFields.Clear
         .SortFields.Add Key:=InvItems.Range("A4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal  'Sort
         .SetRange InvItems.Range("A4:J" & LastResultRow) 'Set Range
         .Apply 'Apply Sort
         End With
    End With
NoItems:
NotSaved:
Invoice_New 'Add New Invoice
End With
End Sub

Sub Invoice_Print() 'RMV_IMPRESSION
    
    Invoice.PrintOut , , , , True, , , , False

End Sub

Sub Prev_Invoice()
With Invoice
    Dim MinInvNumb As Long
    On Error Resume Next
    MinInvNumb = Application.WorksheetFunction.Min(InvList.Range("Inv_ID"))
    On Error GoTo 0
    If MinInvNumb = 0 Then
        MsgBox "Please create and save an Invoice first"
        Exit Sub
    End If
    InvNumb = .Range("N3").Value
    If InvNumb = 0 Or .Range("B20").Value = Empty Then 'On New Invoice
        InvRow = InvList.Range("A99999").End(xlUp).Row 'On Empty Invoice Go to last one created
    Else 'On Existing Inv. find Previous one
        InvRow = InvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).Row - 1
    End If
    If .Range("N3").Value = 1 Or MinInvNumb = 0 Or MinInvNumb = .Range("N3").Value Then
        MsgBox "You are at the first invoice"
        Exit Sub
    End If
    .Range("N3").Value = InvList.Range("A" & InvRow).Value 'Place Inv. ID inside cell
    Invoice_Load
End With
End Sub


Sub Next_Invoice()
With Invoice
    Dim MaxInvNumb As Long
    On Error Resume Next
    MaxInvNumb = Application.WorksheetFunction.Max(InvList.Range("Inv_ID"))
    On Error GoTo 0
    If MaxInvNumb = 0 Then
        MsgBox "Please create and save an Invoice first"
        Exit Sub
    End If
    InvNumb = .Range("N3").Value
    If InvNumb = 0 Or .Range("B20").Value = Empty Then 'On New Invoice
        InvRow = InvList.Range("A4").Value  'On Empty Invoice Go to First one created
    Else 'On Existing Inv. find Previous one
        InvRow = InvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).Row + 1
    End If
    If .Range("N3").Value >= MaxInvNumb Then
        MsgBox "You are at the last invoice"
        Exit Sub
    End If
    .Range("N3").Value = InvList.Range("A" & InvRow).Value 'Place Inv. ID inside cell
    Invoice_Load
End With
End Sub

