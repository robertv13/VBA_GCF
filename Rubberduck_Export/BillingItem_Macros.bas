Attribute VB_Name = "BillingItem_Macros"
Option Explicit
Dim EntryRow As Long, EntryCol As Long, LastRow As Long, LastResultRow As Long, SelRow As Long, InvRow As Long
Dim ServItem As String
Sub BillingEntry_LoadList()
Invoice.Range("B17,C12:I999").ClearContents
With BillEntries
    LastRow = .Range("A99999").End(xlUp).Row
    If LastRow < 4 Then Exit Sub
    .Range("A3:M" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("Q2:R3"), CopyToRange:=.Range("U2:AA2"), Unique:=True
    LastResultRow = .Range("U99999").End(xlUp).Row
    If LastResultRow < 3 Then Exit Sub
    Invoice.Range("C12:I" & LastResultRow + 9).Value = .Range("U3:AA" & LastResultRow).Value 'Bring Over Billing Item Results
    On Error Resume Next
    Invoice.Range("B17").Value = Invoice.Range("C12:C9999").Find(Invoice.Range("B2").Value, , xlFormulas, xlWhole).Row 'Set Selected Row (if applicable)
    On Error GoTo 0
End With
End Sub
Sub BillingEntry_New()
With Invoice
    .Range("B2,B17,E4:F7,H4:H7").ClearContents
    .Range("H4").Value = Date 'Set Current Date as default
    .Range("H7").Value = "No" 'Set Default billed to No
End With
End Sub

Sub BillingEntry_SaveUpdate()
With Invoice
    'Check For Required Fields
    If .Range("B14").Value < 4 Then
        MsgBox "Please make sure to enter a Customer, Project, Service & Date before saving"
        Exit Sub
    End If
    If .Range("B3").Value = Empty Then 'New Billing entry
        .Range("B2").Value = .Range("B4").Value 'Set Entry ID
        EntryRow = BillEntries.Range("A99999").End(xlUp).Row + 1  'First avail row
        BillEntries.Range("A" & EntryRow).Value = .Range("B2").Value 'Set Billing entry ID
        BillEntries.Range("M" & EntryRow).Value = "=Row()"
    Else 'Existing Entry
        EntryRow = .Range("B3").Value
    End If
    For EntryCol = 2 To 12
        BillEntries.Cells(EntryRow, EntryCol).Value = .Range(BillEntries.Cells(1, EntryCol).Value).Value 'Save Data
    Next EntryCol
    BillingEntry_LoadList
    MsgBox "Billing Entry Saved"
End With
End Sub

Sub BillingEntry_Load()
With Invoice
    If .Range("B3").Value = "" Then
        MsgBox "Please select a correct Billing entry"
        Exit Sub
    End If
    .Range("B23").Value = True 'Set Load To True
    EntryRow = .Range("B3").Value 'Set Entry Row
    For EntryCol = 3 To 12 'Not On Project, Customer Or Service ID
    If EntryCol = 4 Or EntryCol = 6 Then GoTo NextCol
        .Range(BillEntries.Cells(1, EntryCol).Value).Value = BillEntries.Cells(EntryRow, EntryCol).Value 'Load Data
NextCol:
    Next EntryCol
    .Range("B23").Value = False 'Set Load to false
End With
End Sub

Sub BillingEntry_Delete()
If MsgBox("Are you sure you want to delete this Billing Item?", vbYesNo, "Delete Billing Item") = vbNo Then Exit Sub
With Invoice
    If .Range("B3").Value = "" Then GoTo NotSaved
    EntryRow = .Range("B3").Value 'Entry Row
    BillEntries.Range(EntryRow & ":" & EntryRow).EntireRow.Delete
NotSaved:
    BillingEntry_New
    BillingEntry_LoadList
End With
End Sub

Sub BillingEntry_AddToInvoice()
With Invoice
    If .Range("B17").Value = Empty Then
        MsgBox "Please select on an item to add to invoice"
    End If
    If .Range("B20").Value = Empty Then
        MsgBox "Please make sure to save the invoice before adding billing items"
        Exit Sub
    End If
    SelRow = .Range("B17").Value 'Selected Row
    If .Range("H" & SelRow).Value = "Yes" Then
        If MsgBox("This item has already been billed. Are you sure you want to add it to the invoice again?", vbYesNo, "Already Billed") = vbNo Then Exit Sub
    End If
    
    EntryRow = .Range("I" & SelRow).Value 'Entry Database Row
    ServItem = .Range("F" & SelRow).Value 'Service Item Name
    .Range("B25").Value = True 'Set Item Load to True
    
    If .Range("B15").Value = True Then ' Total Like Service Items
        InvRow = 0
        On Error Resume Next
        InvRow = .Range("K9:K35").Find(ServItem, , xlValues, xlWhole).Row
        On Error GoTo 0
        If InvRow <> 0 Then
              .Range("M" & InvRow).Value = .Range("M" & InvRow).Value + .Range("G" & SelRow).Value 'Update Hours (maintain service item details
              GoTo ItemAdded
        Else
            GoTo NewRow
        End If
    Else 'Create New rows, even for like service items
NewRow:
            InvRow = .Range("K36").End(xlUp).Row + 1
            If InvRow = 36 Then 'Max Items
                MsgBox "You have reached the maxiumum # of Service items for this invoice"
                Exit Sub
            End If
    End If
    
    .Range("J" & InvRow).Value = .Range("D" & SelRow).Value 'Service Date
    .Range("K" & InvRow).Value = .Range("F" & SelRow).Value 'Service
    .Range("L" & InvRow).Value = BillEntries.Range("J" & EntryRow).Value 'Saved Description
    .Range("M" & InvRow).Value = .Range("G" & SelRow).Value ' Hours
    .Range("N" & InvRow).Value = BillEntries.Range("K" & EntryRow).Value 'Billing Rate
ItemAdded:
    .Range("B25").Value = False 'Set Item Load to false
    BillEntries.Range("L" & EntryRow).Value = "Yes" 'Set Billed Status to Yes
    .Shapes("AddItemBtn").Visible = msoFalse
    BillingEntry_LoadList 'Run Macro to Refresh list
End With
End Sub


Sub BillingEntry_AddAllItems()
MsgBox "This feature will be add for our Patreon members. We would love to have you join here https://www.patreon.com/ExcelForFreelancers"
End Sub

