Attribute VB_Name = "BillingItem_Macros"
Option Explicit
Dim EntryRow As Long, EntryCol As Long, lastRow As Long, LastResultRow As Long, SelRow As Long, InvRow As Long
Const BillingRate As Long = 350
Dim ServItem As String

Sub BillingEntry_LoadList() 'Filter appropriate WIP lines
    If shInvoice.Range("B28").value Then Debug.Print "Now entering - [BillingItem_Macros] - Sub BillingEntry_LoadList() @ " & Time
    Dim LineFrom_Copy As Long
    Dim LineTo_Copy As Long
    Dim BilledCriteria As String
    shInvoice.Range("B17,C12:H9999").ClearContents 'Maximum entries from 12 to 999 in the WIP section = 988 rows
    With shBillEntries
        'Clear destination area
        .Range("U3:Z9999").ClearContents
        lastRow = .Range("A9999").End(xlUp).Row
        If lastRow < 4 Then Exit Sub 'No WIP rows at all
        'Copy line per line, cell per cell
        LineTo_Copy = 3
        For LineFrom_Copy = 4 To lastRow
            If .Cells(LineFrom_Copy, 2).value = .Range("Q3").value Then
                If .Range("R3").value = "<>" Or .Cells(LineFrom_Copy, 12) = "No" Then
                    .Cells(LineTo_Copy, 21).value = .Cells(LineFrom_Copy, 1).value 'Bill Entry ID
                    .Cells(LineTo_Copy, 22).value = .Cells(LineFrom_Copy, 8).value 'Date
                    .Cells(LineTo_Copy, 23).value = .Cells(LineFrom_Copy, 10).value 'Description
                    .Cells(LineTo_Copy, 24).value = .Cells(LineFrom_Copy, 9).value 'Hours
                    .Cells(LineTo_Copy, 25).value = .Cells(LineFrom_Copy, 12).value 'Billed ?
                    .Cells(LineTo_Copy, 26).value = .Cells(LineFrom_Copy, 13).value 'Hours
                    LineTo_Copy = LineTo_Copy + 1
                End If
            End If
        Next LineFrom_Copy
        '.Range("A3:M" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("Q2:R3"), CopyToRange:=.Range("U2:AA2"), Unique:=True
        LastResultRow = .Range("U99999").End(xlUp).Row
        If LastResultRow < 3 Then Exit Sub
        shInvoice.Range("C12:I" & LastResultRow + 9).value = .Range("U3:AA" & LastResultRow).value 'Bring Over Billing Item Results
        On Error Resume Next
        shInvoice.Range("B17").value = shInvoice.Range("C12:C9999").Find(shInvoice.Range("B2").value, , xlFormulas, xlWhole).Row 'Set Selected Row (if applicable)
        On Error GoTo 0
    End With
    If shInvoice.Range("B28").value Then Debug.Print "Now exiting  - [BillingItem_Macros] - Sub BillingEntry_LoadList()" & vbNewLine
End Sub

Sub BillingEntry_New()
    If shInvoice.Range("B28").value Then Debug.Print "Now entering - [BillingItem_Macros] - Sub BillingEntry_New() @ " & Time
    With shInvoice
        .Range("B2,B17,E4:F7,H4:H7").ClearContents
        .Range("H4").value = Date 'Set Current Date as default
        .Range("H7").value = "No" 'Set Default billed to No
    End With
    If shInvoice.Range("B28").value Then Debug.Print "Now exiting  - [BillingItem_Macros] - Sub BillingEntry_New()" & vbNewLine
End Sub

Sub BillingEntry_SaveUpdate()
    If shInvoice.Range("B28").value Then Debug.Print "Now entering - [BillingItem_Macros] - Sub BillingEntry_SaveUpdate() @ " & Time
    With shInvoice
        'Check For Required Fields
        If .Range("B14").value < 4 Then
            MsgBox "Please make sure to enter a Customer, Project, Service & Date before saving"
            Exit Sub
        End If
        If .Range("B3").value = Empty Then 'New Billing entry
            .Range("B2").value = .Range("B4").value 'Set Entry ID
            EntryRow = shBillEntries.Range("A99999").End(xlUp).Row + 1  'First avail row
            shBillEntries.Range("A" & EntryRow).value = .Range("B2").value 'Set Billing entry ID
            shBillEntries.Range("M" & EntryRow).value = "=Row()"
        Else 'Existing Entry
            EntryRow = .Range("B3").value
        End If
        For EntryCol = 2 To 12
            shBillEntries.Cells(EntryRow, EntryCol).value = .Range(shBillEntries.Cells(1, EntryCol).value).value 'Save Data
        Next EntryCol
        BillingEntry_LoadList
        MsgBox "Billing Entry Saved"
    End With
    If shInvoice.Range("B28").value Then Debug.Print "Now exiting  - [BillingItem_Macros] - Sub BillingEntry_SaveUpdate()" & vbNewLine
End Sub

Sub BillingEntry_Load()
    If shInvoice.Range("B28").value Then Debug.Print "Now entering - [BillingItem_Macros] - Sub BillingEntry_Load() @ " & Time
    With shInvoice
        If .Range("B3").value = "" Then
            MsgBox "Veuillez sélectionner une charge valide"
            Exit Sub
        End If
        .Range("B23").value = True 'Set Load To True
        EntryRow = .Range("B3").value 'Set Entry Row
        For EntryCol = 8 To 12 'Not On Project, Customer Or Service ID
            If EntryCol = 4 Or EntryCol = 5 Or EntryCol = 6 Or EntryCol = 7 Then GoTo NextCol
            .Range(shBillEntries.Cells(1, EntryCol).value).value = shBillEntries.Cells(EntryRow, EntryCol).value 'Load Data
NextCol:
        Next EntryCol
        .Range("B23").value = False 'Set Load to false
    End With
    If shInvoice.Range("B28").value Then Debug.Print "Now exiting  - [BillingItem_Macros] - Sub BillingEntry_Load()" & vbNewLine
End Sub

Sub BillingEntry_Delete()
    If shInvoice.Range("B28").value Then Debug.Print "Now entering - Sub BillingEntry_Delete() @ " & Time
    If MsgBox("Are you sure you want to delete this Billing Item?", vbYesNo, "Delete Billing Item") = vbNo Then Exit Sub
    With shInvoice
        If .Range("B3").value = "" Then GoTo NotSaved
        EntryRow = .Range("B3").value 'Entry Row
        shBillEntries.Range(EntryRow & ":" & EntryRow).EntireRow.Delete
NotSaved:
        BillingEntry_New
        BillingEntry_LoadList
    End With
    If shInvoice.Range("B28").value Then Debug.Print "Now entering - [BillingItem_Macros] - Sub BillingEntry_Delete()" & vbNewLine
End Sub

Sub BillingEntry_AddToInvoice()
    If shInvoice.Range("B28").value Then Debug.Print "Now entering - [BillingItem_Macros] - Sub BillingEntry_AddToInvoice() @ " & Time
    With shInvoice
        If .Range("B17").value = Empty Then
            MsgBox "Vous devez sélectionner une charge pour l'ajouter à la facture"
        End If
        If .Range("B20").value = Empty Then
            MsgBox "Assurez-vous de sauvegarder la facture avant d'y ajouter des charges"
            Exit Sub
        End If
        SelRow = .Range("B17").value 'Selected Row
        If .Range("G" & SelRow).value = "Yes" Then
            If MsgBox("Cette charge a déjà été facturé. Êtes-vous certain de vouloir l'ajouter à nouveau ?", vbYesNo, "Charge déjà facturée") = vbNo Then Exit Sub
        End If
        
        EntryRow = .Range("H" & SelRow).value 'Entry Database Row
        ServItem = .Range("F" & SelRow).value 'Service Item Name
        .Range("B25").value = True 'Set Item Load to True
        
        .Range("B15").value = False 'RMV - 2023-09-29 - We dot not allow to add to an existing entry...
'        If .Range("B15").Value = True Then ' Total Like Service Items
'            InvRow = 0
'            On Error Resume Next
'            InvRow = .Range("K9:K35").Find(ServItem, , xlValues, xlWhole).Row
'            On Error GoTo 0
'            If InvRow <> 0 Then
'                .Range("M" & InvRow).Value = .Range("M" & InvRow).Value + .Range("G" & SelRow).Value 'Update Hours (maintain service item details
'                GoTo ItemAdded
'            Else
'                GoTo NewRow
'            End If
'        Else 'Create New rows, even for like service items
NewRow:
            InvRow = .Range("K36").End(xlUp).Row + 1
            InvRow = 10
            'MsgBox InvRow
            If InvRow = 36 Then 'Max Items
                MsgBox "Vous avez atteint le maximum d'entrée sur cette facture"
                Exit Sub
            End If
        'End If
        
        .Range("K" & InvRow).value = .Range("E" & SelRow).value 'Description
        '.Range("J" & InvRow).Value = .Range("D" & SelRow).Value 'Service Date
        '.Range("K" & InvRow).Value = .Range("F" & SelRow).Value 'Service
        '.Range("L" & InvRow).Value = shBillEntries.Range("J" & EntryRow).Value 'Saved Description
        .Range("L" & InvRow).value = .Range("F" & SelRow).value ' Hours
        '.Range("M" & InvRow).Value = shBillEntries.Range("K" & EntryRow).Value 'Billing Rate
        .Range("M" & InvRow).value = BillingRate 'Billing Rate
        '.Range("N" & InvRow).Value = shBillEntries.Range("K" & EntryRow).Value 'Billing Rate
ItemAdded:
        .Range("B25").value = False 'Set Item Load to false
        shBillEntries.Range("L" & EntryRow).value = "Yes" 'Set Billed Status to Yes
        .Shapes("AddItemBtn").Visible = msoFalse
        BillingEntry_LoadList 'Run Macro to Refresh list
    End With
    If shInvoice.Range("B28").value Then Debug.Print "Now exiting  - [BillingItem_Macros] - Sub BillingEntry_AddToInvoice()" & vbNewLine
End Sub

Sub BillingEntry_AddAllItems()
    If shInvoice.Range("B28").value Then Debug.Print "Now entering - [BillingItem_Macros] - Sub BillingEntry_AddAllItems() @ " & Time
    MsgBox "This feature will be add for our Patreon members. We would love to have you join here https://www.patreon.com/ExcelForFreelancers"
    If shInvoice.Range("B28").value Then Debug.Print "Now exiting  - [BillingItem_Macros] - Sub BillingEntry_AddAllItems()" & vbNewLine
End Sub

