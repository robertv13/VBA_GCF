Attribute VB_Name = "Invoice_Macros"
Option Explicit
Dim InvRow As Long, InvItemRow As Long, LastRow As Long, LastItemRow As Long, FieldRow As Long, ItemDBRow As Long
Dim ResultRow As Long, LastResultRow As Long, ItemRow As Long, TermRow As Long, StatusRow As Long, InvNumb As Long
Dim HeadShp As Shape

Sub Invoice_SaveUpdate()
    With Invoice
        If .Range("B3").Value = Empty Then
            MsgBox "Please make sure to add a customer before saving invoice"
            Exit Sub
        End If
        AppEvents_Stop
        'Deterermine New Invoice/Existing Invoice
        If .Range("B5").Value = Empty Then       'new Invoice
            InvRow = InvoiceList.Range("A99999").End(xlUp).Row + 1 'First Avail Row
            .Range("B4").Value = .Range("B6").Value 'Next Inv. #
            InvoiceList.Range("A" & InvRow).Value = .Range("B4").Value 'Next Inv. #
            InvoiceList.Range("J" & InvRow).Value = "=H" & InvRow & "-IFERROR(SUMIF(PayItem_InvID,A" & InvRow & ",PayItem_Amount),0)"
        Else                                     'Existing Invoice
            InvRow = .Range("B5").Value          'Invoice Row
        End If
        For FieldRow = 4 To 9
            InvoiceList.Cells(InvRow, FieldRow - 2).Value = .Range("E" & FieldRow).Value 'Add Invoice Details
        Next FieldRow
        InvoiceList.Range("H" & InvRow).Value = .Range("AD21").Value 'Invoice Total
        InvoiceList.Range("I" & InvRow).Value = .Range("B12").Value 'Total Pages
    End With
    AppEvents_Start
    If Application.Caller = "SaveBtn" Then Invoice_SavedMsg 'Run Invoice Fade Out Message
End Sub

Sub Invoice_SavedMsg()
    With Invoice.Shapes("InvSavedMsg")
        Dim i As Long, Delay As Double, StartTime As Double
        .Visible = msoCTrue
        For i = 1 To 100
            .Fill.Transparency = i / 100
            Delay = 0.009
            StartTime = Timer
            Do
                DoEvents
            Loop While Timer - StartTime < Delay
        Next i
        .Visible = msoFalse
    End With
End Sub

Sub Customer_AddUpdate()
    Dim CustRow As Long, CustCol As Long
    Dim CustomerFld As Control
    Unload AddCustForm
    'If Existing Customer is selected, Load Customer Details
    If Invoice.Range("B3").Value <> Empty Then   'Existing Customer
        CustRow = Invoice.Range("B3").Value
        For CustCol = 2 To 8                     'Loop Through Customer Columns
            Set CustomerFld = AddCustForm.Controls("Field" & CustCol - 1)
            CustomerFld.Value = Customers.Cells(CustRow, CustCol).Value 'Map Data from customers into userform
        Next CustCol
    End If
    AddCustForm.Show
End Sub

Sub Invoice_New()
    With Invoice
        .Range("B1").Value = True                'Set Inv. Load to true
        .Range("B4,F2,E4:F9,J10:N44,P10:Q44").ClearContents
        .Range("B4").Value = .Range("B6").Value  'Set Next invoice #
        .Range("E4").Value = Date                'Set current Date
        .Range("B11,B12").Value = 1              'Set Selected Page and Total Pages to 1
        .Range("B1").Value = False               'Set inv. Load to false
        On Error Resume Next
        TermRow = Admin.Range("J6:J11").Find(Chr(252), , xlValues, xlWhole).Row 'Default Payment Term Row
        On Error GoTo 0
        If TermRow <> 0 Then .Range("E6").Value = Admin.Range("H" & TermRow).Value 'Set Default Term
        On Error Resume Next
        StatusRow = Admin.Range("F6:F12").Find(Chr(252), , xlValues, xlWhole).Row 'Default Status Row
        On Error GoTo 0
        If StatusRow <> 0 Then .Range("E7").Value = Admin.Range("E" & StatusRow).Value 'Set Default Status
        .Range("E5").Select
    End With
End Sub

Sub Invoice_Load()
    With Invoice
        If .Range("B5").Value = Empty Then
            MsgBox "Please enter a correct invoice #"
            Exit Sub
        End If
        AppEvents_Stop
        InvRow = .Range("B5").Value              'Invoice Row
        .Range("B1").Value = True                'Set Inv. Load to true
        .Range("F2,E4:F9,J10:N44,P10:Q44").ClearContents
        .Range("B11").Value = 1                  'Set Default To Load First Page
        For FieldRow = 4 To 9
            .Range("E" & FieldRow).Value = InvoiceList.Cells(InvRow, FieldRow - 2).Value 'Fill Invoice Details
        Next FieldRow
        .Range("B12").Value = InvoiceList.Range("I" & InvRow).Value 'Set Total Pages
        Invoice_PageLoad                         'Load Page Items
        .Range("B1").Value = False               'Set inv. Load to false
        AppEvents_Start
    End With
End Sub

Sub Invoice_PageLoad()
    If Invoice.Range("B5").Value = Empty Then Exit Sub 'Exit on no invoice
    Invoice.Range("B1").Value = True             'Set Inv. Load to true
    Invoice.Range("J10:N44,P10:Q44").ClearContents 'Clear Existing Page Items
    With InvoiceItems
        LastRow = .Range("A99999").End(xlUp).Row
        If LastRow < 4 Then GoTo NoItems
        .Range("A3:K" & LastRow).AdvancedFilter xlFilterCopy, criteriaRange:=.Range("M2:N3"), copytorange:=.Range("P2:W2"), Unique:=True
        LastResultRow = .Range("P99999").End(xlUp).Row
        If LastResultRow < 3 Then GoTo NoItems
        For ResultRow = 3 To LastResultRow
            InvItemRow = .Range("P" & ResultRow).Value 'Get Invoice Row
            Invoice.Range("J" & InvItemRow & ":N" & InvItemRow).Value = .Range("Q" & ResultRow & ":U" & ResultRow).Value 'Item Details
            Invoice.Range("P" & InvItemRow & ":Q" & InvItemRow).Value = .Range("V" & ResultRow & ":W" & ResultRow).Value 'Tax & DB Row
        Next ResultRow
NoItems:
    End With
    Invoice.Range("B1").Value = False            'Set inv. Load to false
End Sub

Sub Invoice_PrevPage()
    If Invoice.Range("B11").Value = 1 Then
        MsgBox "You are on the first page"
        Exit Sub
    End If
    Invoice.Range("B11").Value = Invoice.Range("B11").Value - 1 'Increment Page # Down
    Invoice_PageLoad
End Sub

Sub Invoice_NextPage()
    If Invoice.Range("B11").Value = Invoice.Range("B12").Value Then
        MsgBox "You are on the last page"
        Exit Sub
    End If
    Invoice.Range("B11").Value = Invoice.Range("B11").Value + 1 'Increment Page # Up
    Invoice_PageLoad
End Sub

Sub Invoice_NewPage()
    Invoice_SaveUpdate                           'Save New Last page limit
    Invoice.Range("B11").Value = Invoice.Range("B12").Value + 1 'Set Selected page to new page
    Invoice.Range("B12").Value = Invoice.Range("B12").Value + 1 ' Increment page to + 1
    Invoice_PageLoad
    Invoice_SaveUpdate                           'Save New Last page limit
End Sub

Sub Invoice_PrevInvoice()
    With Invoice
        Dim MinInvNumb As Long
        On Error Resume Next
        MinInvNumb = Application.WorksheetFunction.Min(InvoiceList.Range("Invoice_ID"))
        On Error GoTo 0
        If MinInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("B4").Value
        If InvNumb = 0 Or .Range("B5").Value = Empty Then 'On New Invoice
            InvRow = InvoiceList.Range("A99999").End(xlUp).Row 'On Empty Invoice Go to last one created
        Else                                     'On Existing Inv. find Previous one
            InvRow = InvoiceList.Range("Invoice_ID").Find(InvNumb, , xlValues, xlWhole).Row - 1
        End If
        If .Range("B6").Value = 1 Or MinInvNumb = 0 Or MinInvNumb = .Range("B4").Value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("B4").Value = InvoiceList.Range("A" & InvRow).Value 'Place Inv. ID inside cell
        Invoice_Load
    End With
End Sub

Sub Invoice_NextInvoice()
    With Invoice
        Dim MaxInvNumb As Long
        On Error Resume Next
        MaxInvNumb = Application.WorksheetFunction.Max(InvoiceList.Range("Invoice_ID"))
        On Error GoTo 0
        InvNumb = .Range("B4").Value
        If MaxInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        If InvNumb = MaxInvNumb Then
            MsgBox "You are at the last Invoice"
            Exit Sub
        End If
    
    
        If InvNumb = 0 Or .Range("B5").Value = Empty Then 'On New Invoice
            InvRow = 3                           'On Empty Invoice Go to First one created
        Else                                     'On Existing Inv. find Next one
            InvRow = InvoiceList.Range("Invoice_ID").Find(InvNumb, , xlValues, xlWhole).Row + 1
        End If

        .Range("B4").Value = InvoiceList.Range("A" & InvRow).Value
        Invoice_Load
    End With
End Sub

Sub Invoice_Delete()
    If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Invoice") = vbNo Then Exit Sub
    With Invoice
        AppEvents_Stop
        If .Range("B5").Value = Empty Then GoTo NotSaved
        InvRow = .Range("B5").Value              'Order Row
        InvoiceList.Range(InvRow & ":" & InvRow).EntireRow.Delete 'Delete Invoice row
        'Remove Invoice Items
        With InvoiceItems
            LastRow = .Range("A99999").End(xlUp).Row
            If LastRow < 3 Then GoTo NotSaved
            .Range("A3:K" & LastRow).AdvancedFilter xlFilterCopy, criteriaRange:=.Range("M2:M3"), copytorange:=.Range("P2:W2"), Unique:=True
            LastResultRow = .Range("P99999").End(xlUp).Row
            If LastResultRow < 3 Then GoTo NotSaved
            If LastResultRow < 4 Then GoTo SingleRow
            'Sort based on descending rows
            With .Sort
                .SortFields.Clear
                .SortFields.Add Key:=InvoiceItems.Range("W3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange InvoiceItems.Range("P3:W" & LastResultRow)
                .Apply
            End With
SingleRow:
            For ResultRow = 3 To LastResultRow
                ItemDBRow = .Range("W" & ResultRow).Value 'Invoice Item DB Row
                .Range(ItemDBRow & ":" & ItemDBRow).EntireRow.Delete 'Delete Invoice Item Row
            Next ResultRow
        End With
NotSaved:
        Invoice_New                              'Clear Out All Invoice Fields
        AppEvents_Start
    End With
End Sub

