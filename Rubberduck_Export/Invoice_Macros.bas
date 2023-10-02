Attribute VB_Name = "Invoice_Macros"
Option Explicit
Dim InvRow As Long, InvCol As Long, ItemDBRow As Long, InvItemRow As Long, InvNumb As Long
Dim LastRow As Long, LastItemRow As Long, LastResultRow As Long, ResultRow As Long

Sub Invoice_New()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_New() @ " & Time
    Dim NextInvNum As Long
    With shInvoice
        .Range("J4:K4").ClearContents 'Clear cells for a new Invoice
        .Range("I10:M46,O10:O46,N48,N49,N52").ClearContents
        NextInvNum = .Range("B21").Value
        .Range("N6").Value = NextInvNum 'Set Next Invoice ID
        .Range("B21").Value = NextInvNum + 1
        .Range("B20").Value = ""
    End With
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_New() - Facture # " & shInvoice.Range("N6").Value & " - " & shInvoice.Range("J3").Value & vbNewLine
End Sub

Sub Invoice_SaveUpdate()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_SaveUpdate() @ " & Time
    If shInvoice.Range("B26").Value Then Debug.Print Tab(5); "B18 = " & shInvoice.Range("B18").Value & "   B20 = " & shInvoice.Range("B20").Value
    With shInvoice
        'Check For Mandatory Fields
        If .Range("B18").Value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            If shInvoice.Range("B26").Value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que le nom de client n'est pas encore saisi, sortie de la routine"
            Exit Sub
        End If
        'Determine the row number (InvRow) for InvList
        If .Range("B20").Value = Empty Then 'New Invoice
            InvRow = InvList.Range("A99999").End(xlUp).Row + 1
            InvList.Range("A" & InvRow).Value = shInvoice.Range("N6").Value 'Invoice #
            If shInvoice.Range("B26").Value Then Debug.Print Tab(10); "Cas A (B20 = '""' ) alors InvRow est établi avec les lignes existantes: InvRow = " & InvRow
        Else 'Existing Invoice
            InvRow = .Range("B20").Value 'Set Existing Invoice Row
             If shInvoice.Range("B26").Value Then Debug.Print Tab(10); "Cas B (B20 <> '""') alors B20 est utilisé: InvRow = " & InvRow
        End If
        If shInvoice.Range("B26").Value Then Debug.Print Tab(5); "B20 = " & .Range("B20").Value & "   B21 (Next Invoice #) = " & .Range("B21").Value
        'Load data into InvList (Invoice Header)
        If shInvoice.Range("B26").Value Then Debug.Print Tab(5); "Facture # = " & Format(shInvoice.Range("N6").Value, "00000") & " et InvRow = " & InvRow & " - Posting; dans; InvoiceListing; """
        For InvCol = 2 To 12
            InvList.Cells(InvRow, InvCol).Value = .Range(InvList.Cells(1, InvCol).Value).Value 'Save data into Invoice List
            If shInvoice.Range("B26").Value Then Debug.Print Tab(10); "InvRow = " & InvRow & "   InvCol = " & InvCol & "   From Cell  = " & InvList.Cells(1, InvCol).Value & "   et la valeur = " & .Range(InvList.Cells(1, InvCol).Value).Value
        Next InvCol
        'Load data into InvItems (Save/Update Invoice Items)
        LastItemRow = .Range("K46").End(xlUp).Row
        If LastItemRow < 10 Then GoTo NoItems
        For InvItemRow = 10 To LastItemRow
            If .Range("O" & InvItemRow).Value = "" Then
                ItemDBRow = InvItems.Range("A99999").End(xlUp).Row + 1
                .Range("O" & InvItemRow).Value = ItemDBRow 'Set Item DB Row
                InvItems.Range("A" & ItemDBRow).Value = .Range("N6").Value 'Invoice #
                InvItems.Range("F" & ItemDBRow).Value = InvItemRow 'Set Invoice Row
                InvItems.Range("G" & ItemDBRow).Value = "=Row()"
            Else 'Existing Item
                ItemDBRow = .Range("O" & InvItemRow).Value  'Invoice Item Row
            End If
            'Paste 4 columns with one instruction
            InvItems.Range("B" & ItemDBRow & ":E" & ItemDBRow).Value = .Range("K" & InvItemRow & ":N" & InvItemRow).Value 'Save Invoice Item Details
            If shInvoice.Range("B26").Value Then Debug.Print Tab(15); "C" & ItemDBRow & " = " & InvItems.Range("C" & ItemDBRow).Value & "   D" & ItemDBRow & " = " & InvItems.Range("D" & ItemDBRow).Value & "   E" & ItemDBRow & " = " & InvItems.Range("E" & ItemDBRow).Value
        Next InvItemRow
NoItems:
        MsgBox "La facture '" & Format(.Range("N6").Value, "00000") & "' est enregistrée." & vbNewLine & vbNewLine & "Le total de la facture est " & Trim(Format(.Range("N50").Value, "### ##0.00 $")) & " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    End With
    If shInvoice.Range("B26").Value Then Debug.Print Tab(5); "Total de la facture '" & Format(shInvoice.Range("N6").Value, "00000") & "' (avant taxes) est de " & Format(shInvoice.Range("N50").Value, "### ##0.00 $")
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_SaveUpdate()" & vbNewLine
End Sub

Sub Invoice_Load()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_Load() @ " & Time
    With shInvoice
        If .Range("B20").Value = Empty Then
            MsgBox "Veuillez saisir un numéro de facture"
            Exit Sub
        End If
        .Range("B24").Value = True 'Set Invoice Load to true
        .Range("Q2,J4:J6,N3:N4,M6:N6,I10:M35,O10:O35").ClearContents
        InvRow = .Range("B20").Value
       
        'Assign values from InvList to Invoice worksheet
        For InvCol = 2 To 11 'RMV - 2023-10-01
            If shInvoice.Range("B26").Value And InvCol <> 3 Then Debug.Print "InvRow = " & InvRow & "   InvCol = " & InvCol & " - " & .Range(InvList.Cells(1, InvCol).Value) & " <-- " & InvList.Cells(InvRow, InvCol).Value
            If InvCol <> 3 Then .Range(InvList.Cells(1, InvCol).Value).Value = InvList.Cells(InvRow, InvCol).Value 'Load Invoice List Data
        Next InvCol
        'Load Invoice Items
        With InvItems
            LastRow = .Range("A9999").End(xlUp).Row
            If LastRow < 4 Then Exit Sub
            If shInvoice.Range("B26").Value Then Debug.Print "LastRow = " & LastRow & "   Copie de '" & "A3:G" & LastRow & "   Critère: " & .Range("L3").Value
            .Range("A3:G" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("L2:L3"), CopyToRange:=.Range("N2:S2"), Unique:=True
            LastResultRow = .Range("V9999").End(xlUp).Row
            If shInvoice.Range("B26").Value Then Debug.Print "Based on column 'V' (InvItems), LastResultRow = " & LastResultRow
            If LastResultRow < 3 Then GoTo NoItems
            For ResultRow = 3 To LastResultRow
                InvItemRow = .Range("R" & ResultRow).Value 'Set Invoice Row
                If shInvoice.Range("B26").Value Then Debug.Print Tab(20); "Invoice Item Row (InvItemRow) = " & InvItemRow & _
                    "   shInvoice.Range('K'" & InvItemRow & ")=" & shInvoice.Range("K" & InvItemRow).Value & " devient " & "InvItems.Range('N'" & ResultRow & ") = " & .Range("N" & ResultRow).Value & _
                    "   shInvoice.Range('L'" & InvItemRow & ")=" & shInvoice.Range("L" & InvItemRow).Value & " devient " & "InvItems.Range('O'" & ResultRow & ") = " & .Range("O" & ResultRow).Value & _
                    "   shInvoice.Range('M'" & InvItemRow & ")=" & shInvoice.Range("M" & InvItemRow).Value & " devient " & "InvItems.Range('P'" & ResultRow & ") = " & .Range("P" & ResultRow).Value & _
                shInvoice.Range("K" & InvItemRow & ":M" & InvItemRow).Value = .Range("N" & ResultRow & ":P" & ResultRow).Value 'Item details
                If shInvoice.Range("B26").Value Then Debug.Print Tab(30); "shInvoice.Range('O'" & InvItemRow & ")=" & shInvoice.Range("O" & InvItemRow).Value & " devient " & "InvItems.Range('S'" & ResultRow & ") = " & .Range("S" & ResultRow).Value
                shInvoice.Range("O" & InvItemRow).Value = .Range("S" & ResultRow).Value  'Set Item DB Row
            Next ResultRow
NoItems:
        End With
        .Range("B24").Value = False 'Set Invoice Load To false
    End With
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_Load()" & vbNewLine
End Sub

Sub Invoice_Delete()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_Delete() @ " & Time
    With shInvoice
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
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_Delete()" & vbNewLine
End Sub

Sub Invoice_Print() 'RMV_IMPRESSION
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_Print() @ " & Time
    shInvoice.PrintOut , , , True, True, , , , False
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_Print()" & vbNewLine
End Sub

Sub Prev_Invoice()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Prev_Invoice() @ " & Time
    With shInvoice
        Dim MinInvNumb As Long
        On Error Resume Next
        MinInvNumb = Application.WorksheetFunction.Min(InvList.Range("Inv_ID"))
        On Error GoTo 0
        If MinInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("N6").Value
        If InvNumb = 0 Or .Range("B20").Value = Empty Then 'On New Invoice
            InvRow = InvList.Range("A99999").End(xlUp).Row 'On Empty Invoice Go to last one created
        Else 'On Existing Inv. find Previous one
            InvRow = InvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).Row - 1
        End If
        If .Range("N6").Value = 1 Or MinInvNumb = 0 Or MinInvNumb = .Range("N6").Value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("N3").Value = InvList.Range("A" & InvRow).Value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Prev_Invoice()" & vbNewLine
End Sub

Sub Next_Invoice()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Next_Invoice() @ " & Time
    With shInvoice
        Dim MaxInvNumb As Long
        On Error Resume Next
        MaxInvNumb = Application.WorksheetFunction.Max(InvList.Range("Inv_ID"))
        On Error GoTo 0
        If MaxInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("N6").Value
        If InvNumb = 0 Or .Range("B20").Value = Empty Then 'On New Invoice
            InvRow = InvList.Range("A4").Value  'On Empty Invoice Go to First one created
        Else 'On Existing Inv. find Previous one
            InvRow = InvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).Row + 1
        End If
        If .Range("N6").Value >= MaxInvNumb Then
            MsgBox "You are at the last invoice"
            Exit Sub
        End If
        .Range("N3").Value = InvList.Range("A" & InvRow).Value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Next_Invoice()" & vbNewLine
End Sub

Sub Cacher_Heures()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - Sub Cacher_Heures() @ " & Time
    Range("U65:V66").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Cacher_Heures()" & vbNewLine
End Sub

Sub Montrer_Heures()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Montrer_Heures() @ " & Time
    Range("U65:V66").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Montrer_Heures()" & vbNewLine
End Sub

Sub Retour_Gauche()
    If shInvoice.Range("B26").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Retour_Gauche() @ " & Time
    'ActiveWindow.LargeScroll ToRight:=-1
    'ActiveWindow.SmallScroll Down:=-55
    Range("C1").Select
    Range("P15").Select
    If shInvoice.Range("B26").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Retour_Gauche()" & vbNewLine
End Sub

