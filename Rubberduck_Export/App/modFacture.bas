Attribute VB_Name = "modFacture"
Option Explicit
Dim InvRow As Long, InvCol As Long, ItemDBRow As Long, InvItemRow As Long, InvNumb As Long
Dim lastRow As Long, LastItemRow As Long, lastResultRow As Long, ResultRow As Long

Sub Invoice_New() 'Clear contents
    If wshFACPrep.Range("B27").value = False Then
        With wshFACPrep
            .Range("B24").value = True
            .Range("K3:L6,O3,O5").ClearContents 'Clear cells for a new Invoice
            .Range("J10:Q46").ClearContents
            .Range("O6").value = .Range("FACNextInvoiceNumber").value 'Paste Invoice ID
            .Range("FACNextInvoiceNumber").value = .Range("FACNextInvoiceNumber").value + 1 'Increment Next Invoice ID
            
            Call TEC_Clear
            Call ClearAndFixTotalsFormulaFACPrep
            
            .Range("B20").value = ""
            .Range("B24").value = False
            .Range("B26").value = False
            .Range("B27").value = True 'Set the value to TRUE
        End With
        With wshFACFinale
            .Range("B21,B23:C26,F28").ClearContents
            .Range("A33:G62, D65, E65").ClearContents
            .Range("F28").value = wshFACPrep.Range("O6").value 'Invoice #
            .Range("B68:F80").ClearContents
            
            Call ClearAndFixTotalsFormulaFACFinale
        
        End With
        Call TEC_Clear
        wshFACPrep.Range("E4:F4").ClearContents
        wshFACPrep.Range("E4").Select 'Start inputing values for a NEW invoice
    End If
    If wshFACPrep.Range("B28").value Then Debug.Print vbNewLine & "Le numéro de facture '" & wshFACPrep.Range("O6").value & "' a été assignée"
End Sub

Sub Invoice_Load() 'Retrieve an existing invoice - 2023-12-21 @ 10:16
    If wshFACPrep.Range("B28").value Then Debug.Print vbNewLine & "[modFacture] - Now entering Sub Invoice_Load() @ " & Time
    With wshFACPrep
        If wshFACPrep.Range("B20").value = Empty Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
            GoTo NoItems
        End If
        'Could that invoice been cancelled (more than 1 row) ?
        Call InvoiceGetAllTrans(wshFACPrep.Range("O6").value)
        Dim NbTrans As Integer
        NbTrans = .Range("B31").value
        If NbTrans = 0 Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
            GoTo NoItems
        Else
            If NbTrans > 1 Then
                MsgBox "Cette facture a été annulée! Veuillez saisir un numéro de facture VALIDE pour votre recherche"
                GoTo NoItems
            End If
        End If
        If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Loading info from InvList with row # = " & .Range("B20").value
        .Range("B24").value = True 'Set Invoice Load to true
        .Range("S2,E4:F4,K4:L6,O3,K11:O45,Q11:Q45").ClearContents
        wshFACFinale.Range("C34:F63").ClearContents
        Dim InvListRow As Long
        InvListRow = wshFACPrep.Range("B20").value 'InvListRow = Row associated with the invoice
        'Get values from wshFACInvList (header) and enter them in the wshFACPrep - 2023-12-19 @ 08:29
        .Range("O3").value = wshFACInvList.Range("B" & InvListRow).value
        .Range("K3").value = wshFACInvList.Range("D" & InvListRow).value
        .Range("K4").value = wshFACInvList.Range("E" & InvListRow).value
        .Range("K5").value = wshFACInvList.Range("F" & InvListRow).value
        .Range("K6").value = wshFACInvList.Range("G" & InvListRow).value
        'Get values from wshFACInvList (header) and enter them in the wshFACPrep - 2023-12-19 @ 08:29
        wshFACFinale.Range("B21").value = "Le " & Format(wshFACInvList.Range("B" & InvListRow).value, "d mmmm yyyy")
        wshFACFinale.Range("B23").value = wshFACInvList.Range("D" & InvListRow).value
        wshFACFinale.Range("B24").value = wshFACInvList.Range("E" & InvListRow).value
        wshFACFinale.Range("B25").value = wshFACInvList.Range("F" & InvListRow).value
        wshFACFinale.Range("B26").value = wshFACInvList.Range("G" & InvListRow).value
        'Load Invoice Detail Items
        With wshFACInvItems
            Dim lastRow As Long, lastResultRow As Long
            lastRow = .Range("A999999").End(xlUp).row
            If lastRow < 4 Then Exit Sub 'No Item Lines
            .Range("I3").value = wshFACPrep.Range("O6").value
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Invoice Items - From Range '" & "A3:G" & lastRow & "', Critère = '" & .Range("I3").value & "'"
            wshFACFinale.Range("F28").value = wshFACPrep.Range("O6").value 'Invoice #
            'Advanced Filter to get items specific to ONE invoice
            .Range("A3:G" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("I2:I3"), CopyToRange:=.Range("K2:P2"), Unique:=True
            lastResultRow = .Range("O999").End(xlUp).row
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Based on column 'O' (Inv. Row), the LastResultRow = " & lastResultRow
            If lastResultRow < 3 Then GoTo NoItems
            For ResultRow = 3 To lastResultRow
                InvItemRow = .Range("O" & ResultRow).value
                If wshFACPrep.Range("B28").value Then Debug.Print Tab(10); "Loop = " & ResultRow & " - Desc = " & .Range("K" & ResultRow).value & " - Hrs = " & .Range("L" & ResultRow).value
                wshFACPrep.Range("L" & InvItemRow & ":O" & InvItemRow).value = .Range("K" & ResultRow & ":N" & ResultRow).value 'Description, Hours, Rate & Value
                wshFACPrep.Range("Q" & InvItemRow).value = .Range("P" & ResultRow).value  'Set Item DB Row
                wshFACFinale.Range("C" & InvItemRow + 23 & ":F" & InvItemRow + 23).value = .Range("K" & ResultRow & ":N" & ResultRow).value 'Description, Hours, Rate & Value
            Next ResultRow
        End With
        'Proceed with trailer data (Misc. charges & Taxes)
        .Range("M48").value = wshFACInvList.Range("I" & InvListRow).value
        .Range("O48").value = wshFACInvList.Range("J" & InvListRow).value
        .Range("M49").value = wshFACInvList.Range("K" & InvListRow).value
        .Range("O49").value = wshFACInvList.Range("L" & InvListRow).value
        .Range("M50").value = wshFACInvList.Range("M" & InvListRow).value
        .Range("O50").value = wshFACInvList.Range("N" & InvListRow).value
        .Range("O52").value = wshFACInvList.Range("P" & InvListRow).value
        .Range("O53").value = wshFACInvList.Range("R" & InvListRow).value
        .Range("O57").value = wshFACInvList.Range("T" & InvListRow).value
        
NoItems:
    .Range("B24").value = False 'Set Invoice Load To false
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "[modFacture] - Now exiting Sub Invoice_Load()" & vbNewLine
End Sub

Sub InvoiceGetAllTrans(inv As String)

    Application.ScreenUpdating = False

    wshFACPrep.Range("B31").value = 0
   
    With wshFACInvList
        Dim lastRow As Long, lastResultRow As Long, ResultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last wshFACInvList Row
        If lastRow < 4 Then GoTo Done '3 rows of Header - Nothing to search/filter
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        .Range("V3").value = wshFACPrep.Range("O6").value
        'Advanced Filter setup
        .Range("A3:T" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("V2:V3"), _
            CopyToRange:=.Range("X2:AQ2"), _
            Unique:=True
        lastResultRow = .Range("X999").End(xlUp).row 'How many rows trans for that invoice
        If lastResultRow < 3 Then
            GoTo Done
        End If
'        With .Sort
'            .SortFields.Clear
'            .SortFields.Add Key:=wshFACInvList.Range("X2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based Invoice Number
'            .SortFields.Add Key:=wshBaseHours.Range("Y3"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based On TEC_ID
'            .SetRange wshFACInvList.Range("X2:AQ" & lastResultRow) 'Set Range
'            .Apply 'Apply Sort
'         End With
         wshFACPrep.Range("B31").value = lastResultRow - 2 'Remove Header rows from row count
Done:
    End With
    Application.ScreenUpdating = True

End Sub

Sub ClearAndFixTotalsFormulaFACPrep()

    Application.EnableEvents = False
    
    With wshFACPrep
        .Range("J46:P59").ClearContents
        
        Call SetLabels(.Range("K47"), "FAC_Label_SubTotal_1")
        Call SetLabels(.Range("K51"), "FAC_Label_SubTotal_2")
        Call SetLabels(.Range("K52"), "FAC_Label_TPS")
        Call SetLabels(.Range("K53"), "FAC_Label_TVQ")
        Call SetLabels(.Range("K55"), "FAC_Label_GrandTotal")
        Call SetLabels(.Range("K57"), "FAC_Label_Deposit")
        Call SetLabels(.Range("K59"), "FAC_Label_AmountDue")
        
        .Range("O47").Formula = "=SUM(O11:O45)" 'Fees sub-total
        .Range("O47").Font.Bold = True
        
        .Range("M48").value = wshAdmin.Range("FAC_Label_Frais_1").value 'Misc. # 1 - Descr.
        .Range("O48").value = "" 'Misc. # 1 - Amount
        .Range("M49").value = wshAdmin.Range("FAC_Label_Frais_2").value 'Misc. # 2 - Descr.
        .Range("O49").value = "" 'Misc. # 2 - Amount
        .Range("M50").value = wshAdmin.Range("FAC_Label_Frais_3").value 'Misc. # 3 - Descr.
        .Range("O50").value = "" 'Misc. # 3 - Amount
        
        .Range("O51").Formula = "=sum(O47:O50)" 'Sub-total
        .Range("O51").Font.Bold = True
        
        .Range("N52").value = wshFACPrep.Range("B29").value 'GST Rate
        .Range("N52").NumberFormat = "0.00%"
        .Range("O52").Formula = "=round(o51*n52,2)" 'GST Amnt
        .Range("N53").value = wshFACPrep.Range("B30").value 'PST Rate
        .Range("N53").NumberFormat = "0.000%"
        .Range("O53").Formula = "=round(o51*n53,2)" 'GST Amnt
        .Range("O55").Formula = "=sum(o51:o54)" 'Grand Total"
        .Range("O57").value = "" 'Deposit Amount
        .Range("O59").Formula = "=O55-O57" 'Deposit Amount
        
    End With
    
    Application.EnableEvents = True

End Sub

Sub ClearAndFixTotalsFormulaFACFinale()

    Application.EnableEvents = False
    
    With wshFACFinale
        Call SetLabels(.Range("B68"), "FAC_Label_SubTotal_1")
        Call SetLabels(.Range("B72"), "FAC_Label_SubTotal_2")
        Call SetLabels(.Range("B73"), "FAC_Label_TPS")
        Call SetLabels(.Range("B74"), "FAC_Label_TVQ")
        Call SetLabels(.Range("B76"), "FAC_Label_GrandTotal")
        Call SetLabels(.Range("B78"), "FAC_Label_Deposit")
        Call SetLabels(.Range("B80"), "FAC_Label_AmountDue")

        'Fix formulas to calculate amounts & Copy cells from FAC_Préparation
        .Range("F68").Formula = "=SUM(F33:F62)" 'Fees Sub-Total
        .Range("C69").Formula = "='" & wshFACPrep.Name & "'!M48" 'Misc. Amount # 1 - Description
        .Range("F69").Formula = "='" & wshFACPrep.Name & "'!O48" 'Misc. Amount # 1
        .Range("C70").Formula = "='" & wshFACPrep.Name & "'!M49" 'Misc. Amount # 2 - Description
        .Range("F70").Formula = "='" & wshFACPrep.Name & "'!O49" 'Misc. Amount # 2
        .Range("C71").Formula = "='" & wshFACPrep.Name & "'!M50" 'Misc. Amount # 3 - Description
        .Range("F71").Formula = "='" & wshFACPrep.Name & "'!O50" 'Misc. Amount # 3
        .Range("F72").Formula = "=F68+F69+F70+F71" 'Sub-Total
        .Range("D73").Formula = "='" & wshFACPrep.Name & "'!N52" 'GST Rate
        .Range("F73").Formula = "='" & wshFACPrep.Name & "'!O52" 'GST Amount
        .Range("D74").Formula = "='" & wshFACPrep.Name & "'!N53" 'PST Rate
        .Range("F74").Formula = "='" & wshFACPrep.Name & "'!O53" 'PST Amount
        .Range("F76").Formula = "=F72+F73+F74" 'Total including taxes
        .Range("F78").Formula = "='" & wshFACPrep.Name & "'!O57" 'Deposit Amount
        .Range("F80").Formula = "=F76-F78" 'Total due on that invoice
    End With
    
    Application.EnableEvents = True
    
End Sub

Sub SetLabels(r As Range, l As String)

    r.value = wshAdmin.Range(l).value
    If wshAdmin.Range(l & "_Bold").value = "OUI" Then r.Font.Bold = True

End Sub
Sub MiscCharges()

    ActiveWindow.SmallScroll Down:=14
    wshFACPrep.Range("O48").Select 'Misc Amount 1
    
End Sub
Sub Invoice_SaveUpdate() '2023-12-20 @ 11:31
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFacture] - Sub Invoice_SaveUpdate() @ " & Time
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "B18 (Cust. ID) = " & wshFACPrep.Range("B18").value
    With wshFACPrep
        'Check For Mandatory Fields - Client
        If .Range("B18").value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que le nom de client n'est pas encore saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        'Check For Mandatory Fields - Date de facture
        If .Range("O3").value = Empty Or Len(Trim(.Range("O6").value)) <> 8 Then
            MsgBox "Veuillez vous assurer d'avoir saisi la date de facture AVANT de sauvegarder la facture"
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que la date de facture et/ou le numéro de facture n'ont pas encore été saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        'Determine the row number (InvListRow) for InvList
        If wshFACPrep.Range("B20").value = Empty Then 'New Invoice
            Dim InvListRow As Long
            InvListRow = wshFACInvList.Range("A99999").End(xlUp).row + 1 'First available row
            wshFACPrep.Range("B20").value = InvListRow
            wshFACInvList.Range("A" & InvListRow).value = wshFACPrep.Range("O6").value 'Invoice #
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Cas A (B20 = '' ) alors InvListRow est établi selon les lignes existantes: InvListRow = " & InvListRow
        Else 'Existing Invoice
            InvListRow = .Range("B20").value 'Set Existing Invoice Row
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Cas B (B20 <> '') alors B20 est utilisé - InvListRow = " & InvListRow
        End If
        If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "B20 (Current InvListRow) = " & .Range("B20").value & "   B22 (Search InvListRow) = " & .Range("B22").value
        'Load data into wshFACInvList (Invoice Header)
        If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Facture # = " & wshFACPrep.Range("O6").value & " et Current Inv. Row = " & InvListRow & " - pour posting dans InvoiceListing"
        
        wshFACInvList.Range("B" & InvListRow).value = .Range("O3").value 'Date
        wshFACInvList.Range("C" & InvListRow).value = .Range("B18").value 'Client_ID
        wshFACInvList.Range("D" & InvListRow).value = .Range("K3").value 'Care of
        wshFACInvList.Range("E" & InvListRow).value = .Range("K4").value 'Client Name
        wshFACInvList.Range("F" & InvListRow).value = .Range("K5").value 'Client Address
        wshFACInvList.Range("G" & InvListRow).value = .Range("K6").value 'City, Prov & Postal Code

        wshFACInvList.Range("H" & InvListRow).value = wshFACFinale.Range("F68").value 'Fees sub-total
        wshFACInvList.Range("I" & InvListRow).value = wshFACFinale.Range("C69").value 'Misc. # 1 - Desc
        wshFACInvList.Range("J" & InvListRow).value = wshFACFinale.Range("F69").value 'Misc. # 1
        wshFACInvList.Range("K" & InvListRow).value = wshFACFinale.Range("C70").value 'Misc. # 2 - Desc
        wshFACInvList.Range("L" & InvListRow).value = wshFACFinale.Range("F70").value 'Misc. # 2
        wshFACInvList.Range("M" & InvListRow).value = wshFACFinale.Range("C71").value 'Misc. # 3 - Desc
        wshFACInvList.Range("N" & InvListRow).value = wshFACFinale.Range("F71").value 'Misc. # 3
        
        wshFACInvList.Range("O" & InvListRow).value = wshFACFinale.Range("D73").value 'GST Rate
        wshFACInvList.Range("O" & InvListRow).NumberFormat = "0.00%"
        wshFACInvList.Range("P" & InvListRow).value = wshFACFinale.Range("F73").value 'GST $
        wshFACInvList.Range("Q" & InvListRow).value = wshFACFinale.Range("D74").value 'PST Rate
        wshFACInvList.Range("Q" & InvListRow).NumberFormat = "0.000%"
        wshFACInvList.Range("R" & InvListRow).value = wshFACFinale.Range("F74").value 'GST $
        wshFACInvList.Range("S" & InvListRow).value = wshFACFinale.Range("F76").value 'Grand Total
        wshFACInvList.Range("T" & InvListRow).value = wshFACFinale.Range("F78").value 'Deposit received
        
        'Load data into wshInvItems (Save/Update Invoice Items) - Columns A, F & G - TO-DO_RMV - 2023-12-17 @ 15:38 - Duplicate entries !!!
        LastItemRow = .Range("L46").End(xlUp).row
        If LastItemRow < 11 Then GoTo NoItems
        For InvItemRow = 11 To LastItemRow
            If .Range("Q" & InvItemRow).value = "" Then
                ItemDBRow = wshFACInvItems.Range("A99999").End(xlUp).row + 1
                .Range("Q" & InvItemRow).value = ItemDBRow 'Set Item DB Row
                wshFACInvItems.Range("A" & ItemDBRow).value = .Range("O6").value 'Invoice #
                wshFACInvItems.Range("F" & ItemDBRow).value = InvItemRow 'Set Invoice Row
                wshFACInvItems.Range("G" & ItemDBRow).value = "=Row()"
            Else 'Existing Item
                ItemDBRow = .Range("Q" & InvItemRow).value  'Invoice Item Row
            End If
            'Paste 4 columns with one instruction - Columns B, C, D & E
            wshFACInvItems.Range("B" & ItemDBRow & ":E" & ItemDBRow).value = .Range("L" & InvItemRow & ":O" & InvItemRow).value 'Save Invoice Item Details
        Next InvItemRow
NoItems:
        MsgBox "La facture '" & .Range("O6").value & "' est enregistrée." & vbNewLine & vbNewLine & "Le total de la facture est " & Trim(Format(.Range("O51").value, "### ##0.00 $")) & " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    End With
    wshFACPrep.Range("B27").value = False
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Total de la facture '"; wshFACPrep.Range("O6") & "' (avant taxes) est de " & Format(wshFACPrep.Range("O51").value, "### ##0.00 $")
Fast_Exit_Sub:
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFacture] - Sub Invoice_SaveUpdate()" & vbNewLine
    
'    Dim myShape As Shape
'    Set myShape = ActiveSheet.Shapes("Rectangle 18")
    'Deactivate the shape
    'myShape.OLEFormat.Object.Enabled = False
    Call FromFAC2GL(InvListRow)
    
End Sub

Sub ClientChange(ClientName As String)

    wshFACPrep.Range("B18").value = GetID_FromClientName(ClientName)
    
    With wshFACPrep
        .Range("K3").value = "Monsieur Robert M. Vigneault"
        .Range("K4").value = ClientName
        .Range("K5").value = "15 chemin des Mésanges" 'Address 1
        .Range("K6").value = "Mansonville, QC  J0E 1X0" 'Ville, Province & Code postal
    End With
    With wshFACFinale
        .Range("B21").value = "Le " & Format(wshFACPrep.Range("O3").value, "d mmmm yyyy")
        .Range("B23").value = wshFACPrep.Range("K3").value 'Contact from wshFACPrep
        .Range("B24").value = wshFACPrep.Range("K4").value 'Client from wshFACPrep
        .Range("B25").value = wshFACPrep.Range("K5").value 'Address 1 from wshFACPrep
        .Range("B26").value = wshFACPrep.Range("K6").value
    End With
    
    TEC_Load
    
    wshFACPrep.Range("O3").Select 'Move on to Invoice Date

End Sub

Sub DateChange(d As String)

    If InStr(1, wshFACPrep.Range("O6").value, "-") = 0 Then
        Dim y As String
        y = Right(Year(d), 2)
        wshFACPrep.Range("O6").value = y & "-" & wshFACPrep.Range("O6").value
        wshFACFinale.Range("F28").value = wshFACPrep.Range("O6").value
    End If
    wshFACFinale.Range("B21").value = "Le " & Format(d, "d mmmm yyyy")
    
    'Must Get GST & PST rates and store them in wshFACPrep 'B' column
    Dim DateTaxRates As Date
    DateTaxRates = d
    wshFACPrep.Range("B29").value = GetTaxRate(DateTaxRates, "F")
    wshFACPrep.Range("B30").value = GetTaxRate(DateTaxRates, "P")
        
    wshFACPrep.Range("L11").Select 'Move on to Services Entry
    
End Sub

Sub TEC_Clear()

    Dim lastRow As Long
    lastRow = wshFACPrep.Range("D999").End(xlUp).row
    If lastRow > 7 Then wshFACPrep.Range("D8:I" & lastRow).ClearContents
    
End Sub

Sub TEC_Load()

    'Set Criteria, before Filtering entries
    wshBaseHours.Range("T3").value = wshFACPrep.Range("B18").value
    TECByClient_FilterAndSort (wshFACPrep.Range("B18").value)
    
    'Reset Criteria, after Filtering entries
    wshBaseHours.Range("T3").value = ""
    
    CopyFromFilteredEntriesToFACPrep
    
    wshFACPrep.Range("O3").Select
    
End Sub

Sub TECByClient_FilterAndSort(id As Long) 'RMV-2023-12-21 @ 11:00
    
    Application.ScreenUpdating = False

    Call TEC_Import '2023-12-15 @ 17:02
    
    With wshBaseHours
        Dim lastRow As Long, lastResultRow As Long, ResultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last BaseHours Row
        If lastRow < 3 Then Exit Sub 'Nothing to filter
        Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        .Range("A2:Q" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("T2:W3"), _
            CopyToRange:=.Range("Y2:AL2"), _
            Unique:=True
        lastResultRow = .Range("Y999999").End(xlUp).row
        If lastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If lastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshBaseHours.Range("AA3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add Key:=wshBaseHours.Range("Y3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshBaseHours.Range("Y3:AL" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
NoSort:
    End With
    Application.ScreenUpdating = True
End Sub

Sub CopyFromFilteredEntriesToFACPrep()

    Dim lastRow As Long
    lastRow = wshBaseHours.Range("Y99999").End(xlUp).row
    Dim row As Long
    row = 8
    
    Dim i As Integer
    With wshBaseHours
        For i = 3 To lastRow
            If .Range("AH" & i).value = False And .Range("AJ" & i).value = False Then
                wshFACPrep.Range("D" & row).value = .Range("AA" & i).value 'Date
                wshFACPrep.Range("E" & row).value = .Range("Z" & i).value 'Date
                wshFACPrep.Range("F" & row).value = .Range("AC" & i).value 'Description
                wshFACPrep.Range("G" & row).value = .Range("AD" & i).value 'Heures
                wshFACPrep.Range("H" & row).value = .Range("AH" & i).value 'Facturée ou pas
                wshFACPrep.Range("I" & row).value = .Range("Y" & i).value 'TEC_ID
                row = row + 1
            End If
        Next i
    End With
End Sub



Sub Invoice_Delete()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFacture] - Sub Invoice_Delete() @ " & Time
    With wshFACPrep
        If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Invoice") = vbNo Then Exit Sub
        If .Range("B20").value = Empty Then GoTo NotSaved
        InvRow = .Range("B20").value 'Set Invoice Row
        wshFACInvList.Range(InvRow & ":" & InvRow).EntireRow.Delete
        With InvItems
            lastRow = .Range("A99999").End(xlUp).row
            If lastRow < 4 Then Exit Sub
            .Range("A3:J" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("N2:N3"), CopyToRange:=.Range("P2:W2"), Unique:=True
            lastResultRow = .Range("V99999").End(xlUp).row
            If lastResultRow < 3 Then GoTo NoItems
    '        If LastResultRow < 4 Then GoTo SkipSort
    '        'Sort Rows Descending
    '         With .Sort
    '         .SortFields.Clear
    '         .SortFields.Add Key:=wshFACInvItems.Range("W3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal  'Sort
    '         .SetRange wshFACInvItems.Range("P3:W" & LastResultRow) 'Set Range
    '         .Apply 'Apply Sort
    '         End With
SkipSort:
            For ResultRow = 3 To lastResultRow
                ItemDBRow = .Range("V" & ResultRow).value 'Set Invoice Database Row
                .Range("A" & ItemDBRow & ":J" & ItemDBRow).ClearContents 'Clear Fields (deleting creates issues with results
            Next ResultRow
            'Resort DB to remove spaces
            With .Sort
                .SortFields.Clear
                .SortFields.Add Key:=wshFACInvItems.Range("A4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal  'Sort
                .SetRange wshFACInvItems.Range("A4:J" & lastResultRow) 'Set Range
                .Apply 'Apply Sort
            End With
        End With
NoItems:
NotSaved:
    Call Invoice_New 'Add New Invoice
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFacture] - Sub Invoice_Delete()" & vbNewLine
End Sub

Sub Previsualisation_PDF() 'RMV - 2023-12-17 @ 14:33

    wshFACFinale.PrintOut , , , True, True, , , , False
    
End Sub

Sub Creation_PDF_Email() 'RMV - 2023-12-17 @ 14:35
    
    Call Create_PDF_Email_Sub(wshFACPrep.Range("O6").value)

End Sub

Sub Create_PDF_Email_Sub(NoFacture As String)
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFacture] - Create_PDF_Email_Sub(NoFacture As String) @ " & Time
    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF et préparation du courriel pour envoyer la facture
    Dim result As Boolean
    result = Create_PDF_Email_Function(NoFacture, "CreateEmail")
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFacture] - Create_PDF_Email_Sub(NoFacture As String)" & vbNewLine
End Sub

Function Create_PDF_Email_Function(NoFacture As String, Optional action As String = "SaveOnly") As Boolean
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Now entering - [modFacture] - Function Create_PDF_Email_Function" & _
        "(NoFacture As Long, Optional action As String = """"SaveOnly"""") As Boolean @ " & Time
    Dim SaveAs As String

    Application.ScreenUpdating = False

    'Construct the SaveAs filename
    'NoFactFormate = Format(NoFacture, "000000")
    SaveAs = wshAdmin.Range("FolderPDFInvoice").value & Application.PathSeparator & _
                     NoFacture & ".pdf" '2023-12-19 @ 07:28

    'Set Print Quality
    On Error Resume Next
    ActiveSheet.PageSetup.PrintQuality = 600
    Err.Clear
    On Error GoTo 0

    'Adjust Document Properties - 2023-10-06 @ 09:54
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.1)
        .RightMargin = Application.InchesToPoints(0.1)
        .TopMargin = Application.InchesToPoints(0.1)
        .BottomMargin = Application.InchesToPoints(0.1)
    End With
    
    'Create the PDF file and Save It
    On Error GoTo RefLibError
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=SaveAs, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    On Error GoTo 0
    
    'Construct & Display the Email, allowing the user to modify the Email
    If action = "CreateEmail" Then
        On Error GoTo SaveOnly
        
        Dim outlookApp As Outlook.Application
        Dim myMail As Outlook.MailItem
        
        Set outlookApp = New Outlook.Application
        Set myMail = outlookApp.CreateItem(olMailItem)

        Dim source_file As String
        source_file = wshAdmin.Range("FolderPDFInvoice").value & Application.PathSeparator & _
                      NoFacture & ".pdf" '2023-12-19 @ 07:22
        
        With myMail
            .To = "robertv13@hotmail.com"
            .CC = "robertv13@me.com"
            .BCC = "robertv13@gmail.com"
            .Subject = "TEST - GC FISCALITÉ INC. - Facturation - TEST"
            .Body = "Bonjour," & vbNewLine & vbNewLine & "Vous trouverez ci-joint notre note d'honoraires." & _
                vbNewLine & vbNewLine & "Merci" & vbNewLine & vbNewLine & vbNewLine & "Guillaume Charron, CPA, CA, M. Fisc." & _
                vbNewLine & "Président"
            .Attachments.Add source_file
           
            .Display 'Affiche le courriel, ce qui permet de corriger AVANT l'envoi
            'myMail.Send
        End With
        
        Set outlookApp = Nothing
        Set myMail = Nothing

    End If
    
SaveOnly:
    Create_PDF_Email_Function = True 'Return value
    GoTo EndMacro
    
RefLibError:
    MsgBox "Incapable de préparer le courriel. La librairie n'est pas disponible"
    Create_PDF_Email_Function = False 'Return value

EndMacro:
    Application.ScreenUpdating = True
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Now exiting  - [modFacture] - Create_PDF_Email_Function(NoFacture As Long, Optional action As String = """"SaveOnly"""") As Boolean" & vbNewLine
End Function

Sub Prev_Invoice() 'TO-DO-RMV 2023-12-17
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFacture] - Sub Prev_Invoice() @ " & Time
    With wshFACPrep
        Dim MinInvNumb As Long
        On Error Resume Next
        MinInvNumb = Application.WorksheetFunction.Min(wshFACInvList.Range("Inv_ID"))
        On Error GoTo 0
        If MinInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("N6").value
        If InvNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            InvRow = wshFACInvList.Range("A99999").End(xlUp).row 'On Empty Invoice Go to last one created
        Else 'On Existing Inv. find Previous one
            InvRow = wshFACInvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).row - 1
        End If
        If .Range("N6").value = 1 Or MinInvNumb = 0 Or MinInvNumb = .Range("N6").value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFACInvList.Range("A" & InvRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFacture] - Sub Prev_Invoice()" & vbNewLine
End Sub

Sub Next_Invoice() 'TO-DO-RMV 2023-12-17
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFacture] - Sub Next_Invoice() @ " & Time
    With wshFACPrep
        Dim MaxInvNumb As Long
        On Error Resume Next
        MaxInvNumb = Application.WorksheetFunction.Max(wshFACInvList.Range("Inv_ID"))
        On Error GoTo 0
        If MaxInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("N6").value
        If InvNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            InvRow = wshFACInvList.Range("A4").value  'On Empty Invoice Go to First one created
        Else 'On Existing Inv. find Previous one
            InvRow = wshFACInvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).row + 1
        End If
        If .Range("N6").value >= MaxInvNumb Then
            MsgBox "You are at the last invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFACInvList.Range("A" & InvRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFacture] - Sub Next_Invoice()" & vbNewLine
End Sub

Sub Cacher_Heures()
    With wshFACFinale.Range("D34:F66")
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
End Sub

Sub Montrer_Heures()
    With wshFACFinale.Range("D34:F66")
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
End Sub

Sub Goto_Onglet_Preparation_Facture()
    wshFACPrep.Select
    wshFACPrep.Range("C1").Select
End Sub

Sub ExportAllFacInvList() '2023-12-21 @ 14:36
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceRange As Range

    Application.ScreenUpdating = False
    
    'Work with the source range
    Set wsSource = wshFACInvList
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Range("A99999").End(xlUp).row
    wsSource.Range("A4:T" & lastUsedRow).Copy

    'Open the target workbook
    Workbooks.Open Filename:=wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"

    'Set references to the target workbook and target worksheet
    Set wb = Workbooks("GCF_BD_Sortie.xlsx")
    Set wsTarget = wb.Sheets("FACTURES")

    'PasteSpecial directly to the target range
    wsTarget.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    wb.Close SaveChanges:=True
    
    Application.ScreenUpdating = True
    
End Sub

Sub FromFAC2GL(r As Long) '2023-12-21 @ 22:33

    Dim Montant As Double
    Dim DateFact As Date
    Dim NoFacture As String
    Dim nomClient As String
    
    NoFacture = wshFACInvList.Range("A" & r).value
    DateFact = wshFACInvList.Range("B" & r).value
    nomClient = wshFACInvList.Range("E" & r).value
    
    Dim Rng As Range
    Set Rng = wshGLFACTrans.Range("C2:C999999")
    Dim newID As Long
    newID = WorksheetFunction.Max(Rng) + 1

    'AR amount
    Montant = wshFACInvList.Range("S" & r).value
    If Montant Then Call GLPost(Montant, newID, "1100", "Comptes Clients", DateFact)
    
    'Professionnal Fees
    Montant = -wshFACInvList.Range("H" & r).value
    If Montant Then Call GLPost(Montant, newID, "4000", "Revenus", DateFact)
    
    'Miscellaneous Amount # 1
    Montant = -wshFACInvList.Range("J" & r).value
    If Montant Then Call GLPost(Montant, newID, "5009", "Frais divers # 1", DateFact)
    
    'Miscellaneous Amount # 2
    Montant = -wshFACInvList.Range("L" & r).value
    If Montant Then Call GLPost(Montant, newID, "5008", "Frais divers # 2", DateFact)
    
    'Miscellaneous Amount # 3
    Montant = -wshFACInvList.Range("N" & r).value
    If Montant Then Call GLPost(Montant, newID, "5002", "Frais divers # 3", DateFact)
    
    'TPS à payer
    Montant = -wshFACInvList.Range("P" & r).value
    If Montant Then Call GLPost(Montant, newID, "2200", "TPS à payer", DateFact)
    
    'TVQ à payer
    Montant = -wshFACInvList.Range("R" & r).value
    If Montant Then Call GLPost(Montant, newID, "2201", "TVQ à payer", DateFact)
    
    Call GLPost(0, newID, "", NoFacture + "-" & nomClient, DateFact)
    Call GLPost(0, newID, "", "", DateFact)
    
    Call AdjustJETrans(newID)
    
End Sub

Sub GLPost(m As Double, noEJ, GL As String, GLDesc As String, d As Date)

    Dim rowGLTrans As Long, maxID As Double, newID As Long
    'Détermine la prochaine ligne disponible dans la table
    rowGLTrans = wshGLFACTrans.Range("C999999").End(xlUp).row + 1  'Last Used + 1 = First Empty Row

    wshGLFACTrans.Range("C" & rowGLTrans).value = noEJ
    wshGLFACTrans.Range("D" & rowGLTrans).value = d
    wshGLFACTrans.Range("E" & rowGLTrans).value = noEJ
    wshGLFACTrans.Range("F" & rowGLTrans).value = "Facturation"
    wshGLFACTrans.Range("G" & rowGLTrans).value = GL
    wshGLFACTrans.Range("H" & rowGLTrans).value = GLDesc
    If m > 0 Then
        wshGLFACTrans.Range("I" & rowGLTrans).value = m
    ElseIf m < 0 Then
        wshGLFACTrans.Range("J" & rowGLTrans).value = -m
    End If
    wshGLFACTrans.Range("K" & rowGLTrans).value = ""
    wshGLFACTrans.Range("L" & rowGLTrans).Formula = "=ROW()"

End Sub

Sub AdjustJETrans(JENumber As Long) '2023-12-22 @ 08:18
    
    Dim firstRow As Long, lastRow As Long, r As Long
    Dim nrJE_All As Range
    Set nrJE_All = Range("nrJE_All")
    firstRow = Application.WorksheetFunction.Match(JENumber, nrJE_All, 0) + 1
    r = firstRow
    
    'Determine the last row for a given Journal Entry
    Do While wshGLFACTrans.Cells(r, 3).value = JENumber
        r = r + 1
    Loop
    lastRow = r - 1
    
    With wshGLFACTrans
        'Les lignes subséquentes sont en police blanche...
        .Range("D" & (firstRow + 1) & ":F" & lastRow).Font.Color = vbWhite
        
        'We adjust Numeric Formats for the amounts
        .Range("I" & firstRow & ":J" & (lastRow - 2)).NumberFormat = "#,###,##0.00 $"
        
        'Ajoute des bordures (cadre extérieur) à l'ensemble des lignes de l'écriture
        Dim r1 As Range
        Set r1 = .Range("D" & firstRow & ":K" & (lastRow - 1))
        r1.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=vbBlack
        
        With .Range("H" & (lastRow - 1) & ":K" & (lastRow - 1))
            .Merge
            .HorizontalAlignment = xlLeft
            .Font.Italic = True
            .Font.Bold = True
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
            .Borders(xlInsideVertical).LineStyle = xlNone
        End With
    End With
End Sub

Sub LoopUntilColumnChange()
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim targetColumn As Long
    Dim targetValue As Variant
    
    ' Set your worksheet
    Set ws = ThisWorkbook.Sheets("YourSheetName") ' Replace with your actual sheet name
    
    ' Set the target column and value
    targetColumn = 3 ' Change this to your target column number (e.g., column C)
    targetValue = "YourTargetValue" ' Change this to your target value
    
    ' Initialize the starting row
    currentRow = 1
    
    ' Loop until the target column changes its value
    Do While ws.Cells(currentRow, targetColumn).value <> targetValue
        ' Your code for each row goes here
        ' You can access cell values using ws.Cells(currentRow, ColumnNumber)
        
        ' Move to the next row
        currentRow = currentRow + 1
    Loop
End Sub


