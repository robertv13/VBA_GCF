Attribute VB_Name = "modFacture"
Option Explicit
Dim InvRow As Long, InvCol As Long, ItemDBRow As Long, InvItemRow As Long, InvNumb As Long
Dim lastRow As Long, LastItemRow As Long, LastResultRow As Long, ResultRow As Long

Sub Invoice_New()
    If wshFACPrep.Range("B27").value = False Then
        With wshFACPrep
            .Range("B24").value = True
            .Range("K3:L6,O3,O5").ClearContents 'Clear cells for a new Invoice
            .Range("J10:P46,O48,O49,O50,O53").ClearContents
            .Range("O6").value = .Range("FACNextInvoiceNUmber").value 'Paste Invoice ID
            .Range("FACNextInvoiceNUmber").value = .Range("FACNextInvoiceNUmber").value + 1 'Increment Next Invoice ID
            .Range("B20").value = ""
            .Range("B24").value = False
            .Range("B26").value = False
            .Range("B27").value = True 'Set the value to TRUE
        End With
        With wshFACFinale
            .Range("B21,B24:B26").ClearContents
            .Range("A33:F63").ClearContents
            .Range("C65,D65").ClearContents
            .Range("E69:E71,E78").value = 0
            .Range("E28").value = wshFACPrep.Range("O6").value
        End With
        Call TEC_Clear
        wshFACPrep.Range("E4:F4").ClearContents
        wshFACPrep.Range("E4").Select 'Start inputing values for a NEW invoice
    End If
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Le numéro de facture '" & wshFACPrep.Range("O6").value & "' a été assignée"
End Sub

Sub Invoice_SaveUpdate()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_SaveUpdate() @ " & Time
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "B18 (Cust. ID) = " & wshFACPrep.Range("B18").value & "   B20 (Current Inv. Row) = " & wshFACPrep.Range("B20").value
    With wshFACPrep
        'Check For Mandatory Fields - Client
        If .Range("B18").value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que le nom de client n'est pas encore saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        'Check For Mandatory Fields - Date de facture, Date due & Taux horaire
        If .Range("N3").value = Empty Or .Range("N5").value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir saisi la date de facture et le taux horaire avant de sauvegarder la facture"
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que la date de facture et le taux horaire n'ont pas encore été saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        'Determine the row number (InvRow) for InvList
        If .Range("B20").value = Empty Then 'New Invoice
            InvRow = InvList.Range("A99999").End(xlUp).row + 1 'First available row
            wshFACPrep.Range("B20").value = InvRow 'RMV - 2023-10-02 @ 14:39
            InvList.Range("A" & InvRow).value = wshFACPrep.Range("N6").value 'Invoice #
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(10); "Cas A (B20 = '""' ) alors InvRow est établi avec les lignes existantes: InvRow = " & InvRow
        Else 'Existing Invoice
            InvRow = .Range("B20").value 'Set Existing Invoice Row
             If wshFACPrep.Range("B28").value Then Debug.Print Tab(10); "Cas B (B20 <> '""') alors B20 est utilisé: InvRow = " & InvRow
        End If
        If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "B20 (Current Inv. Row) = " & .Range("B20").value & "   B21 (Next Invoice #) = " & .Range("B21").value
        'Load data into InvList (Invoice Header)
        If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Facture # = " & Format(wshFACPrep.Range("N6").value, "000000") & " et Current Inv. Row = " & InvRow & " - Posting; dans; InvoiceListing; """
        'wshFACPrep
        For InvCol = 2 To 5
            InvList.Cells(InvRow, InvCol).value = .Range(InvList.Cells(1, InvCol).value).value 'Save data into Invoice List
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(10); "InvListCol = " & InvCol & "   from wshFACPrep.Cell  = " & InvList.Cells(1, InvCol).value & "   et la valeur = " & .Range(InvList.Cells(1, InvCol).value).value
        Next InvCol
        'wshFACPrep
        For InvCol = 6 To 13
            InvList.Cells(InvRow, InvCol).value = wshFACFinale.Range(InvList.Cells(1, InvCol).value).value 'Save data into Invoice List
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(10); "InvListCol = " & InvCol & "   from wshFACPrep.Cell  = " & InvList.Cells(1, InvCol).value & "   et la valeur = " & wshFACFinale.Range(InvList.Cells(1, InvCol).value).value
        Next InvCol
        'Load data into InvItems (Save/Update Invoice Items) - Columns A, F & G
        LastItemRow = .Range("K46").End(xlUp).row
        If LastItemRow < 10 Then GoTo NoItems
        For InvItemRow = 10 To LastItemRow
            If .Range("O" & InvItemRow).value = "" Then
                ItemDBRow = InvItems.Range("A99999").End(xlUp).row + 1
                .Range("O" & InvItemRow).value = ItemDBRow 'Set Item DB Row
                InvItems.Range("A" & ItemDBRow).value = .Range("N6").value 'Invoice #
                InvItems.Range("F" & ItemDBRow).value = InvItemRow 'Set Invoice Row
                InvItems.Range("G" & ItemDBRow).value = "=Row()"
            Else 'Existing Item
                ItemDBRow = .Range("O" & InvItemRow).value  'Invoice Item Row
            End If
            'Paste 4 columns with one instruction - Columns B, C, D & E
            InvItems.Range("B" & ItemDBRow & ":E" & ItemDBRow).value = .Range("K" & InvItemRow & ":N" & InvItemRow).value 'Save Invoice Item Details
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(15); "Détail (InvItems) - B" & ItemDBRow & " = " & InvItems.Range("B" & ItemDBRow).value
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(20); "  C" & ItemDBRow & " = " & InvItems.Range("C" & ItemDBRow).value & "   D" & ItemDBRow & " = " & InvItems.Range("D" & ItemDBRow).value & "   E" & ItemDBRow & " = " & InvItems.Range("E" & ItemDBRow).value
        Next InvItemRow
NoItems:
        MsgBox "La facture '" & Format(.Range("N6").value, "000000") & "' est enregistrée." & vbNewLine & vbNewLine & "Le total de la facture est " & Trim(Format(.Range("N51").value, "### ##0.00 $")) & " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    End With
    wshFACPrep.Range("B27").value = False
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Total de la facture '" & Format(wshFACPrep.Range("N6").value, "000000") & "' (avant taxes) est de " & Format(wshFACPrep.Range("N51").value, "### ##0.00 $")
Fast_Exit_Sub:
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_SaveUpdate()" & vbNewLine
End Sub

Sub TEC_Clear()

    Dim lastRow As Long
    lastRow = wshFACPrep.Range("D999").End(xlUp).row
    wshFACPrep.Range("D8:I" & lastRow).ClearContents
    
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

Sub TECByClient_FilterAndSort(id As Long)
    
    TEC_Import '2023-12-15 @ 17:02
    
    With wshBaseHours
        Dim lastRow As Long, LastResultRow As Long, ResultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last BaseHours Row
        If lastRow < 2 Then Exit Sub 'Nothing to filter
        Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        .Range("A2:Q" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("T2:W3"), _
            CopyToRange:=.Range("Y2:AL2"), _
            Unique:=True
        LastResultRow = .Range("Y999999").End(xlUp).row
        If LastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If LastResultRow < 4 Then GoTo NoSort
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
            .SetRange wshBaseHours.Range("Y3:AL" & LastResultRow) 'Set Range
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

Sub Invoice_Load()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_Load() @ " & Time
    With wshFACPrep
        If .Range("B20").value = Empty Then
            MsgBox "Veuillez saisir un numéro de facture"
            Exit Sub
        End If
        .Range("B24").value = True 'Set Invoice Load to true
        .Range("Q2,J4:J6,N3:N4,M6:N6,I10:M35,O10:O35").ClearContents
        InvRow = .Range("B20").value
       
        'Assign values from InvList to Invoice worksheet
        For InvCol = 2 To 11 'RMV - 2023-10-01
            If wshFACPrep.Range("B28").value And InvCol <> 3 Then Debug.Print "InvRow = " & InvRow & "   InvCol = " & InvCol & " - " & .Range(InvList.Cells(1, InvCol).value) & " <-- " & InvList.Cells(InvRow, InvCol).value
            If InvCol <> 3 Then .Range(InvList.Cells(1, InvCol).value).value = InvList.Cells(InvRow, InvCol).value 'Load Invoice List Data
        Next InvCol
        'Load Invoice Items
        With InvItems
            lastRow = .Range("A9999").End(xlUp).row
            If lastRow < 4 Then Exit Sub
            If wshFACPrep.Range("B28").value Then Debug.Print "LastRow = " & lastRow & "   Copie de '" & "A3:G" & lastRow & "   Critère: " & .Range("L3").value
            .Range("A3:G" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("L2:L3"), CopyToRange:=.Range("N2:S2"), Unique:=True
            LastResultRow = .Range("V9999").End(xlUp).row
            If wshFACPrep.Range("B28").value Then Debug.Print "Based on column 'V' (InvItems), LastResultRow = " & LastResultRow
            If LastResultRow < 3 Then GoTo NoItems
            For ResultRow = 3 To LastResultRow
                InvItemRow = .Range("R" & ResultRow).value 'Set Invoice Row
                If wshFACPrep.Range("B28").value Then Debug.Print Tab(20); "Invoice Item Row (InvItemRow) = " & InvItemRow & _
                    "   wshFACPrep.Range('K'" & InvItemRow & ")=" & wshFACPrep.Range("K" & InvItemRow).value & " devient " & "InvItems.Range('N'" & ResultRow & ") = " & .Range("N" & ResultRow).value & _
                    "   wshFACPrep.Range('L'" & InvItemRow & ")=" & wshFACPrep.Range("L" & InvItemRow).value & " devient " & "InvItems.Range('O'" & ResultRow & ") = " & .Range("O" & ResultRow).value & _
                    "   wshFACPrep.Range('M'" & InvItemRow & ")=" & wshFACPrep.Range("M" & InvItemRow).value & " devient " & "InvItems.Range('P'" & ResultRow & ") = " & .Range("P" & ResultRow).value & _
                wshFACPrep.Range("K" & InvItemRow & ":M" & InvItemRow).value = .Range("N" & ResultRow & ":P" & ResultRow).value 'Item details
                If wshFACPrep.Range("B28").value Then Debug.Print Tab(30); "wshFACPrep.Range('O'" & InvItemRow & ")=" & wshFACPrep.Range("O" & InvItemRow).value & " devient " & "InvItems.Range('S'" & ResultRow & ") = " & .Range("S" & ResultRow).value
                wshFACPrep.Range("O" & InvItemRow).value = .Range("S" & ResultRow).value  'Set Item DB Row
            Next ResultRow
NoItems:
        End With
        .Range("B24").value = False 'Set Invoice Load To false
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_Load()" & vbNewLine
End Sub

Sub Invoice_Delete()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_Delete() @ " & Time
    With wshFACPrep
        If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Invoice") = vbNo Then Exit Sub
        If .Range("B20").value = Empty Then GoTo NotSaved
        InvRow = .Range("B20").value 'Set Invoice Row
        InvList.Range(InvRow & ":" & InvRow).EntireRow.Delete
        With InvItems
            lastRow = .Range("A99999").End(xlUp).row
            If lastRow < 4 Then Exit Sub
            .Range("A3:J" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("N2:N3"), CopyToRange:=.Range("P2:W2"), Unique:=True
            LastResultRow = .Range("V99999").End(xlUp).row
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
                ItemDBRow = .Range("V" & ResultRow).value 'Set Invoice Database Row
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
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_Delete()" & vbNewLine
End Sub

Sub Previsualisation_PDF() 'RMV - 2023-10-04 @ 15:45
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Previsualisation_PDF() @ " & Time
    wshFACFinale.PrintOut , , , True, True, , , , False
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Previsualisation_PDF()" & vbNewLine
End Sub

Sub Creation_PDF_Email() 'RMV - 2023-10-04 @ 15:46
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Creation_PDF_Email() @ " & Time
    Create_PDF_Email_Sub wshFACPrep.Range("N6").value
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Creation_PDF_Email()" & vbNewLine
End Sub

Sub Create_PDF_Email_Sub(NoFacture As Long)
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Create_PDF_Email_Sub(NoFacture As Long) @ " & Time
    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF et préparation du courriel pour envoyer la facture
    Dim result As Boolean
    result = Create_PDF_Email_Function(NoFacture, "CreateEmail")
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Create_PDF_Email_Sub(NoFacture As Long)" & vbNewLine
End Sub

Function Create_PDF_Email_Function(NoFacture As Long, Optional action As String = "SaveOnly") As Boolean
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Function Create_PDF_Email_Function" & _
        "(NoFacture As Long, Optional action As String = """"SaveOnly"""") As Boolean @ " & Time
    Dim NoFactFormate As String, PathName As String, SaveAs As String

    Application.ScreenUpdating = False

    'Construct the SaveAs filename
    NoFactFormate = Format(NoFacture, "000000")
    PathName = ActiveWorkbook.Path & "\" & "Factures_PDF"
    SaveAs = PathName & "\" & NoFactFormate & ".pdf"

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
    
    'Construct & Displat the Email, allowing the user to modify the Email
    If action = "CreateEmail" Then
        On Error GoTo SaveOnly
        
        Dim outlookApp As Outlook.Application
        Dim myMail As Outlook.MailItem
        
        Set outlookApp = New Outlook.Application
        Set myMail = outlookApp.CreateItem(olMailItem)

        Dim source_file As String
        source_file = "C:\VBA\GC_FISCALITÉ\Factures_PDF\" & NoFactFormate & ".pdf"
        
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
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Create_PDF_Email_Function(NoFacture As Long, Optional action As String = """"SaveOnly"""") As Boolean" & vbNewLine
End Function

Sub Prev_Invoice()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Prev_Invoice() @ " & Time
    With wshFACPrep
        Dim MinInvNumb As Long
        On Error Resume Next
        MinInvNumb = Application.WorksheetFunction.Min(InvList.Range("Inv_ID"))
        On Error GoTo 0
        If MinInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("N6").value
        If InvNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            InvRow = InvList.Range("A99999").End(xlUp).row 'On Empty Invoice Go to last one created
        Else 'On Existing Inv. find Previous one
            InvRow = InvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).row - 1
        End If
        If .Range("N6").value = 1 Or MinInvNumb = 0 Or MinInvNumb = .Range("N6").value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("N3").value = InvList.Range("A" & InvRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Prev_Invoice()" & vbNewLine
End Sub

Sub Next_Invoice()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Next_Invoice() @ " & Time
    With wshFACPrep
        Dim MaxInvNumb As Long
        On Error Resume Next
        MaxInvNumb = Application.WorksheetFunction.Max(InvList.Range("Inv_ID"))
        On Error GoTo 0
        If MaxInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("N6").value
        If InvNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            InvRow = InvList.Range("A4").value  'On Empty Invoice Go to First one created
        Else 'On Existing Inv. find Previous one
            InvRow = InvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).row + 1
        End If
        If .Range("N6").value >= MaxInvNumb Then
            MsgBox "You are at the last invoice"
            Exit Sub
        End If
        .Range("N3").value = InvList.Range("A" & InvRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Next_Invoice()" & vbNewLine
End Sub

Sub Cacher_Heures()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - Sub Cacher_Heures() @ " & Time
    wshFACFinale.Range("C64:D65").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Cacher_Heures()" & vbNewLine
End Sub

Sub Montrer_Heures()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Montrer_Heures() @ " & Time
    wshFACFinale.Range("C64:D65").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Montrer_Heures()" & vbNewLine
End Sub

Sub Goto_Onglet_Preparation_Facture()
    wshFACPrep.Select
    wshFACPrep.Range("C1").Select
End Sub

Sub Goto_Onglet_Facture_Finale()
    wshFACFinale.Select
End Sub

