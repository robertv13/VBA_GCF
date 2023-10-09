Attribute VB_Name = "Invoice_Macros"
Option Explicit
Dim InvRow As Long, InvCol As Long, ItemDBRow As Long, InvItemRow As Long, InvNumb As Long
Dim LastRow As Long, LastItemRow As Long, LastResultRow As Long, ResultRow As Long

Sub Invoice_New()
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_New() @ " & Time
    If shInvoice.Range("B27").Value = False Then
        With shInvoice
            .Range("B24").Value = True
            .Range("J3,J4:K4,J5,J6,N3,N5").ClearContents 'Clear cells for a new Invoice
            .Range("I10:N46,O10:O46,N48,N49,N50,N53").ClearContents
            .Range("N6").Value = .Range("B21").Value 'Paste Invoice ID
            .Range("B21").Value = .Range("B21").Value + 1 'Increment Next Invoice ID
            .Range("B20").Value = ""
            .Range("B26").Value = False
            .Range("B27").Value = True
            .Range("B24").Value = False
        End With
        With shFactureFinale
            .Range("B21,B24:B26").ClearContents
            .Range("A33:F63").ClearContents
            .Range("C65,D65").ClearContents
            .Range("E68:E71,E78").Value = 0
            .Range("E28").Value = shInvoice.Range("N6").Value
        End With
    End If
    If shInvoice.Range("B28").Value Then Debug.Print Tab(5); "Le numéro de facture '" & shInvoice.Range("N6").Value & "' a été assignée"
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_New()" & vbNewLine
End Sub

Sub Invoice_SaveUpdate()
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_SaveUpdate() @ " & Time
    If shInvoice.Range("B28").Value Then Debug.Print Tab(5); "B18 (Cust. ID) = " & shInvoice.Range("B18").Value & "   B20 (Current Inv. Row) = " & shInvoice.Range("B20").Value
    With shInvoice
        'Check For Mandatory Fields - Client
        If .Range("B18").Value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            If shInvoice.Range("B28").Value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que le nom de client n'est pas encore saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        'Check For Mandatory Fields - Date de facture, Date due & Taux horaire
        If .Range("N3").Value = Empty Or .Range("N5").Value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir saisi la date de facture et le taux horaire avant de sauvegarder la facture"
            If shInvoice.Range("B28").Value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que la date de facture et le taux horaire n'ont pas encore été saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        'Determine the row number (InvRow) for InvList
        If .Range("B20").Value = Empty Then 'New Invoice
            InvRow = InvList.Range("A99999").End(xlUp).Row + 1 'First available row
            shInvoice.Range("B20").Value = InvRow 'RMV - 2023-10-02 @ 14:39
            InvList.Range("A" & InvRow).Value = shInvoice.Range("N6").Value 'Invoice #
            If shInvoice.Range("B28").Value Then Debug.Print Tab(10); "Cas A (B20 = '""' ) alors InvRow est établi avec les lignes existantes: InvRow = " & InvRow
        Else 'Existing Invoice
            InvRow = .Range("B20").Value 'Set Existing Invoice Row
             If shInvoice.Range("B28").Value Then Debug.Print Tab(10); "Cas B (B20 <> '""') alors B20 est utilisé: InvRow = " & InvRow
        End If
        If shInvoice.Range("B28").Value Then Debug.Print Tab(5); "B20 (Current Inv. Row) = " & .Range("B20").Value & "   B21 (Next Invoice #) = " & .Range("B21").Value
        'Load data into InvList (Invoice Header)
        If shInvoice.Range("B28").Value Then Debug.Print Tab(5); "Facture # = " & Format(shInvoice.Range("N6").Value, "000000") & " et Current Inv. Row = " & InvRow & " - Posting; dans; InvoiceListing; """
        'shInvoice
        For InvCol = 2 To 5
            InvList.Cells(InvRow, InvCol).Value = .Range(InvList.Cells(1, InvCol).Value).Value 'Save data into Invoice List
            If shInvoice.Range("B28").Value Then Debug.Print Tab(10); "InvListCol = " & InvCol & "   from shInvoice.Cell  = " & InvList.Cells(1, InvCol).Value & "   et la valeur = " & .Range(InvList.Cells(1, InvCol).Value).Value
        Next InvCol
        'shInvoice
        For InvCol = 6 To 13
            InvList.Cells(InvRow, InvCol).Value = shFactureFinale.Range(InvList.Cells(1, InvCol).Value).Value 'Save data into Invoice List
            If shInvoice.Range("B28").Value Then Debug.Print Tab(10); "InvListCol = " & InvCol & "   from shInvoice.Cell  = " & InvList.Cells(1, InvCol).Value & "   et la valeur = " & shFactureFinale.Range(InvList.Cells(1, InvCol).Value).Value
        Next InvCol
        'Load data into InvItems (Save/Update Invoice Items) - Columns A, F & G
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
            'Paste 4 columns with one instruction - Columns B, C, D & E
            InvItems.Range("B" & ItemDBRow & ":E" & ItemDBRow).Value = .Range("K" & InvItemRow & ":N" & InvItemRow).Value 'Save Invoice Item Details
            If shInvoice.Range("B28").Value Then Debug.Print Tab(15); "Détail (InvItems) - B" & ItemDBRow & " = " & InvItems.Range("B" & ItemDBRow).Value
            If shInvoice.Range("B28").Value Then Debug.Print Tab(20); "  C" & ItemDBRow & " = " & InvItems.Range("C" & ItemDBRow).Value & "   D" & ItemDBRow & " = " & InvItems.Range("D" & ItemDBRow).Value & "   E" & ItemDBRow & " = " & InvItems.Range("E" & ItemDBRow).Value
        Next InvItemRow
NoItems:
        MsgBox "La facture '" & Format(.Range("N6").Value, "000000") & "' est enregistrée." & vbNewLine & vbNewLine & "Le total de la facture est " & Trim(Format(.Range("N51").Value, "### ##0.00 $")) & " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    End With
    shInvoice.Range("B27").Value = False
    If shInvoice.Range("B28").Value Then Debug.Print Tab(5); "Total de la facture '" & Format(shInvoice.Range("N6").Value, "000000") & "' (avant taxes) est de " & Format(shInvoice.Range("N51").Value, "### ##0.00 $")
Fast_Exit_Sub:
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_SaveUpdate()" & vbNewLine
End Sub

Sub Invoice_Load()
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_Load() @ " & Time
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
            If shInvoice.Range("B28").Value And InvCol <> 3 Then Debug.Print "InvRow = " & InvRow & "   InvCol = " & InvCol & " - " & .Range(InvList.Cells(1, InvCol).Value) & " <-- " & InvList.Cells(InvRow, InvCol).Value
            If InvCol <> 3 Then .Range(InvList.Cells(1, InvCol).Value).Value = InvList.Cells(InvRow, InvCol).Value 'Load Invoice List Data
        Next InvCol
        'Load Invoice Items
        With InvItems
            LastRow = .Range("A9999").End(xlUp).Row
            If LastRow < 4 Then Exit Sub
            If shInvoice.Range("B28").Value Then Debug.Print "LastRow = " & LastRow & "   Copie de '" & "A3:G" & LastRow & "   Critère: " & .Range("L3").Value
            .Range("A3:G" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("L2:L3"), CopyToRange:=.Range("N2:S2"), Unique:=True
            LastResultRow = .Range("V9999").End(xlUp).Row
            If shInvoice.Range("B28").Value Then Debug.Print "Based on column 'V' (InvItems), LastResultRow = " & LastResultRow
            If LastResultRow < 3 Then GoTo NoItems
            For ResultRow = 3 To LastResultRow
                InvItemRow = .Range("R" & ResultRow).Value 'Set Invoice Row
                If shInvoice.Range("B28").Value Then Debug.Print Tab(20); "Invoice Item Row (InvItemRow) = " & InvItemRow & _
                    "   shInvoice.Range('K'" & InvItemRow & ")=" & shInvoice.Range("K" & InvItemRow).Value & " devient " & "InvItems.Range('N'" & ResultRow & ") = " & .Range("N" & ResultRow).Value & _
                    "   shInvoice.Range('L'" & InvItemRow & ")=" & shInvoice.Range("L" & InvItemRow).Value & " devient " & "InvItems.Range('O'" & ResultRow & ") = " & .Range("O" & ResultRow).Value & _
                    "   shInvoice.Range('M'" & InvItemRow & ")=" & shInvoice.Range("M" & InvItemRow).Value & " devient " & "InvItems.Range('P'" & ResultRow & ") = " & .Range("P" & ResultRow).Value & _
                shInvoice.Range("K" & InvItemRow & ":M" & InvItemRow).Value = .Range("N" & ResultRow & ":P" & ResultRow).Value 'Item details
                If shInvoice.Range("B28").Value Then Debug.Print Tab(30); "shInvoice.Range('O'" & InvItemRow & ")=" & shInvoice.Range("O" & InvItemRow).Value & " devient " & "InvItems.Range('S'" & ResultRow & ") = " & .Range("S" & ResultRow).Value
                shInvoice.Range("O" & InvItemRow).Value = .Range("S" & ResultRow).Value  'Set Item DB Row
            Next ResultRow
NoItems:
        End With
        .Range("B24").Value = False 'Set Invoice Load To false
    End With
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_Load()" & vbNewLine
End Sub

Sub Invoice_Delete()
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Invoice_Delete() @ " & Time
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
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Invoice_Delete()" & vbNewLine
End Sub

Sub Previsualisation_PDF() 'RMV - 2023-10-04 @ 15:45
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Previsualisation_PDF() @ " & Time
    shFactureFinale.PrintOut , , , True, True, , , , False
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Previsualisation_PDF()" & vbNewLine
End Sub

Sub Creation_PDF_Email() 'RMV - 2023-10-04 @ 15:46
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Creation_PDF_Email() @ " & Time
    Create_PDF_Email_Sub shInvoice.Range("N6").Value
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Creation_PDF_Email()" & vbNewLine
End Sub

Sub Create_PDF_Email_Sub(NoFacture As Long)
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Create_PDF_Email_Sub(NoFacture As Long) @ " & Time
    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF et préparation du courriel pour envoyer la facture
    Dim Result As Boolean
    Result = Create_PDF_Email_Function(NoFacture, "CreateEmail")
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Create_PDF_Email_Sub(NoFacture As Long)" & vbNewLine
End Sub

Function Create_PDF_Email_Function(NoFacture As Long, Optional action As String = "SaveOnly") As Boolean
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Function Create_PDF_Email_Function" & _
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
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Create_PDF_Email_Function(NoFacture As Long, Optional action As String = """"SaveOnly"""") As Boolean" & vbNewLine
End Function

Sub Prev_Invoice()
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Prev_Invoice() @ " & Time
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
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Prev_Invoice()" & vbNewLine
End Sub

Sub Next_Invoice()
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Next_Invoice() @ " & Time
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
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Next_Invoice()" & vbNewLine
End Sub

Sub Cacher_Heures()
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - Sub Cacher_Heures() @ " & Time
    shFactureFinale.Range("C64:D65").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Cacher_Heures()" & vbNewLine
End Sub

Sub Montrer_Heures()
    If shInvoice.Range("B28").Value Then Debug.Print "Now entering - [Invoice_Macros] - Sub Montrer_Heures() @ " & Time
    shFactureFinale.Range("C64:D65").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    If shInvoice.Range("B28").Value Then Debug.Print "Now exiting  - [Invoice_Macros] - Sub Montrer_Heures()" & vbNewLine
End Sub

Sub Goto_Onglet_Preparation_Facture()
    shInvoice.Select
    shInvoice.Range("C1").Select
End Sub

Sub Goto_Onglet_Facture_Finale()
    shFactureFinale.Select
End Sub

