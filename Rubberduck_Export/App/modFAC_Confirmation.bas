Attribute VB_Name = "modFAC_Confirmation"
Option Explicit

Public invNo As String

Sub Get_Invoice_Data(noFact As String)

    'Save original worksheet
    Dim oWorkSheet As Worksheet: Set oWorkSheet = ActiveSheet
    
    'Reference to A/R master file
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    Dim result As Variant
    Dim rngToSearch As Range: Set rngToSearch = ws.Range("A1").CurrentRegion.Offset(0, 0).Resize(lastUsedRow, 1)
    result = Application.WorksheetFunction.XLookup(noFact, _
                                                   rngToSearch, _
                                                   rngToSearch, _
                                                   "Not Found", _
                                                   0, _
                                                   1)
    
    If result <> "Not Found" Then
        Dim matchedRow As Long
        matchedRow = Application.Match(noFact, rngToSearch, 0)
        
        Call Display_Invoice_info(ws, matchedRow)
        
        Call Insert_PDF_WIP_Icons
        
        Dim resultArr As Variant
        resultArr = Fn_Get_TEC_Invoiced_By_This_Invoice(noFact)
        
        If Not IsEmpty(resultArr) Then
            Dim TECSummary() As Variant
            ReDim TECSummary(1 To 10, 1 To 3)
            Call Get_TEC_Summary_For_That_Invoice(resultArr, TECSummary)
            
            Dim FeesSummary() As Variant
            ReDim FeesSummary(1 To 5, 1 To 3)
            Call Get_Fees_Summary_For_That_Invoice(resultArr, FeesSummary)
        End If
        oWorkSheet.Activate
    Else
        MsgBox "La facture n'existe pas"
        GoTo Clean_Exit
    End If
    
Clean_Exit:
    Set oWorkSheet = Nothing
    Set rngToSearch = Nothing
    Set ws = Nothing

End Sub

Sub Insert_PDF_WIP_Icons()

    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Dim i As Long
    Dim iconPath As String
    iconPath = wshAdmin.Range("F5").value & Application.PathSeparator & "Resources"
    
    Dim pic As Picture
    Dim cell As Range
    
    '1. Insert the PDF icon
    
    'Set the cell where the icon should be inserted
    Set cell = ws.Cells(7, 12) 'Set the cell where the icon should be inserted
            
    Set pic = ws.Pictures.Insert(iconPath & Application.PathSeparator & "AdobeAcrobatReader.png")
    With pic
        .Name = "PDF"
        .Top = cell.Top + 10
        .Left = cell.Left + 10
        .Height = 50 'cell.Height
        .Width = 50 'cell.width
        .Placement = xlMoveAndSize
        .OnAction = "FAC_Confirmation_Display_PDF_Invoice"
    End With
    
    '2. Insert the WIP icon
    
    'Set the cell where the icon should be inserted
    Set cell = ws.Cells(14, 5) 'Set the cell where the icon should be inserted
    
    Set pic = ws.Pictures.Insert(iconPath & Application.PathSeparator & "WIP.png")
    With pic
        .Name = "WIP"
        .Top = cell.Top + 10
        .Left = cell.Left + 10
        .Height = 50 'cell.Height
        .Width = 50 'cell.width
        .Placement = xlMoveAndSize
        .OnAction = "FAC_Confirmation_Report_Detailed_TEC"
    End With
    
    'Libérer la mémoire
    Set cell = Nothing
    Set pic = Nothing
    Set ws = Nothing
    
End Sub

Sub FAC_Confirmation_Display_PDF_Invoice()

    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    'Assuming the invoice number is at 'F5'
    Dim fullPDFFileName As String
    fullPDFFileName = wshAdmin.Range("F5").value & FACT_PDF_PATH & _
        Application.PathSeparator & ws.Cells(5, 6).value & ".pdf"
    
    'Open the invoice using Adobe Acrobat Reader
    If fullPDFFileName <> "" Then
        Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " & Chr(34) & fullPDFFileName & Chr(34), vbNormalFocus
    Else
        MsgBox "Je ne retrouve pas cette facture", vbExclamation
    End If
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub Display_Invoice_info(wsF As Worksheet, r As Long)

    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    'Display all fields from FAC_Entête
    With ws
        .Range("L5").value = wsF.Cells(r, 2).value
    
        ws.Range("F7").value = wsF.Cells(r, 5).value
        ws.Range("F8").value = wsF.Cells(r, 6).value
        ws.Range("F9").value = wsF.Cells(r, 7).value
        ws.Range("F10").value = wsF.Cells(r, 8).value
        ws.Range("F11").value = wsF.Cells(r, 9).value
        
        ws.Range("L13").value = wsF.Cells(r, 10).value
        ws.Range("L14").value = wsF.Cells(r, 12).value
        ws.Range("L15").value = wsF.Cells(r, 14).value
        ws.Range("L16").value = wsF.Cells(r, 16).value
        ws.Range("L17").formula = "=SUM(L13:L16)"
        
        ws.Range("L18").value = wsF.Cells(r, 18).value
        ws.Range("L19").value = wsF.Cells(r, 20).value
        ws.Range("L21").formula = "=SUM(L17:L19)"
        
        ws.Range("L23").value = wsF.Cells(r, 22).value
        ws.Range("L25").formula = "=L21 - L23"
        
    End With
    
    'Take care of invoice type (to be confirmed OR already confirmed)
    If wsF.Cells(r, 3).value = "AC" Then
        ws.Range("H5").value = "À CONFIRMER"
        ws.Shapes("btnFAC_Confirmation").Visible = True
    Else
        ws.Range("H5").value = ""
        ws.Shapes("btnFAC_Confirmation").Visible = False
    End If
    
    'Make OK button visible
    ws.Shapes("btnFAC_Confirmation_OK").Visible = True
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Application.EnableEvents = True

End Sub

Sub FAC_Confirmation_Report_Detailed_TEC()

    'Utilisation d'un AdvancedFilter directement dans TEC_Local (BI:BX)
    Call Get_Detail_TEC_Invoice_AF(invNo)

    Dim ws As Worksheet: Set ws = wshTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "BI").End(xlUp).row
    
    'Est-ce que nous avons des TEC pour cette facture ?
    If lastUsedRow < 3 Then
        GoTo Nothing_to_Print
    End If
    
    Call FAC_Confirmation_Creer_Rapport_TEC_Factures

    Exit Sub
    
Nothing_to_Print:
    MsgBox "Il n'y a aucun TEC associé à la facture '" & invNo & "'"

    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub FAC_Confirmation_Creer_Rapport_TEC_Factures()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Creer_Rapport_TEC_Factures", 0)
    
    Dim cheminFichier As String
    
    'Tenter d'assigner la feuille qui existe peut-être
    Dim strRapport As String
    strRapport = "Rapport TEC facturés"
    Dim wsRapport As Worksheet
    On Error Resume Next ' Eviter erreur si la feuille existe déjà
    Set wsRapport = ThisWorkbook.Sheets(strRapport)
    On Error GoTo 0
    
    'Si la feuille "Rapport TEC facturés" n'existe pas, la créer
    If wsRapport Is Nothing Then
        Set wsRapport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsRapport.Name = strRapport
    Else
        wsRapport.Cells.Clear 'Vider la feuille si elle existe déjà
    End If
    
    'Mettre en forme la feuille de rapport
    With wsRapport
        ' Titre du rapport
        .Range("A1").value = "TEC facturés pour la facture '" & invNo & "'"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.size = 12
        
        'Ajouter une date de génération du rapport
        .Range("A2").value = "Date de création : " & Format(Now, "dd/mm/yyyy")
        .Range("A2").Font.Italic = True
        .Range("A2").Font.size = 10
        
        'Entête du rapport (A4:D4)
        .Range("A4").value = "Date"
        .Range("B4").value = "Prof."
        .Range("C4").value = "Description"
        .Range("D4").value = "Heures"
        With .Range("A4:D4")
            .Font.Bold = True
            .Font.Italic = True
            .Font.Color = vbWhite
            .HorizontalAlignment = xlCenter
        End With
        
        'Corps du rapport
        .Range("A5:D999").VerticalAlignment = xlTop
        With .Range("A4:D4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        'Supposons que nous résumons des données d'une autre feuille
        Dim wsSource As Worksheet
        Set wsSource = wshTEC_Local 'Utilisation des résultats du AF (BI:BX)
        
        'Copier quelques données de la source
        Dim rngResult As Range
        Set rngResult = wsSource.Range("BI1").CurrentRegion.Offset(2, 0)
        'Redimensionner la plage après l'offset pour ajuster la taille (réduire le nombre de lignes)
        Set rngResult = rngResult.Resize(rngResult.rows.count - 2)
        'Transfert des données vers un tableau
        Dim tableau As Variant
        tableau = rngResult.value
        
        Dim r As Long
        r = 4 'Nombre de lignes d'entête
        
        Dim i As Long
        For i = LBound(tableau, 1) To UBound(tableau, 1)
            r = r + 1
            wsRapport.Cells(r, 1) = tableau(i, 4)
            wsRapport.Cells(r, 2) = tableau(i, 3)
            wsRapport.Cells(r, 3) = tableau(i, 7)
            wsRapport.Cells(r, 4) = tableau(i, 8)
        Next i

        'Ajouter une bordure aux données
        .Range("A4:D" & r).Borders.LineStyle = xlContinuous
        With .Range("A5:D" & r).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Range("A5:D" & r).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        
        .Range("A4:D" & r).Font.Name = "Aptos Narrow"
        .Range("A4:D" & r).Font.size = 10
        
        .columns("A").ColumnWidth = 10
        .Range("A4:A" & r).HorizontalAlignment = xlCenter
        
        .columns("B").ColumnWidth = 6
        .Range("B4:B" & r).HorizontalAlignment = xlCenter
        
        .columns("C").ColumnWidth = 72
        .columns("C").WrapText = True
        
        .columns("D").ColumnWidth = 7
        .columns("D").NumberFormat = "##0.00"
        
    End With

    'Configurer la mise en page pour l'impression ou l'export en PDF
    With wsRapport.PageSetup
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        
        'Ajuster la marge des en-têtes et pieds de page (1 cm)
        .HeaderMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(1)
        
        .Orientation = xlPortrait 'Portrait
        .FitToPagesWide = 1 'Ajuster sur une page en largeur
        .FitToPagesTall = False ' Ne pas ajuster en hauteur
        .PrintArea = "A1:D" & r ' Définir la zone d'impression
        .CenterHorizontally = True ' Centrer horizontalement
        .CenterVertically = False ' Centrer verticalement
    End With
    
    MsgBox "Le rapport a été généré sur la feuille " & strRapport
    
    'On se déplace à la feuille contenant le rapport
    wsRapport.Activate
    
    'Libérer la mémoire
    Set rngResult = Nothing
    Set wsRapport = Nothing
    Set wsSource = Nothing
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Creer_Rapport_TEC_Factures", startTime)
    
End Sub

Sub Get_Detail_TEC_Invoice_AF(noFact As String) '2024-10-20 @ 11:11

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:Get_Detail_TEC_Invoice_AF", 0)

    'Voir la feuille TEC_Local
    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'AdvancedFilter par Numéro de Facture
    
    'Définir le range des critères (Numéro de Facture)
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("BG2:BG3")
    ws.Range("BG3").value = CStr(noFact)
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("BI1").CurrentRegion
    rngResult.Offset(2, 0).Clear
    Set rngResult = ws.Range("BI2:BX2")
    
    'AdvanceFilter
    ws.Range("tblTEC_Local[#All]").AdvancedFilter _
                                            action:=xlFilterCopy, _
                                            criteriaRange:=rngCriteria, _
                                            CopyToRange:=rngResult, _
                                            Unique:=False
        
    'Tri des informations
    Dim lastResultRow As Long
    lastResultRow = ws.Cells(ws.rows.count, "BI").End(xlUp).row
    
    'Est-il nécessaire de trier les résultats ?
    If lastResultRow > 3 Then
        With ws.Sort 'Sort - Date, ProfID, TEC_ID
            .SortFields.Clear
            'First sort On Date
            .SortFields.Add key:=ws.Range("BL2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Prof_ID
            .SortFields.Add key:=ws.Range("BJ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Third, sort On TecID
            .SortFields.Add key:=ws.Range("BI2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange ws.Range("BI3:BW" & lastResultRow)
            .Apply 'Apply Sort
         End With
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Free memory
    Set rngCriteria = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:Get_Detail_TEC_Invoice_AF", startTime)
    
End Sub

Sub FAC_Entête_AdvancedFilter_AC_C() '2024-07-19 @ 13:58

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    With ws
        'Setup the destination Range and clear it before applying AdvancedFilter
        Dim lastUsedRow As Long
        Dim destinationRng As Range: Set destinationRng = .Range("AY2:BP2")
        lastUsedRow = ws.Cells(ws.rows.count, "AY").End(xlUp).row
        If lastUsedRow > 2 Then
            ws.Range("AY3:BP" & lastUsedRow).ClearContents
        End If
        
        'Setup source data including headers
        lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
        If lastUsedRow < 3 Then Exit Sub 'No data to filter
        Dim sourceRng As Range: Set sourceRng = .Range("A2:V" & lastUsedRow)
        
        'Define the criteria range including headers
        Dim criteriaRng As Range: Set criteriaRng = ws.Range("AW2:AW3")
    
        ' Apply the advanced filter
        sourceRng.AdvancedFilter xlFilterCopy, criteriaRng, destinationRng, False
        
        lastUsedRow = ws.Cells(ws.rows.count, "AY").End(xlUp).row
        If lastUsedRow < 4 Then Exit Sub
        With ws.Sort 'Sort - Inv_No
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("AY3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Invoice Number
            .SetRange ws.Range("AY3:BP" & lastUsedRow) 'Set Range
            .Apply 'Apply Sort
         End With
     End With

    'Libérer la mémoire
    Set criteriaRng = Nothing
    Set destinationRng = Nothing
    Set sourceRng = Nothing
    Set ws = Nothing

End Sub

Sub Show_Unconfirmed_Invoice()

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    Application.ScreenUpdating = False
    
    'Clear contents or the area
    Dim lastUsedRow As Long
    lastUsedRow = wshFAC_Confirmation.Cells(wshFAC_Confirmation.rows.count, "P").End(xlUp).row
    If lastUsedRow > 3 Then
        wshFAC_Confirmation.Range("P4:AA" & lastUsedRow).ClearContents
    End If

    'Set criteria for AvancedFilter
    ws.Range("AW3").value = "AC"
    
    Call FAC_Entête_AdvancedFilter_AC_C
    
    Dim lastUsedRowAF As Long
    lastUsedRowAF = ws.Cells(ws.rows.count, "AY").End(xlUp).row
    If lastUsedRowAF < 3 Then
        GoTo Clean_Exit
    End If
    
'    wshFAC_Confirmation.Unprotect
    
    Application.EnableEvents = False
    
    Dim i As Integer
    For i = 3 To lastUsedRowAF
        With wshFAC_Confirmation
            wshFAC_Confirmation.Cells(i + 1, 16).Locked = False
            .Cells(i + 1, 16).value = ws.Cells(i, 51)
            .Cells(i + 1, 17).value = ws.Cells(i, 52)
            .Cells(i + 1, 18).value = ws.Cells(i, 55)
            .Cells(i + 1, 19).value = ws.Cells(i, 67)
            .Cells(i + 1, 20).value = ws.Cells(i, 56)
            .Cells(i + 1, 21).value = ws.Cells(i, 58)
            .Cells(i + 1, 22).value = ws.Cells(i, 60)
            .Cells(i + 1, 23).value = ws.Cells(i, 62)
            .Cells(i + 1, 24).value = ws.Cells(i, 64)
            .Cells(i + 1, 25).value = ws.Cells(i, 66)
            .Cells(i + 1, 26).value = ws.Cells(i, 68)
        End With
    Next i
    
    Application.EnableEvents = True
    
    With wshFAC_Confirmation
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    wshFAC_Confirmation.Range("F5").value = ""
    
    Application.ScreenUpdating = True
    
Clean_Exit:
    Set ws = Nothing

End Sub

Sub Get_TEC_Summary_For_That_Invoice(arr As Variant, ByRef TECSummary As Variant)

    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    'Setup a Dictionary to summarize the hours by Professionnal
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")

    Dim pro As String
    Dim hres As Double
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        pro = wsTEC.Cells(arr(i), 3).value
        hres = wsTEC.Cells(arr(i), 8).value
        If hres <> 0 Then
            If dictHours.Exists(pro) Then
                dictHours(pro) = dictHours(pro) + hres
            Else
                dictHours.Add pro, hres
            End If
        End If
    Next i
    
    Dim profID As Long
    Dim rowInWorksheet As Long: rowInWorksheet = 13
    Dim prof As Variant
    Application.EnableEvents = False
    If dictHours.count <> 0 Then
        For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) 'Sort dictionary by hours in descending order
            Dim strProf As String
            strProf = prof
            profID = Fn_GetID_From_Initials(strProf)
            hres = dictHours(prof)
            Dim tauxHoraire As Currency
            tauxHoraire = Fn_Get_Hourly_Rate(profID, wshFAC_Confirmation.Range("L5").value)
            wshFAC_Confirmation.Cells(rowInWorksheet, 6) = strProf
            wshFAC_Confirmation.Cells(rowInWorksheet, 7) = _
                    CDbl(Format$(hres, "0.00"))
            wshFAC_Confirmation.Cells(rowInWorksheet, 8) = _
                    CDbl(Format$(tauxHoraire, "# ##0.00 $"))
            rowInWorksheet = rowInWorksheet + 1
    '        Debug.Print "Summary : " & strProf & " = " & hres & " @ " & tauxHoraire
    '        Cells(rowSelected, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
    '        rowSelected = rowSelected + 1
        Next prof
    End If
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set dictHours = Nothing
    Set prof = Nothing
    Set wsTEC = Nothing
    
End Sub

Sub Get_TEC_Total_For_That_Invoice(arr As Variant, ByRef TECTotal As Double)

    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    'Setup a Dictionary to summarize the hours by Professionnal
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")

    Dim pro As String
    Dim hres As Double
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        pro = wsTEC.Cells(arr(i), 3).value
        hres = wsTEC.Cells(arr(i), 8).value
        If hres <> 0 Then
            If dictHours.Exists(pro) Then
                dictHours(pro) = dictHours(pro) + hres
            Else
                dictHours.Add pro, hres
            End If
        End If
    Next i
    
    Dim profID As Long
    Dim rowInWorksheet As Long: rowInWorksheet = 13
    Dim prof As Variant
    Application.EnableEvents = False
    If dictHours.count <> 0 Then
        For Each prof In dictHours
            Dim strProf As String
            strProf = prof
            profID = Fn_GetID_From_Initials(strProf)
            hres = dictHours(prof)
            Dim tauxHoraire As Currency
            tauxHoraire = Fn_Get_Hourly_Rate(profID, wshFAC_Confirmation.Range("L5").value)
            wshFAC_Confirmation.Cells(rowInWorksheet, 6) = strProf
            wshFAC_Confirmation.Cells(rowInWorksheet, 7) = _
                    CDbl(Format$(hres, "0.00"))
            wshFAC_Confirmation.Cells(rowInWorksheet, 8) = _
                    CDbl(Format$(tauxHoraire, "# ##0.00 $"))
            rowInWorksheet = rowInWorksheet + 1
    '        Debug.Print "Summary : " & strProf & " = " & hres & " @ " & tauxHoraire
    '        Cells(rowSelected, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
    '        rowSelected = rowSelected + 1
        Next prof
    End If
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set dictHours = Nothing
    Set prof = Nothing
    Set wsTEC = Nothing
    
End Sub

Sub Get_Fees_Summary_For_That_Invoice(arr As Variant, ByRef FeesSummary As Variant)

    Dim wsFees As Worksheet: Set wsFees = wshFAC_Sommaire_Taux
    
    'Determine the last used row
    Dim lastUsedRow As Long
    lastUsedRow = wsFees.Cells(wsFees.rows.count, "A").End(xlUp).row
    
    'Get Invoice number
    Dim invNo As String
    invNo = Trim(wshFAC_Confirmation.Range("F5").value)
    
    'Use Range.Find to locate the first cell with the InvoiceNo
    Dim cell As Range
    Set cell = wsFees.Range("A2:A" & lastUsedRow).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the invNo was found at all
    Dim firstAddress As String
    Dim rowFeesSummary As Long: rowFeesSummary = 20
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Application.EnableEvents = False
        Do
            'Display values in the worksheet
            wshFAC_Confirmation.Range("F" & rowFeesSummary).value = wsFees.Cells(cell.row, 3).value
            wshFAC_Confirmation.Range("G" & rowFeesSummary).value = _
                        CDbl(Format$(wsFees.Cells(cell.row, 4).value, "##0.00"))
            wshFAC_Confirmation.Range("H" & rowFeesSummary).value = _
                        CDbl(Format$(wsFees.Cells(cell.row, 5).value, "##,##0.00 $"))
            rowFeesSummary = rowFeesSummary + 1
            'Find the next cell with the invNo
            Set cell = wsFees.Range("A2:A" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set wsFees = Nothing
    
End Sub

Sub FAC_Confirmation_Clear_Cells_And_PDF_Icon()

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshFAC_Confirmation:FAC_Confirmation_Clear_Cells_And_PDF_Icon", 0)
    
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Application.ScreenUpdating = False
    
    ws.Range("F5,H5,L5,F7:I11,L13:L19,L21,L23,L25,F13:H17,F20:H24").ClearContents
    
    Dim pic As Picture
    For Each pic In ws.Pictures
        On Error Resume Next
        pic.Delete
        On Error GoTo 0
    Next pic
    
    Application.ScreenUpdating = True
    
    'Hide both buttons
    ws.Shapes("btnFAC_Confirmation").Visible = False
    ws.Shapes("btnFAC_Confirmation_OK").Visible = False
    
    Call Show_Unconfirmed_Invoice
    
    'Libérer la mémoire
    Set pic = Nothing
    Set ws = Nothing

    Application.EnableEvents = True
    
    wshFAC_Confirmation.Range("F5").Select
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Clear_Cells_And_PDF_Icon", startTime)

End Sub

Sub FAC_Confirmation_OK_Button_Click()

    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Call FAC_Confirmation_Clear_Cells_And_PDF_Icon
    
    ws.Range("F5").Select
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub FAC_Confirmation_Confirm_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshFAC_Confirmation:FAC_Confirmation_Confirm_Click", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Dim invNo As String
    invNo = ws.Range("F5").value
    
    ws.Shapes("btnFAC_Confirmation").Visible = False
    
    Dim answerYesNo As Long
    answerYesNo = MsgBox("Êtes-vous certain de vouloir CONFIRMER cette facture ? ", _
                         vbYesNo + vbQuestion, "Confirmation de facture")
    If answerYesNo = vbNo Then
        MsgBox _
            Prompt:="Cette facture ne sera PAS CONFIRMÉE ! ", _
            Title:="Confirmation", _
            Buttons:=vbCritical
            GoTo Clean_Exit
    End If
    
    If answerYesNo = vbYes Then
    
        Call FAC_Confirmation_Facture(invNo)
        
    End If
    
Clean_Exit:

    Call FAC_Confirmation_Clear_Cells_And_PDF_Icon

    wshFAC_Confirmation.Range("F5").Select
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Confirm_Click", startTime)

End Sub

Sub FAC_Confirmation_Get_GL_Posting(invNo)

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshFAC_Confirmation:FAC_Confirmation_Get_GL_Posting", 0)
    
    Dim wsGL As Worksheet: Set wsGL = wshGL_Trans
    
    Dim lastUsedRow
    lastUsedRow = wsGL.Range("A99999").End(xlUp).row
    Dim rngToSearch As Range: Set rngToSearch = wsGL.Range("D1:D" & lastUsedRow)
    
    'Use Range.Find to locate the first cell with the invNo
    Dim cell As Range
    Set cell = wsGL.Range("D2:D" & lastUsedRow).Find(What:="FACTURE:" & invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the invNo was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Dim r As Long
        r = 38
        Application.EnableEvents = False
        Do
            'Save the information for invoice deletion
            r = r + 1
            'Find the next cell with the invNo
            Set cell = wsGL.Range("D2:D" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If

    'Libérer la mémoire
    Set cell = Nothing
    Set rngToSearch = Nothing
    Set wsGL = Nothing
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Get_GL_Posting", startTime)

End Sub

Sub FAC_Confirmation_Facture(invNo As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("wshFAC_Confirmation:FAC_Confirmation_Facture(" & invNo & ")", 0)
    
    'Update the type of invoice (Master)
    Call FAC_Confirmation_Update_BD_MASTER(invNo)
    
    'Update the type of invoice (Locally)
    Call FAC_Confirmation_Update_Locally(invNo)
    
    'Do the G/L posting
    Call FAC_Confirmation_GL_Posting(invNo)
    
'    MsgBox "Cette facture a été confirmée avec succès", vbInformation

    'Clear the cells on the current Worksheet
    Call FAC_Confirmation_Clear_Cells_And_PDF_Icon
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Facture(" & invNo & ")", startTime)
    
End Sub

Sub FAC_Confirmation_Update_BD_MASTER(invoice As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Update_BD_MASTER(" & invoice & ")", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Entête"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
              ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim SQL As String
    'Open the recordset for the specified invoice
    SQL = "SELECT * FROM [" & destinationTab & "$] WHERE Inv_No = '" & invoice & "'"
    rs.Open SQL, conn, 2, 3
    If Not rs.EOF Then
        'Update AC_ouC with 'C'
        rs.Fields("AC_C").value = "C"
        rs.update
    Else
        'Handle the case where the specified invoice is not found
        MsgBox "La facture '" & invoice & "' n'existe pas!", vbCritical
    End If
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Update_BD_MASTER", startTime)

End Sub

Sub FAC_Confirmation_Update_Locally(invoice As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Update_Locally(" & invoice & ")", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=invoice, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim r As Long, rowToBeUpdated As Long, TECID As Long
    If Not foundRange Is Nothing Then
        r = foundRange.row
        ws.Cells(r, 3).value = "C"
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entête."
    End If
    
    'Libérer la mémoire
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Update_Locally", startTime)

End Sub

Sub FAC_Confirmation_GL_Posting(invoice As String) '2024-08-18 @17:15

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:FAC_Confirmation_GL_Posting(" & invoice & ")", 0)

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=invoice, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim r As Long
    If Not foundRange Is Nothing Then
        r = foundRange.row
        Dim dateFact As Date
        dateFact = Left(ws.Cells(r, 2).value, 10)
        Dim hono As Currency
        hono = ws.Cells(r, 10).value
        Dim misc1 As Currency, misc2 As Currency, misc3 As Currency
        misc1 = ws.Cells(r, 12).value
        misc2 = ws.Cells(r, 14).value
        misc3 = ws.Cells(r, 16).value
        Dim tps As Currency, tvq As Currency
        tps = ws.Cells(r, 18).value
        tvq = ws.Cells(r, 20).value
        
        Dim descGL_Trans As String, source As String
        descGL_Trans = ws.Cells(r, 6).value
        source = "FACTURE:" & invoice
        
        Dim MyArray(1 To 7, 1 To 4) As String
        
        'AR amount
        If hono + misc1 + misc2 + misc3 + tps + tvq Then
            MyArray(1, 1) = "1100"
            MyArray(1, 2) = "Comptes clients"
            MyArray(1, 3) = hono + misc1 + misc2 + misc3 + tps + tvq
            MyArray(1, 4) = ""
        End If
        
        'Professional Fees (hono)
        If hono Then
            MyArray(2, 1) = "4000"
            MyArray(2, 2) = "Revenus de consultation"
            MyArray(2, 3) = -hono
            MyArray(2, 4) = ""
        End If
        
        'Miscellaneous Amount # 1 (misc1)
        If misc1 Then
            MyArray(3, 1) = "4010"
            MyArray(3, 2) = "Revenus - Frais de poste"
            MyArray(3, 3) = -misc1
            MyArray(3, 4) = ""
        End If
        
        'Miscellaneous Amount # 2 (misc2)
        If misc2 Then
            MyArray(4, 1) = "4015"
            MyArray(4, 2) = "Revenus - Sous-traitants"
            MyArray(4, 3) = -misc2
            MyArray(4, 4) = ""
        End If
        
        'Miscellaneous Amount # 3 (misc3)
        If misc3 Then
            MyArray(5, 1) = "4020"
            MyArray(5, 2) = "Revenus - Autres Frais"
            MyArray(5, 3) = -misc3
            MyArray(5, 4) = ""
        End If
        
        'GST to pay (tps)
        If tps Then
            MyArray(6, 1) = "1202"
            MyArray(6, 2) = "TPS percues"
            MyArray(6, 3) = -tps
            MyArray(6, 4) = ""
        End If
        
        'PST to pay (tvq)
        If tvq Then
            MyArray(7, 1) = "1203"
            MyArray(7, 2) = "TVQ percues"
            MyArray(7, 3) = -tvq
            MyArray(7, 4) = ""
        End If
        
        Dim glEntryNo As Long
        Call GL_Posting_To_DB(dateFact, descGL_Trans, source, MyArray, glEntryNo)
        
        Call GL_Posting_Locally(dateFact, descGL_Trans, source, MyArray, glEntryNo)
        
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entête.", vbCritical
    End If
    
    'Libérer la mémoire
    On Error Resume Next
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_GL_Posting", startTime)

End Sub

Sub FAC_Confirmation_Back_To_FAC_Menu()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:Back_To_FAC_Menu", 0)
   
    wshFAC_Confirmation.Unprotect '2024-08-21 @ 05:06
    
    Application.EnableEvents = False
    wshFAC_Confirmation.Range("F5").ClearContents
    Application.EnableEvents = True
    
    wshFAC_Confirmation.Visible = xlSheetHidden

    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select
    
    Call Log_Record("modFAC_Confirmation:Back_To_FAC_Menu", startTime)
    
End Sub


