Attribute VB_Name = "modGL_TB"
Option Explicit

Sub GL_TB_Build_Trial_Balance() '2024-03-05 @ 13:34
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_Build_Trial_Balance()")
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Clear TB cells - Contents & formats
    Dim lastUsedRow As Long
    lastUsedRow = wshGL_BV.Range("D99999").End(xlUp).row
    wshGL_BV.Range("D4" & ":G" & lastUsedRow + 2).clear

    'Clear Detail transaction section
    wshGL_BV.Range("L4:T99999").ClearContents
    With wshGL_BV.Range("S4:S9999").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Add the cut-off date in the header (printing purposes)
    Dim minDate As Date, dateCutOff As Date
    wshGL_BV.Range("C2").value = "Au " & CDate(Format(wshGL_BV.Range("J1").value, "dd-mm-yyyy"))

    minDate = CDate("01/01/2023")
    dateCutOff = CDate(wshGL_BV.Range("J1").value)
    wshGL_BV.Range("B2").value = 3
    wshGL_BV.Range("B10").value = 0
    
    'Step # 1 - Use AdvancedFilter on GL_Trans for ALL accounts and transactions
    '           between the 2 dates
    Call GL_TB_AdvancedFilter_By_GL("", minDate, dateCutOff)
    'The SORT method does not sort correctly the GLNo, since there is NUMBER and NUMBER+LETTER !!!
    
    lastUsedRow = wshGL_Trans.Range("T999999").End(xlUp).row
    If lastUsedRow < 2 Then Exit Sub
    
    'The Chart of Account will drive the results, so the sort order is determined by COA
    Dim arr As Variant
    arr = Fn_Get_Chart_Of_Accounts(2) 'Returns array with 2 columns (Code, Description)
    
    Dim dictSolde As Dictionary 'GLNo dictionary
    Set dictSolde = New Dictionary
    Dim arrSolde() As Variant 'GLbalance
    ReDim arrSolde(1 To UBound(arr, 1), 1 To 2)
    Dim newRowID As Long: newRowID = 1
    Dim currRowID As Long
    
    'Parse every line of the result (AdvancedFilter in GL_Trans)
    Dim i As Long, glNo As String, dtct As Currency, MyValue As String, t1 As Currency, t2 As Currency
    For i = 2 To lastUsedRow
        With wshGL_Trans
            glNo = .Range("T" & i).value
            dtct = .Range("V" & i).value - .Range("W" & i).value
            t1 = t1 + .Range("V" & i).value
            t2 = t2 + .Range("W" & i).value
        End With
        If Not dictSolde.Exists(glNo) Then
            dictSolde.add glNo, newRowID
            arrSolde(newRowID, 1) = glNo
'            Debug.Print glNo & "   " & newRowID
            newRowID = newRowID + 1
        End If
        currRowID = dictSolde(glNo)
        'Update the summary array
        arrSolde(currRowID, 2) = arrSolde(currRowID, 2) + dtct
    Next i
    
    Dim sumDT As Currency, sumCT As Currency, GLNoPlusDesc As String
    Dim currRow As Long: currRow = 4
    wshGL_BV.Range("D4:D" & UBound(arrSolde, 1)).HorizontalAlignment = xlCenter
    wshGL_BV.Range("F4:G" & UBound(arrSolde, 1) + 3).HorizontalAlignment = xlRight
    
    Dim r As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        glNo = arr(i, 1)
        If glNo <> "" Then
            r = dictSolde.item(glNo) 'Get the value of the item associated with GLNo
            If r <> 0 Then
                wshGL_BV.Range("D" & currRow).value = glNo
                wshGL_BV.Range("E" & currRow).value = arr(i, 2)
                If arrSolde(r, 2) >= 0 Then
                    wshGL_BV.Range("F" & currRow).value = Format(arrSolde(r, 2), "###,###,##0.00")
                    sumDT = sumDT + arrSolde(r, 2)
                Else
                    wshGL_BV.Range("G" & currRow).value = Format(-arrSolde(r, 2), "###,###,##0.00")
                    sumCT = sumCT - arrSolde(r, 2)
                End If
                currRow = currRow + 1
            End If
        End If
    Next i

    currRow = currRow + 1
    wshGL_BV.Range("B2").value = currRow
    
    'Output Debit total
    Dim rng As Range
    Set rng = wshGL_BV.Range("F" & currRow)
    Call GL_TB_Display_TB_Totals(rng, sumDT) 'Débit total - 2024-06-09 @ 07:51
    
    'Output Credit total
    Set rng = wshGL_BV.Range("G" & currRow)
    Call GL_TB_Display_TB_Totals(rng, sumCT) 'Débit total - 2024-06-09 @ 07:51
    
    wshGL_BV.Range("B2").value = currRow

    'Setup page for printing purposes
    Dim CenterHeaderTxt As String
    CenterHeaderTxt = wshAdmin.Range("NomEntreprise")
    With ActiveSheet.PageSetup
        .CenterHeader = "&""Calibri,Bold""&20 " & CenterHeaderTxt
        .PrintArea = "$D$1:$G$" & currRow
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    Application.EnableEvents = True
    
    ActiveWindow.ScrollRow = 1
  
    Call Output_Timer_Results("modGL_TB:GL_TB_Build_Trial_Balance()", timerStart)

End Sub

Sub GL_TB_Display_TB_Totals(rng As Range, t As Currency) '2024-06-09 @ 07:45

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_Display_TB_Totals()")
    
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets("GL_BV")
    
    With rng
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        .value = t
        .Font.Bold = True
        .NumberFormat = "#,##0.00 $"
    End With
    
'        If .Cells(r, c).value <> .Cells(r, c + 1).value Then
'            Call Erreur_Totaux_DT_CT
'        End If
    
    Call Output_Timer_Results("modGL_TB:GL_TB_Display_TB_Totals()", timerStart)

End Sub

Sub GL_TB_Display_Trans_For_Selected_Account(GLAcct As String, GLDesc As String, minDate As Date, maxDate As Date) 'Display GL Trans for a specific account

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_Display_Trans_For_Selected_Account()")
    
    Dim ws As Worksheet: Set ws = wshGL_BV
    
    'Clear the display area & display the account number & description
    With ws
        .Range("L4:T99999").clear '2024-06-08 @ 15:28
        .Range("L2").value = "Du " & minDate & " au " & maxDate
    
        .Range("L4").Font.Bold = True
        .Range("L4").value = GLAcct & " - " & GLDesc
        .Range("B6").value = GLAcct
        .Range("B7").value = GLDesc
    End With
    
    'Use the Advanced Filter Result already prepared for TB
    Dim row As Range, foundRow As Long, lastResultUsedRow As Long
    lastResultUsedRow = wshGL_Trans.Range("T99999").End(xlUp).row
    If lastResultUsedRow <= 2 Then
        GoTo Exit_Sub
    End If
    foundRow = 0
    
    'Find the first occurence of GlACct in AdvancedFilter Results on GL_Trans
    Dim foundCell As Range, searchRange As Range
    Set searchRange = wshGL_Trans.Range("T1:T" & lastResultUsedRow)
    Set foundCell = searchRange.Find(What:=GLAcct, LookIn:=xlValues, LookAt:=xlWhole)
    foundRow = foundCell.row
    
    'Check if the target value was found
    If foundRow = 0 Then
        MsgBox "Il n'existe aucune transaction pour ce compte (période choisie)."
        Exit Sub
    End If
    
    Dim rowGLDetail As Long
    rowGLDetail = 5
    With ws.Range("S4")
        .value = 0
        .Font.Bold = True
        .NumberFormat = "#,##0.00 $"
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
    End With
    
    Dim d As Date, OK As Integer
    
    With ws
        Do Until wshGL_Trans.Range("T" & foundRow).value <> GLAcct
            'Traitement des transactions détaillées
            d = Format(wshGL_Trans.Range("Q" & foundRow).Value2, "dd-mm-yyyy")
            If d >= minDate And d <= maxDate Then
                .Range("M" & rowGLDetail).value = wshGL_Trans.Range("Q" & foundRow).value
                .Range("N" & rowGLDetail).value = wshGL_Trans.Range("P" & foundRow).value
                .Range("N" & rowGLDetail).HorizontalAlignment = xlCenter
                .Range("O" & rowGLDetail).value = wshGL_Trans.Range("R" & foundRow).value
                .Range("P" & rowGLDetail).value = wshGL_Trans.Range("S" & foundRow).value
                .Range("Q" & rowGLDetail).value = wshGL_Trans.Range("V" & foundRow).value
                .Range("R" & rowGLDetail).value = wshGL_Trans.Range("W" & foundRow).value
                .Range("S" & rowGLDetail).value = ws.Range("S" & rowGLDetail - 1).value + _
                    wshGL_Trans.Range("V" & foundRow).value - wshGL_Trans.Range("W" & foundRow).value
                .Range("T" & rowGLDetail).Value2 = wshGL_Trans.Range("X" & foundRow).value
                foundRow = foundRow + 1
                rowGLDetail = rowGLDetail + 1
                OK = OK + 1
            Else
                foundRow = foundRow + 1
            End If
        Loop
    End With

    With ws.Range("S" & rowGLDetail - 1)
        .Font.Bold = True
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
    End With
        
    'Set columns width for the detailled transactions list
    Dim rng As Range
    lastResultUsedRow = ws.Range("M9999").End(xlUp).row
    Set rng = ws.Range("M5:M" & lastResultUsedRow)
    rng.ColumnWidth = 11
    Set rng = ws.Range("N5:N" & lastResultUsedRow)
    rng.ColumnWidth = 8
    Set rng = ws.Range("O5:O" & lastResultUsedRow)
    rng.ColumnWidth = 40
    Set rng = ws.Range("P5:P" & lastResultUsedRow)
    rng.ColumnWidth = 16
    Set rng = ws.Range("Q5:S" & lastResultUsedRow)
    rng.ColumnWidth = 16
    Set rng = ws.Range("T5:T" & lastResultUsedRow)
    rng.ColumnWidth = 30

    Dim visibleRows As Long
    visibleRows = ActiveWindow.VisibleRange.rows.count
    If lastResultUsedRow > visibleRows Then
        ActiveWindow.ScrollRow = lastResultUsedRow - visibleRows + 3 'Move to the bottom of the worksheet
    Else
        ActiveWindow.ScrollRow = 1
    End If
    
    'Create a Conditional Formating for the displayed transactions
    ws.Unprotect
    With ws.Range("M5:T" & lastResultUsedRow)
        .FormatConditions.add _
            Type:=xlExpression, _
            Formula1:="=ET($M5<>"""";MOD(LIGNE();2)=1)"
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
    
    ws.Protect UserInterfaceOnly:=True
    
Exit_Sub:

    Call Output_Timer_Results("modGL_TB:GL_TB_Display_Trans_For_Selected_Account()", timerStart)

End Sub

Sub GL_TB_AdvancedFilter_By_GL(glNo As String, minDate As Date, maxDate As Date)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_AdvancedFilter_By_GL()")

    With wshGL_Trans
        Dim rgResult As Range, rgData As Range, rgCriteria As Range, rgCopyToRange As Range
        Set rgResult = .Range("P2").CurrentRegion
        rgResult.Offset(1).ClearContents
        
        Set rgData = .Range("A1").CurrentRegion
        .Range("L3").value = ""
        .Range("M3").value = ">=" & Format(minDate, "mm-dd-yyyy")
        .Range("N3").value = "<=" & Format(maxDate, "mm-dd-yyyy")
        
        Set rgCriteria = .Range("L2:N3")
        Set rgCopyToRange = .Range("P1:Y1")
        
        rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgCopyToRange
        
        Dim lastResultUsedRow
        lastResultUsedRow = .Range("P99999").End(xlUp).row
        If lastResultUsedRow < 3 Then GoTo NoSort
        
        'Sort GL_Trans AdvancedFilter results (Range("P2:Y??"))
        With .Sort
                .SortFields.clear
                .SortFields.add key:=wshGL_Trans.Range("T2:T" & lastResultUsedRow), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortTextAsNumbers 'Returns Numeric only (first), then numeric and letters
                .SortFields.add key:=wshGL_Trans.Range("Q2:Q" & lastResultUsedRow), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Sort Based On Date
                .SortFields.add key:=wshGL_Trans.Range("P2:P" & lastResultUsedRow), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Sort Based On JE Number
                .SetRange wshGL_Trans.Range("P2:Y" & lastResultUsedRow) 'Set Range
                .Apply 'Apply Sort
        End With
    End With

NoSort:

    Call Output_Timer_Results("modGL_TB:GL_TB_AdvancedFilter_By_GL()", timerStart)

End Sub

Sub GL_TB_Sub_Totals(glNo As String, GLDesc As String, s As Currency)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_Sub_Totals()")

    Dim r As Long
    With wshGL_BV
        r = .Range("B2").value + 1
        .Range("D" & r).HorizontalAlignment = xlCenter
        .Range("D" & r).value = glNo
        .Range("E" & r).value = GLDesc
        If s > 0 Then
            .Range("F" & r).value = s
        ElseIf s < 0 Then
            .Range("G" & r).value = -s
        End If
        .Range("B2").value = wshGL_BV.Range("B2").value + 1
    End With
    
    Call Output_Timer_Results("modGL_TB:GL_TB_Sub_Totals()", timerStart)

End Sub

Sub GL_TB_Determine_From_And_To_Date(period As String)

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_Determine_From_And_To_Date()")

    Select Case period
        Case "Mois"
            wshGL_BV.Range("B8").value = wshAdmin.Range("MoisDe").value
            wshGL_BV.Range("B9").value = wshAdmin.Range("MoisA").value
        Case "Mois dernier"
            wshGL_BV.Range("B8").value = wshAdmin.Range("MoisPrecDe").value
            wshGL_BV.Range("B9").value = wshAdmin.Range("MoisPrecA").value
        Case "Trimestre"
            wshGL_BV.Range("B8").value = wshAdmin.Range("TrimDe").value
            wshGL_BV.Range("B9").value = wshAdmin.Range("TrimA").value
        Case "Trimestre dernier"
            wshGL_BV.Range("B8").value = wshAdmin.Range("TrimPrecDe").value
            wshGL_BV.Range("B9").value = wshAdmin.Range("TrimPrecA").value
        Case "Année"
            wshGL_BV.Range("B8").value = wshAdmin.Range("AnneeDe").value
            wshGL_BV.Range("B9").value = wshAdmin.Range("AnneeA").value
        Case "Année dernière"
            wshGL_BV.Range("B8").value = wshAdmin.Range("AnneePrecDe").value
            wshGL_BV.Range("B9").value = wshAdmin.Range("AnneePrecA").value
        Case "Dates Manuelles"
            wshGL_BV.Range("B8").value = CDate(Format("01-01-2023", "dd-mm-yyyy"))
            wshGL_BV.Range("B9").value = CDate(Format("12-31-2023", "dd-mm-yyyy"))
        Case "Toutes les dates"
            wshGL_BV.Range("B8").value = CDate(Format(wshGL_BV.Range("B3").value, "dd-mm-yyyy"))
            wshGL_BV.Range("B9").value = CDate(Format(wshGL_BV.Range("B4").value, "dd-mm-yyyy"))
    End Select
    
    Call Output_Timer_Results("modGL_TB:GL_TB_Determine_From_And_To_Date()", timerStart)

End Sub

Sub GL_TB_Setup_And_Print()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_Setup_And_Print()")
    
    Dim lastRow As Long, printRange As Range, shp As shape
    lastRow = Range("D999").End(xlUp).row + 2
    If lastRow < 4 Then Exit Sub
    Set printRange = wshGL_BV.Range("D1:G" & lastRow)
    
    Dim pagesRequired As Integer
    pagesRequired = Int((lastRow - 1) / 60) + 1
    
    Set shp = ActiveSheet.Shapes("GL_BV_Print")
    shp.Visible = msoFalse
    
    Call GL_TB_SetUp_And_Print_Document(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    Set printRange = Nothing
    Set shp = Nothing
    
    Call Output_Timer_Results("modGL_TB:GL_TB_Setup_And_Print()", timerStart)

End Sub

Sub GL_TB_Setup_And_Print_Trans()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_Setup_And_Print_Trans()")
    
    Dim lastRow As Long, printRange As Range, shp As shape
    lastRow = Range("M9999").End(xlUp).row
    If lastRow < 4 Then Exit Sub
    Set printRange = wshGL_BV.Range("L1:T" & lastRow)
    
    Dim pagesRequired As Integer
    pagesRequired = Int((lastRow - 1) / 80) + 1
    
    Set shp = ActiveSheet.Shapes("GL_BV_Print_Trans")
    shp.Visible = msoFalse
    
    Call GL_TB_SetUp_And_Print_Document(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    Call Output_Timer_Results("modGL_TB:GL_TB_Setup_And_Print_Trans()", timerStart)

End Sub

Sub GL_TB_SetUp_And_Print_Document(myPrintRange As Range, pagesTall As Integer)
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_TB:GL_TB_SetUp_And_Print_Document()")
    
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .PaperSize = xlPaperLetter
        .Orientation = xlPortrait
        .PrintArea = myPrintRange.Address 'Parameter 1
        .FitToPagesWide = 1
        .FitToPagesTall = pagesTall 'Parameter 2
        'Page Header & Footer
        .LeftHeader = ""
        .CenterHeader = "&""Aptos Narrow,Gras""&20 " & wshAdmin.Range("NomEntreprise").value
        .RightHeader = ""
        .LeftFooter = "&9&D - &T"
        .CenterFooter = ""
        .RightFooter = "&9Page &P de &N"
        'Page Margins
        .LeftMargin = Application.InchesToPoints(0.16)
        .RightMargin = Application.InchesToPoints(0.16)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .CenterHorizontally = True
        .CenterVertically = False
        'Header and Footer margins
        .HeaderMargin = Application.InchesToPoints(0.16)
        .FooterMargin = Application.InchesToPoints(0.16)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintInPlace
'        .PrintQuality = -3
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With

'    Application.Dialogs(xlDialogPrint).show
'    ActiveSheet.PageSetup.PrintArea = ""
    
    wshGL_BV.PrintOut , , 1, True, True, , , , False
 
    Call Output_Timer_Results("modGL_TB:GL_TB_SetUp_And_Print_Document()", timerStart)
 
End Sub

Sub GL_TB_Back_To_Menu()
    
    wshGL_BV.Visible = xlSheetHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
End Sub

