Attribute VB_Name = "modGL_BV"
Option Explicit

Public dynamicShape As Shape

Sub shp_GL_BV_Actualiser_Click()

    Dim ws As Worksheet
    Set ws = wshGL_BV
    
    Call GL_Trial_Balance_Build(ws, ws.Range("J1").value)

End Sub

Sub GL_Trial_Balance_Build(ws As Worksheet, dateCutOff As Date) '2024-11-18 @ 07:50
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_Trial_Balance_Build", ws.Name & " " & dateCutOff, 0)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Clear TB cells - Contents & formats
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row
    ws.Unprotect '2024-08-24 @ 16:38
    Application.EnableEvents = False
    ws.Range("D4" & ":G" & lastUsedRow + 2).Clear
    Application.EnableEvents = True

    'Clear Detail transaction section
    ws.Range("L4").CurrentRegion.offset(3, 0).Clear
    
    'Add the cut-off date in the header (printing purposes)
    Dim minDate As Date
    ws.Range("C2").value = "Au " & Format$(dateCutOff, wshAdmin.Range("B1").value)

    Application.EnableEvents = False
    ws.Range("B2").value = 3
    ws.Range("B10").value = 0
    Application.EnableEvents = True
    
    'Step # 1 - Use AdvancedFilter on GL_Trans for ALL accounts and transactions between the 2 dates
    Dim rngResultAF As Range
    Call GL_Get_Account_Trans_AF("", #7/31/2024#, dateCutOff, rngResultAF)

    'The SORT method does not sort correctly the GLNo, since there is NUMBER and NUMBER+LETTER !!!
    lastUsedRow = rngResultAF.Rows.count
    If lastUsedRow < 2 Then Exit Sub
    
    'The Chart of Account will drive the results, so the sort order is determined by COA
    Dim arr As Variant
    arr = Fn_Get_Plan_Comptable(2) 'Returns array with 2 columns (Code, Description)
    
    Dim dictSoldesParGL As Dictionary: Set dictSoldesParGL = New Dictionary
    Dim arrSolde() As Variant 'GLbalances
    ReDim arrSolde(1 To UBound(arr, 1), 1 To 2)
    Dim newRowID As Long: newRowID = 1
    Dim currRowID As Long
    
    'Parse every line of the result (AdvancedFilter in GL_Trans)
    Dim i As Long, glNo As String, MyValue As String, t1 As Currency, t2 As Currency
    For i = 2 To lastUsedRow
        glNo = rngResultAF.Cells(i, 5)
        If Not dictSoldesParGL.Exists(glNo) Then
            dictSoldesParGL.Add glNo, newRowID
            arrSolde(newRowID, 1) = glNo
            newRowID = newRowID + 1
        End If
        currRowID = dictSoldesParGL(glNo)
        'Update the summary array
        arrSolde(currRowID, 2) = arrSolde(currRowID, 2) + rngResultAF.Cells(i, 7).value - rngResultAF.Cells(i, 8).value
    Next i
    
    t1 = Application.WorksheetFunction.Sum(rngResultAF.Columns(7))
    t2 = Application.WorksheetFunction.Sum(rngResultAF.Columns(8))
    
    Dim sumDT As Currency, sumCT As Currency, GLNoPlusDesc As String
    Dim currRow As Long: currRow = 4
    ws.Range("D4:D" & UBound(arrSolde, 1)).HorizontalAlignment = xlCenter
    ws.Range("F4:G" & UBound(arrSolde, 1) + 3).HorizontalAlignment = xlRight
    
    Dim r As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        glNo = arr(i, 1)
        If glNo <> "" Then
            r = dictSoldesParGL.item(glNo) 'Get the value of the item associated with GLNo
            If r <> 0 Then
                ws.Range("D" & currRow).value = glNo
                ws.Range("E" & currRow).value = arr(i, 2)
                If arrSolde(r, 2) >= 0 Then
                    ws.Range("F" & currRow).value = Format$(arrSolde(r, 2), "###,###,##0.00")
                    sumDT = sumDT + arrSolde(r, 2)
                Else
                    ws.Range("G" & currRow).value = Format$(-arrSolde(r, 2), "###,###,##0.00")
                    sumCT = sumCT - arrSolde(r, 2)
                End If
                currRow = currRow + 1
            End If
        End If
    Next i

    currRow = currRow + 1
    ws.Range("B2").value = currRow
    
    'Unprotect the active cells of the TB area
    With ws '2024-08-21 @ 07:10
        .Unprotect
        .Range("D4:G" & currRow - 2).Locked = False
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    'Output Debit total
    Dim rng As Range
    Set rng = ws.Range("F" & currRow)
    Call GL_BV_Display_TB_Totals(rng, sumDT) 'Débit total - 2024-06-09 @ 07:51
    
    'Output Credit total
    Set rng = ws.Range("G" & currRow)
    Call GL_BV_Display_TB_Totals(rng, sumCT) 'Débit total - 2024-06-09 @ 07:51
    
    'Setup page for printing purposes
    Dim CenterHeaderTxt As String
    CenterHeaderTxt = wshAdmin.Range("NomEntreprise")
    With ActiveSheet.PageSetup
        .CenterHeader = "&""Calibri,Bold""&16 " & CenterHeaderTxt
        .PrintArea = "$D$1:$G$" & currRow
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    Application.EnableEvents = True
    
    ActiveWindow.ScrollRow = 4
    
    Application.EnableEvents = False
    ws.Range("C4").Select
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set dictSoldesParGL = Nothing
    Set rng = Nothing
    
    Call Log_Record("modGL_BV:GL_Trial_Balance_Build", "", startTime)

End Sub

Sub GL_BV_Display_TB_Totals(rng As Range, t As Currency) '2024-06-09 @ 07:45

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Display_TB_Totals", "", 0)
    
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets("GL_BV")
    
    With rng
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        .value = t
        .Font.Bold = True
        .NumberFormat = "#,##0.00 $"
    End With
    
    Call Log_Record("modGL_BV:GL_BV_Display_TB_Totals", "", startTime)

End Sub

Sub GL_BV_Display_Trans_For_Selected_Account(GLAcct As String, GLDesc As String, minDate As Date, maxDate As Date) 'Display GL Trans for a specific account

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Display_Trans_For_Selected_Account", GLAcct & " De " & minDate & " à " & maxDate, 0)
    
    Dim ws As Worksheet: Set ws = wshGL_BV
    
    'Clear the display area & display the account number & description
    With ws
        .Range("L4:T99999").Clear '2024-06-08 @ 15:28
        .Range("L2").value = "Du " & Format$(minDate, wshAdmin.Range("B1").value) & " au " & Format$(maxDate, wshAdmin.Range("B1").value)
    
        .Range("L4").Font.Bold = True
        .Range("L4").value = GLAcct & " - " & GLDesc
        .Range("B6").value = GLAcct
        .Range("B7").value = GLDesc
    End With
    
    'Use the AdvancedFilter Result already prepared for TB
    Dim row As Range, foundRow As Long, lastResultUsedRow As Long
    lastResultUsedRow = wshGL_Trans.Cells(wshGL_Trans.Rows.count, "P").End(xlUp).row
    If lastResultUsedRow <= 2 Then
        GoTo Exit_Sub
    End If
    foundRow = 0
    
    'Find the first occurence of GlACct in AdvancedFilter Results on GL_Trans
    Dim searchRange As Range: Set searchRange = wshGL_Trans.Range("T1:T" & lastResultUsedRow)
    Dim foundCell As Range: Set foundCell = searchRange.Find(What:=GLAcct, LookIn:=xlValues, LookAt:=xlWhole)
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
    
    Dim d As Date, OK As Long
    
    Application.ScreenUpdating = False
    
    With ws
        'On assume que les résultats de GL_Trans sont triés par numéro de compte, par date & par no écriture
        Do Until wshGL_Trans.Range("T" & foundRow).value <> GLAcct
            'Traitement des transactions détaillées
            d = Format$(wshGL_Trans.Range("Q" & foundRow).Value2, wshAdmin.Range("B1").value)
            If d >= minDate And d <= maxDate Then
                .Range("M" & rowGLDetail).value = wshGL_Trans.Range("Q" & foundRow).Value2
                .Range("M" & rowGLDetail).NumberFormat = wshAdmin.Range("B1").value
                .Range("N" & rowGLDetail).value = wshGL_Trans.Range("P" & foundRow).value
                .Range("N" & rowGLDetail).HorizontalAlignment = xlCenter
                .Range("O" & rowGLDetail).value = wshGL_Trans.Range("R" & foundRow).value
                .Range("P" & rowGLDetail).value = wshGL_Trans.Range("S" & foundRow).value
                .Range("Q" & rowGLDetail).NumberFormat = "#,##0.00"
                .Range("Q" & rowGLDetail).value = wshGL_Trans.Range("V" & foundRow).value
                .Range("R" & rowGLDetail).NumberFormat = "#,##0.00"
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
        
    Dim rng As Range
    lastResultUsedRow = ws.Cells(ws.Rows.count, "M").End(xlUp).row
    Set rng = ws.Range("M5:T" & lastResultUsedRow)
    
    'Fix font size & Family for the detailled transactions list
    Call Fix_Font_Size_And_Family(rng, "Aptos Narrow", 9)
    
    'Set columns width for the detailled transactions list
    Set rng = ws.Range("M5:M" & lastResultUsedRow)
    rng.ColumnWidth = 9
    rng.HorizontalAlignment = xlCenter
    
    Set rng = ws.Range("N5:N" & lastResultUsedRow)
    rng.ColumnWidth = 7
    Set rng = ws.Range("O5:O" & lastResultUsedRow)
    rng.ColumnWidth = 45
    Set rng = ws.Range("P5:P" & lastResultUsedRow)
    rng.ColumnWidth = 20
    Set rng = ws.Range("Q5:S" & lastResultUsedRow)
    rng.ColumnWidth = 15
    Set rng = ws.Range("T5:T" & lastResultUsedRow)
    rng.ColumnWidth = 35

    Dim visibleRows As Long
    visibleRows = ActiveWindow.visibleRange.Rows.count
    If lastResultUsedRow > visibleRows Then
        ActiveWindow.ScrollRow = lastResultUsedRow - visibleRows + 5 'Move to the bottom of the worksheet
    Else
        ActiveWindow.ScrollRow = 1
    End If
    
    'Create a Conditional Formating for the displayed transactions
    ws.Unprotect
    With ws.Range("M5:T" & lastResultUsedRow)
        On Error Resume Next
        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=ET($M5<>"""";MOD(LIGNE();2)=1)"
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
        End With
        .FormatConditions(1).StopIfTrue = False
        On Error GoTo 0
    End With
    
    'Unprotect the active cells of the transactions details area
    With wshGL_BV '2024-08-21 @ 07:15
        .Unprotect
        .Range("L4:T" & lastResultUsedRow).Locked = False
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

    Call GL_BV_Ajouter_Shape_Retour
    
Exit_Sub:

    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set foundCell = Nothing
    Set rng = Nothing
    Set searchRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modGL_BV:GL_BV_Display_Trans_For_Selected_Account", "", startTime)

End Sub

Sub GL_BV_Sub_Totals(glNo As String, GLDesc As String, s As Currency)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Sub_Totals", "", 0)

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
    
    Call Log_Record("modGL_BV:GL_BV_Sub_Totals", "", startTime)

End Sub

Sub GL_BV_Determine_From_And_To_Date(period As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Determine_From_And_To_Date", "", 0)

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
            wshGL_BV.Range("B8").value = CDate(Format$("07-31-2024", "dd/mm/yyyy"))
            wshGL_BV.Range("B9").value = CDate(Format$("07-31-2025", "dd/mm/yyyy"))
        Case "Toutes les dates"
            wshGL_BV.Range("B8").value = CDate(Format$(wshGL_BV.Range("B3").value, "dd/mm/yyyy"))
            wshGL_BV.Range("B9").value = CDate(Format$(wshGL_BV.Range("B4").value, "dd/mm/yyyy"))
    End Select
    
    Call Log_Record("modGL_BV:GL_BV_Determine_From_And_To_Date", "", startTime)

End Sub

Sub shp_GL_BV_Impression_BV_Click()

    Call GL_BV_Setup_And_Print

End Sub

Sub GL_BV_Setup_And_Print()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Setup_And_Print", "", 0)
    
    Dim lastRow As Long
    lastRow = wshGL_BV.Cells(wshGL_BV.Rows.count, "D").End(xlUp).row + 2
    If lastRow < 4 Then Exit Sub
    
    Dim printRange As Range
    Set printRange = wshGL_BV.Range("D1:G" & lastRow)
    
    Dim pagesRequired As Long
    pagesRequired = Int((lastRow - 1) / 60) + 1
    
    Dim shp As Shape: Set shp = wshGL_BV.Shapes("GL_BV_Print")
    shp.Visible = msoFalse
    
    Call GL_BV_SetUp_And_Print_Document(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    'Libérer la mémoire
    Set printRange = Nothing
    Set shp = Nothing
    
    Call Log_Record("modGL_BV:GL_BV_Setup_And_Print", "", startTime)

End Sub

Sub shp_GL_BV_Setup_And_Print_Trans_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:shp_GL_BV_Setup_And_Print_Trans_Click", "", 0)
    
    Call GL_BV_Setup_And_Print_Trans

    Call Log_Record("modGL_BV:shp_GL_BV_Setup_And_Print_Trans_Click", "", startTime)

End Sub

Sub GL_BV_Setup_And_Print_Trans()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Setup_And_Print_Trans", "", 0)
    
    Dim lastRow As Long
    lastRow = wshGL_BV.Cells(wshGL_BV.Rows.count, "M").End(xlUp).row
    If lastRow < 4 Then Exit Sub
    
    Dim printRange As Range
    Set printRange = wshGL_BV.Range("L1:T" & lastRow)
    
    Dim pagesRequired As Long
    pagesRequired = Int((lastRow - 1) / 80) + 1
    
    Dim shp As Shape: Set shp = ActiveSheet.Shapes("GL_BV_Print_Trans")
    shp.Visible = msoFalse
    
    Call GL_BV_SetUp_And_Print_Document(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    'Libérer la mémoire
    Set printRange = Nothing
    Set shp = Nothing
    
    Call Log_Record("modGL_BV:GL_BV_Setup_And_Print_Trans", "", startTime)

End Sub

Sub GL_BV_SetUp_And_Print_Document(myPrintRange As Range, pagesTall As Long)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_SetUp_And_Print_Document", "", 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    With ActiveSheet.PageSetup
'        .PrintTitleRows = ""
'        .PrintTitleColumns = ""
        .PaperSize = xlPaperLetter
        .Orientation = xlPortrait
        .PrintArea = myPrintRange.Address 'Parameter 1
        .FitToPagesWide = 1
        .FitToPagesTall = pagesTall 'Parameter 2
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1 is completed", -1)
        
        'Page Header & Footer
'        .LeftHeader = ""
        .CenterHeader = "&""Aptos Narrow,Gras""&18 " & wshAdmin.Range("NomEntreprise").value
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1.A is completed", -1)
        
'        .RightHeader = ""
        .LeftFooter = "&9&D - &T"
'        .CenterFooter = ""
        .RightFooter = "&9Page &P de &N"
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1.B is completed", -1)
        
        'Page Margins
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 is starting", -1)
        .LeftMargin = Application.InchesToPoints(0.16)
        .RightMargin = Application.InchesToPoints(0.16)
         Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Left & Right) margins", -1)
         
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
         Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Top & Bottom) margins", -1)
         
        .CenterHorizontally = True
        .CenterVertically = False
         Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Center Horizontal & Vertical)", -1)
         
        'Header and Footer margins
        .HeaderMargin = Application.InchesToPoints(0.16)
        .FooterMargin = Application.InchesToPoints(0.16)
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Header & Footer) margins", -1)
        
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintInPlace
'        .PrintQuality = -3
'        .Draft = False
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = False
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Speed Measure", -1)
    
    wshGL_BV.PrintPreview '2024-08-15 @ 14:53
 
    Call Log_Record("modGL_BV:GL_BV_SetUp_And_Print_Document", "", startTime)
 
End Sub

Sub Erase_Non_Required_Shapes() '2024-08-15 @ 14:42

    Dim ws As Worksheet: Set ws = wshGL_BV
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        If InStr(shp.Name, "Rounded Rectangle ") Then
            shp.Delete
        End If
    Next shp

    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub Test_Get_All_Shapes() '2024-08-15 @ 14:42

    Dim ws As Worksheet: Set ws = wshGL_BV
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub wshGL_BV_Display_JE_Trans_With_Shape()

    Call wshGL_BV_Create_Dynamic_Shape
    Call wshGL_BV_Adjust_The_Shape
    Call GL_BV_Show_Dynamic_Shape
    
End Sub

Sub wshGL_BV_Create_Dynamic_Shape()

    'Check if the shape has already been created
    If dynamicShape Is Nothing Then
        'Create the text box shape
        wshGL_BV.Unprotect
        Set dynamicShape = wshGL_BV.Shapes.AddShape(msoShapeRoundedRectangle, 2000, 100, 600, 100)
    End If

    'Libérer la mémoire
'    Set dynamicShape = Nothing
    
End Sub

Sub wshGL_BV_Adjust_The_Shape()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:wshGL_BV_Adjust_The_Shape", "", 0)
    
    Dim lastResultRow As Long
    lastResultRow = wshGL_Trans.Cells(wshGL_Trans.Rows.count, "AC").End(xlUp).row
    If lastResultRow < 2 Then Exit Sub
    
    Dim rowSelected As Long
    rowSelected = wshGL_BV.Range("B10").value
    
    Dim texteOneLine As String, texteFull As String
    
    Dim i As Long, maxLength As Long
    With wshGL_Trans
        For i = 2 To lastResultRow
            If i = 2 Then
                texteFull = "Entrée #: " & .Range("AC2").value & vbCrLf
                texteFull = texteFull & "Desc    : " & .Range("AE2").value & vbCrLf
                If Trim$(.Range("AF2").value) <> "" Then
                    texteFull = texteFull & "Source  : " & .Range("AF2").value & vbCrLf & vbCrLf
                Else
                    texteFull = texteFull & vbCrLf
                End If
            End If
            texteOneLine = Fn_Pad_A_String(.Range("AG" & i).value, " ", 5, "R") & _
                            " - " & Fn_Pad_A_String(.Range("AH" & i).value, " ", 35, "R") & _
                            "  " & Fn_Pad_A_String(Format$(.Range("AI" & i).value, "#,##0.00 $"), " ", 14, "L") & _
                            "  " & Fn_Pad_A_String(Format$(.Range("AJ" & i).value, "#,##0.00 $"), " ", 14, "L")
            If Trim$(.Range("AF" & i).value) = Trim$(wshGL_BV.Range("B6").value) Then
                texteOneLine = " * " & texteOneLine
            Else
                texteOneLine = "   " & texteOneLine
            End If
            texteOneLine = Fn_Pad_A_String(texteOneLine, " ", 79, "R")
            If Trim$(.Range("AK" & i).value) <> "" Then
                texteOneLine = texteOneLine & Trim$(.Range("AK" & i).value)
            End If
            If Len(texteOneLine) > maxLength Then
                maxLength = Len(texteOneLine)
            End If
            texteFull = texteFull & texteOneLine & vbCrLf
        Next i
    End With
    If Right$(texteFull, Len(texteFull) - 1) = vbCrLf Then
        texteFull = Left$(texteFull, Len(texteFull) - 2)
    End If
    
    Dim dynamicShape As Shape: Set dynamicShape = wshGL_BV.Shapes("JE_Detail_Trans")

    'Set shape properties
    With dynamicShape
        .Fill.ForeColor.RGB = RGB(249, 255, 229)
        .Line.Weight = 2
        .Line.ForeColor.RGB = vbBlue
        .TextFrame.Characters.Text = texteFull
        .TextFrame.Characters.Font.Color = vbBlack
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.size = 10
        .TextFrame.MarginLeft = 4
        .TextFrame.MarginRight = 4
        .TextFrame.MarginTop = 3
        .TextFrame.MarginBottom = 3
        If maxLength < 80 Then maxLength = 80
        .Width = ((maxLength * 6.1))
'            .Height = ((lastResultRow + 4) * 12) + 3 + 3
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .Left = wshGL_BV.Range("N" & rowSelected).Left + 4
        .Top = wshGL_BV.Range("N" & rowSelected + 1).Top + 4
    End With
        
    'Libérer la mémoire
    Set dynamicShape = Nothing
      
    Call Log_Record("modGL_BV:wshGL_BV_Adjust_The_Shape", "", startTime)
      
End Sub

Sub GL_BV_Show_Dynamic_Shape()

    Dim shp As Shape: Set shp = wshGL_BV.Shapes("JE_Detail_Trans")
    shp.Visible = msoTrue
    
'    If Not dynamicShape Is Nothing Then
'        dynamicShape.Visible = True
'    End If

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub GL_BV_Hide_Dynamic_Shape()

    Dim shp As Shape: Set shp = wshGL_BV.Shapes("JE_Detail_Trans")
    shp.Visible = msoFalse

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub shp_GL_BV_Exit_Click()

    Call GL_BV_Back_To_Menu

End Sub

Sub GL_BV_Back_To_Menu()
    
    Call Erase_Non_Required_Shapes
    
    wshGL_BV.Visible = xlSheetHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
End Sub


