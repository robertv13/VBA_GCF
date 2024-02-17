Attribute VB_Name = "modGL"
Option Explicit

Sub UpdateBV() 'Button 'Actualiser'
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call GL_Trans_Import_All
    
    Application.ScreenUpdating = True
    
    Dim minDate As Date, dateCutOff As Date, LastRow As Long, solde As Currency
    Dim planComptable As Range
    Set planComptable = wshAdmin.Range("dnrPlanComptable")
    
    'Clear Detail transaction section
    wshBV.Range("L4:T9999").ClearContents
'    wshBV.Range("L4:T99999").ClearComments
    
    'Clear contents & formats for TB cells
    LastRow = wshBV.Range("D99999").End(xlUp).row
    With wshBV.Range("D4" & ":G" & LastRow + 2)
        .ClearContents
        .ClearFormats
    End With
    
    'Add the cut-off date in the header (printing purposes)
    wshBV.Range("C2").value = "Au " & CDate(Format(wshBV.Range("B4").value, "dd-mm-yyyy"))

    minDate = CDate("01/01/2023")
    dateCutOff = CDate(wshBV.Range("J1").value)
    wshBV.Range("B2").value = 4
    
    Call AdvancedFilterGLTrans("", minDate, dateCutOff)
    
    LastRow = wshGL_Trans.Range("T99999").End(xlUp).row
    If LastRow < 2 Then Exit Sub
    Dim r As Long, BreakGLNo As String, oldDesc As String
    BreakGLNo = wshGL_Trans.Range("T2").value
    oldDesc = wshGL_Trans.Range("U2").value
    
    For r = 2 To LastRow
        If wshGL_Trans.Range("T" & r).value <> BreakGLNo Then
            Call GL_Trans_Sub_Total(BreakGLNo, oldDesc, solde)
            BreakGLNo = wshGL_Trans.Range("T" & r).value
            oldDesc = wshGL_Trans.Range("U" & r).value
            solde = 0
        End If
        solde = solde + wshGL_Trans.Range("V" & r).value - wshGL_Trans.Range("W" & r).value
    Next r
    
    Call GL_Trans_Sub_Total(BreakGLNo, oldDesc, solde)
    
    r = wshBV.Range("B2").value + 1
    
    DisplayTBTotals r, 6 'Débit
    DisplayTBTotals r, 7 'Crédit
    
    'Setup page for printing purposes
    Dim CenterHeaderTxt As String
    CenterHeaderTxt = wshAdmin.Range("NomEntreprise")
    With ActiveSheet.PageSetup
        .CenterHeader = "&""Calibri,Bold""&20 " & CenterHeaderTxt
        .PrintArea = "$D$1:$G$" & r
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    wshBV.Range("B2").value = r - 1
  
    Application.EnableEvents = True
  
End Sub

Sub DisplayTBTotals(r As Long, c As Long)

    Dim sumRange As Range

    With wshBV.Cells(r, c).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With wshBV.Cells(r, c).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    wshBV.Cells(r, c).Font.Bold = True
    wshBV.Cells(r, c).NumberFormat = "#,##0.00 $"
    Set sumRange = Range(Cells(4, c), Cells(r - 1, c))
    
    wshBV.Cells(r, c).value = Application.WorksheetFunction.Sum(sumRange)

End Sub

Sub GLTransDisplay(GLAcct As String, GLDesc As String, minDate As Date, maxDate As Date) 'Display GL Trans for a specific account

    'Clear the display area & display the account number & description
    wshBV.Range("M4:T99999").ClearFormats
    wshBV.Range("M4:T99999").ClearContents
    wshBV.Range("L2").value = "Du " & minDate & " au " & maxDate
    
    wshBV.Range("L4").Font.Bold = True
    wshBV.Range("L4").value = GLAcct & " - " & GLDesc
    wshBV.Range("B6").value = GLAcct
    wshBV.Range("B7").value = GLDesc
    
    'Use the Advanced Filter Result already prepared for TB
    Dim row As Range, foundRow As Long, LastResultRow As Long
    LastResultRow = wshGL_Trans.Range("T99999").End(xlUp).row
    foundRow = 0
    
    'Loop through each row in the search range - RMV - 2024-01-05 - À améliorer
    For Each row In wshGL_Trans.Range("T2:T" & LastResultRow).Rows
        If row.Cells(1, 1).value = GLAcct Then
            'Store the row number and exit the loop
            foundRow = row.row
            Exit For
        End If
    Next row
    
    ' Check if the target value was found
    If foundRow = 0 Then
        MsgBox "Il n'existe aucune transaction pour ce compte (période choisie)."
        Exit Sub
    End If
    
    Dim rowGLDetail As Long
    rowGLDetail = 5
    wshBV.Range("S4").value = 0
    wshBV.Range("S4").Font.Bold = True
    wshBV.Range("S4").NumberFormat = "#,##0.00 $"
    With wshBV.Range("S4").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    Dim d As Date, OK As Integer
    
    With wshBV
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
            .Range("S" & rowGLDetail).value = wshBV.Range("S" & rowGLDetail - 1).value + _
                wshGL_Trans.Range("V" & foundRow).value - wshGL_Trans.Range("W" & foundRow).value
'            With .Range("S" & rowGLDetail).Font
'                .Name = "Aptos Narrow"
'                .Size = 11
'            End With
            .Range("T" & rowGLDetail).Value2 = wshGL_Trans.Range("X" & foundRow).value
            foundRow = foundRow + 1
            rowGLDetail = rowGLDetail + 1
            OK = OK + 1
        Else
            foundRow = foundRow + 1
'            If d < minDate Then Debug.Print Tab(5); d & " < " & minDate
'            If d > maxDate Then Debug.Print Tab(5); d & " > " & maxDate
        End If
    Loop
    End With

        wshBV.Range("S" & rowGLDetail - 1).Font.Bold = True
        With wshBV.Range("S" & rowGLDetail - 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With

End Sub

Sub AdvancedFilterGLTrans(GLNo As String, minDate As Date, maxDate As Date)

    With wshGL_Trans
        Dim rgResult As Range, rgData As Range, rgCriteria As Range, rgCopyToRange As Range
        Set rgResult = .Range("P2").CurrentRegion
        Call ClearRangeBorders(rgResult)
        rgResult.ClearContents
        Set rgData = .Range("A1").CurrentRegion
        .Range("L3").value = GLNo
        .Range("M3").value = ">=" & Format(minDate, "mm-dd-yyyy")
        .Range("N3").value = "<=" & Format(maxDate, "mm-dd-yyyy")
        
        Set rgCriteria = .Range("L2:N3")
        Set rgCopyToRange = .Range("P1")
        
        rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgCopyToRange
        
        Dim LastResultRow
        LastResultRow = .Range("S999999").End(xlUp).row
        If LastResultRow < 3 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshGL_Trans.Range("T1"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortTextAsNumbers 'Sort Based On GLNo
            .SortFields.Add Key:=wshGL_Trans.Range("Q1"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add Key:=wshGL_Trans.Range("P1"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On JE Number
            .SetRange wshGL_Trans.Range("P2:Y" & LastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
    End With

NoSort:

End Sub

Sub GL_Trans_Sub_Total(GLNo As String, GLDesc As String, s As Currency)

    Dim r As Long
    r = wshBV.Range("B2").value
    wshBV.Range("D" & r).HorizontalAlignment = xlCenter
    wshBV.Range("D" & r).value = GLNo
    wshBV.Range("E" & r).value = GLDesc
    If s > 0 Then
        wshBV.Range("F" & r).value = s
    ElseIf s < 0 Then
        wshBV.Range("G" & r).value = -s
    End If
    With wshBV.Range("D" & r & ":G" & r).Font
        .Name = "Aptos Narrow"
        .Size = 11
    End With
    wshBV.Range("B2").value = wshBV.Range("B2").value + 1
    
End Sub

Sub GL_Trans_Import_All() '2024-02-14 @ 06:14
    
    Application.ScreenUpdating = False
    
    Dim saveLastRow As Long
    saveLastRow = wshGL_Trans.Range("A999999").End(xlUp).row
    
    'Clear all cells, but the headers, in the target worksheet
    wshGL_Trans.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import GLTrans from 'GCF_DB_Sortie.xlsx'
    Dim sourceWorkbook As String, sourceTab As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx" '2024-02-13 @ 15:09
    sourceTab = "GL_Trans"
                     
    'Set up source and destination ranges
    Dim sourceRange As Range
    Set sourceRange = Workbooks.Open(sourceWorkbook).Worksheets(sourceTab).UsedRange

    Dim destinationRange As Range
    Set destinationRange = wshGL_Trans.Range("A1")

    'Copy data, using Range to Range, then close the BD_Sortie file
    sourceRange.Copy destinationRange
    wshGL_Trans.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Workbooks("GCF_BD_Sortie.xlsx").Close SaveChanges:=False

    Dim LastRow As Long
    LastRow = wshGL_Trans.Range("A999999").End(xlUp).row
    
    'Adjust Formats for all new rows
    With wshGL_Trans
        .Range("A" & 2 & ":J" & LastRow).HorizontalAlignment = xlCenter
        .Range("B" & 2 & ":B" & LastRow).NumberFormat = "dd/mm/yyyy"
        .Range("C" & 2 & ":C" & LastRow & _
            ", D" & 2 & ":D" & LastRow & _
            ", F" & 2 & ":F" & LastRow & _
            ", I" & 2 & ":I" & LastRow) _
                .HorizontalAlignment = xlLeft
        With .Range("G" & 2 & ":H" & LastRow)
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00 $"
        End With
        With .Range("A" & 2 & ":A" & LastRow) _
            .Range("J" & 2 & ":J" & LastRow).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
    Dim firstRowJE As Long, lastRowJE As Long
    Dim r As Long
    
    For r = 2 To LastRow 'RMV - 2024-01-05
        With wshGL_Trans.Range("A" & r & ":J" & r) 'No_EJ & No.Ligne
            .Font.ThemeColor = xlThemeColorLight1
            .Font.TintAndShade = -4.99893185216834E-02
            .Interior.Pattern = xlNone
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
        End With
        wshGL_Trans.Range("J" & r).formula = "=ROW()"
    Next r

    Application.ScreenUpdating = True
    
End Sub

Sub DetermineFromAndToDate(period As String)

    Select Case period
        Case "Mois"
            wshBV.Range("B8").value = wshAdmin.Range("MoisDe").value
            wshBV.Range("B9").value = wshAdmin.Range("MoisA").value
        Case "Mois dernier"
            wshBV.Range("B8").value = wshAdmin.Range("MoisPrecDe").value
            wshBV.Range("B9").value = wshAdmin.Range("MoisPrecA").value
        Case "Trimestre"
            wshBV.Range("B8").value = wshAdmin.Range("TrimDe").value
            wshBV.Range("B9").value = wshAdmin.Range("TrimA").value
        Case "Trimestre dernier"
            wshBV.Range("B8").value = wshAdmin.Range("TrimPrecDe").value
            wshBV.Range("B9").value = wshAdmin.Range("TrimPrecA").value
        Case "Année"
            wshBV.Range("B8").value = wshAdmin.Range("AnneeDe").value
            wshBV.Range("B9").value = wshAdmin.Range("AnneeA").value
        Case "Année dernière"
            wshBV.Range("B8").value = wshAdmin.Range("AnneePrecDe").value
            wshBV.Range("B9").value = wshAdmin.Range("AnneePrecA").value
        Case "Dates Manuelles"
            wshBV.Range("B8").value = CDate(Format("01-01-2023", "dd-mm-yyyy"))
            wshBV.Range("B9").value = CDate(Format("12-31-2023", "dd-mm-yyyy"))
        Case "Toutes les dates"
            wshBV.Range("B8").value = CDate(Format(wshBV.Range("B3").value, "dd-mm-yyyy"))
            wshBV.Range("B9").value = CDate(Format(wshBV.Range("B4").value, "dd-mm-yyyy"))
    End Select
'            Debug.Print "Period is '" & period & "' so MinDate = " & wshBV.Range("B8").value & _
'                "  maxDate = " & wshBV.Range("B9").value
End Sub

Sub SetUpAndPrintTransactions()
    
    Dim LastRow As Long, printRange As Range, Shp As Shape
    LastRow = Range("M9999").End(xlUp).row
    If LastRow < 4 Then Exit Sub
    Set printRange = wshBV.Range("L1:T" & LastRow)
    
    Dim pagesRequired As Integer
    pagesRequired = Int((LastRow - 1) / 60) + 1
    
    Set Shp = ActiveSheet.Shapes("ImprimerTransactions")
    Shp.Visible = msoFalse
    
    Call SetUpAndPrintDocument(printRange, pagesRequired)
    
    Shp.Visible = msoTrue
    
End Sub

Sub SetUpAndPrintTB()
    
    Dim LastRow As Long, printRange As Range, Shp As Shape
    LastRow = Range("D9999").End(xlUp).row + 2
    If LastRow < 4 Then Exit Sub
    Set printRange = wshBV.Range("D1:G" & LastRow)
    
    Dim pagesRequired As Integer
    pagesRequired = Int((LastRow - 1) / 60) + 1
    
    Set Shp = ActiveSheet.Shapes("ImprimerBV")
    Shp.Visible = msoFalse
    
    Call SetUpAndPrintDocument(printRange, pagesRequired)
    
    Shp.Visible = msoTrue
    
End Sub

Sub SetUpAndPrintDocument(myPrintRange As Range, pagesTall As Integer)
    
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

    Application.Dialogs(xlDialogPrint).show
    ActiveSheet.PageSetup.PrintArea = ""
 
'    wshBV.PrintOut , , , True, True, , , , False
 
End Sub

Sub Back_To_GL_Menu()

    wshMenuCOMPTA.Activate
    wshMenuCOMPTA.Range("A1").Select
    
End Sub


