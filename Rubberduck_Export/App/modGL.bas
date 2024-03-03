Attribute VB_Name = "modGL"
Option Explicit

Sub UpdateBV() 'Button 'Actualiser'
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call GL_Trans_Import_All
    
    Application.ScreenUpdating = True
    
    Dim minDate As Date, dateCutOff As Date, lastRow As Long, solde As Currency
    Dim planComptable As Range
    Set planComptable = wshAdmin.Range("dnrPlanComptable")
    
    'Clear Detail transaction section
    wshGL_BV.Range("L4:T9999").ClearContents
'    wshGL_BV.Range("L4:T99999").ClearComments
    
    'Clear contents & formats for TB cells
    lastRow = wshGL_BV.Range("D99999").End(xlUp).row
    With wshGL_BV.Range("D4" & ":G" & lastRow + 2)
        .ClearContents
        .ClearFormats
    End With
    
    'Add the cut-off date in the header (printing purposes)
    wshGL_BV.Range("C2").value = "Au " & CDate(Format(wshGL_BV.Range("B4").value, "dd-mm-yyyy"))

    minDate = CDate("01/01/2023")
    dateCutOff = CDate(wshGL_BV.Range("J1").value)
    wshGL_BV.Range("B2").value = 4
    
    Call GL_Trans_Advanced_Filter("", minDate, dateCutOff)
    
    lastRow = wshGL_Trans.Range("T99999").End(xlUp).row
    If lastRow < 2 Then Exit Sub
    Dim r As Long, BreakGLNo As String, oldDesc As String
    BreakGLNo = wshGL_Trans.Range("T2").value
    oldDesc = wshGL_Trans.Range("U2").value
    
    For r = 2 To lastRow
        If wshGL_Trans.Range("T" & r).value <> BreakGLNo Then
            Call GL_Trans_Sub_Total(BreakGLNo, oldDesc, solde)
            BreakGLNo = wshGL_Trans.Range("T" & r).value
            oldDesc = wshGL_Trans.Range("U" & r).value
            solde = 0
        End If
        solde = solde + wshGL_Trans.Range("V" & r).value - wshGL_Trans.Range("W" & r).value
    Next r
    
    Call GL_Trans_Sub_Total(BreakGLNo, oldDesc, solde)
    
    r = wshGL_BV.Range("B2").value + 1
    
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

    wshGL_BV.Range("B2").value = r - 1
  
    Application.EnableEvents = True
  
End Sub

Sub DisplayTBTotals(r As Long, c As Long)

    Dim sumRange As Range

    With wshGL_BV.Cells(r, c).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With wshGL_BV.Cells(r, c).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    wshGL_BV.Cells(r, c).Font.Bold = True
    wshGL_BV.Cells(r, c).NumberFormat = "#,##0.00 $"
    Set sumRange = Range(Cells(4, c), Cells(r - 1, c))
    
    wshGL_BV.Cells(r, c).value = Application.WorksheetFunction.Sum(sumRange)

End Sub

Sub GLTransDisplay(GLAcct As String, GLDesc As String, minDate As Date, maxDate As Date) 'Display GL Trans for a specific account

    'Clear the display area & display the account number & description
    wshGL_BV.Range("M4:T99999").ClearFormats
    wshGL_BV.Range("M4:T99999").ClearContents
    wshGL_BV.Range("L2").value = "Du " & minDate & " au " & maxDate
    
    wshGL_BV.Range("L4").Font.Bold = True
    wshGL_BV.Range("L4").value = GLAcct & " - " & GLDesc
    wshGL_BV.Range("B6").value = GLAcct
    wshGL_BV.Range("B7").value = GLDesc
    
    'Use the Advanced Filter Result already prepared for TB
    Dim row As Range, foundRow As Long, lastResultRow As Long
    lastResultRow = wshGL_Trans.Range("T99999").End(xlUp).row
    foundRow = 0
    
    'Loop through each row in the search range - RMV - 2024-01-05 - À améliorer
    For Each row In wshGL_Trans.Range("T2:T" & lastResultRow).Rows
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
    wshGL_BV.Range("S4").value = 0
    wshGL_BV.Range("S4").Font.Bold = True
    wshGL_BV.Range("S4").NumberFormat = "#,##0.00 $"
    With wshGL_BV.Range("S4").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    Dim d As Date, OK As Integer
    
    With wshGL_BV
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
            .Range("S" & rowGLDetail).value = wshGL_BV.Range("S" & rowGLDetail - 1).value + _
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

        wshGL_BV.Range("S" & rowGLDetail - 1).Font.Bold = True
        With wshGL_BV.Range("S" & rowGLDetail - 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With

End Sub

Sub GL_Trans_Advanced_Filter(GLNo As String, minDate As Date, maxDate As Date)

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
        
        Dim lastResultRow
        lastResultRow = .Range("S999999").End(xlUp).row
        If lastResultRow < 3 Then GoTo NoSort
        With .Sort
            .SortFields.clear
            .SortFields.add Key:=wshGL_Trans.Range("T1"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortTextAsNumbers 'Sort Based On GLNo
            .SortFields.add Key:=wshGL_Trans.Range("Q1"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.add Key:=wshGL_Trans.Range("P1"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On JE Number
            .SetRange wshGL_Trans.Range("P2:Y" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
    End With

NoSort:

End Sub

Sub GL_Trans_Sub_Total(GLNo As String, GLDesc As String, s As Currency)

    Dim r As Long
    r = wshGL_BV.Range("B2").value
    wshGL_BV.Range("D" & r).HorizontalAlignment = xlCenter
    wshGL_BV.Range("D" & r).value = GLNo
    wshGL_BV.Range("E" & r).value = GLDesc
    If s > 0 Then
        wshGL_BV.Range("F" & r).value = s
    ElseIf s < 0 Then
        wshGL_BV.Range("G" & r).value = -s
    End If
    With wshGL_BV.Range("D" & r & ":G" & r).Font
        .name = "Aptos Narrow"
        .Size = 11
    End With
    wshGL_BV.Range("B2").value = wshGL_BV.Range("B2").value + 1
    
End Sub

Sub DetermineFromAndToDate(period As String)

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
'            Debug.Print "Period is '" & period & "' so MinDate = " & wshGL_BV.Range("B8").value & _
'                "  maxDate = " & wshGL_BV.Range("B9").value
End Sub

Sub SetUpAndPrintTransactions()
    
    Dim lastRow As Long, printRange As Range, Shp As Shape
    lastRow = Range("M9999").End(xlUp).row
    If lastRow < 4 Then Exit Sub
    Set printRange = wshGL_BV.Range("L1:T" & lastRow)
    
    Dim pagesRequired As Integer
    pagesRequired = Int((lastRow - 1) / 60) + 1
    
    Set Shp = ActiveSheet.Shapes("ImprimerTransactions")
    Shp.Visible = msoFalse
    
    Call SetUpAndPrintDocument(printRange, pagesRequired)
    
    Shp.Visible = msoTrue
    
End Sub

Sub SetUpAndPrintTB()
    
    Dim lastRow As Long, printRange As Range, Shp As Shape
    lastRow = Range("D9999").End(xlUp).row + 2
    If lastRow < 4 Then Exit Sub
    Set printRange = wshGL_BV.Range("D1:G" & lastRow)
    
    Dim pagesRequired As Integer
    pagesRequired = Int((lastRow - 1) / 60) + 1
    
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
 
'    wshGL_BV.PrintOut , , , True, True, , , , False
 
End Sub

Sub Back_To_GL_Menu()

    wshMenuCOMPTA.Activate
    wshMenuCOMPTA.Range("A1").Select
    
End Sub


