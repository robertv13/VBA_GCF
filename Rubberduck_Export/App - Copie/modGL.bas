Attribute VB_Name = "modGL"
Option Explicit

Sub GL_Build_TB() '2024-03-05 @ 13:34
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Call GL_Trans_Import_All is it mandatory ???
    
    Dim minDate As Date, dateCutOff As Date, lastUsedRow As Long, solde As Currency
    Dim planComptable As Range
    Set planComptable = wshAdmin.Range("dnrPlanComptableDescription")
    
    'Clear Detail transaction section
    wshGL_BV.Range("L4:T9999").Clearcontents
    With wshGL_BV.Range("S4:S9999").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Clear contents & formats for TB cells
    lastUsedRow = wshGL_BV.Range("D99999").End(xlUp).row
    wshGL_BV.Range("D4" & ":G" & lastUsedRow + 2).clear
    
    'Add the cut-off date in the header (printing purposes)
    wshGL_BV.Range("C2").Value = "Au " & CDate(Format(wshGL_BV.Range("B4").Value, "dd-mm-yyyy"))

    minDate = CDate("01/01/2023")
    dateCutOff = CDate(wshGL_BV.Range("J1").Value)
    wshGL_BV.Range("B2").Value = 3
    wshGL_BV.Range("B10").Value = 0
    
    Call GL_Trans_Advanced_Filter("", minDate, dateCutOff) 'Get all transactions between the 2 dates
    
    lastUsedRow = wshGL_Trans.Range("T999999").End(xlUp).row
    If lastUsedRow < 2 Then Exit Sub
    Dim r As Long, BreakGLNo As String, oldDesc As String
    BreakGLNo = wshGL_Trans.Range("T2").Value
    oldDesc = wshGL_Trans.Range("U2").Value
    
    For r = 2 To lastUsedRow
        If wshGL_Trans.Range("T" & r).Value <> BreakGLNo Then
            Call GL_Trans_Sub_Total(BreakGLNo, oldDesc, solde)
            BreakGLNo = wshGL_Trans.Range("T" & r).Value
            oldDesc = wshGL_Trans.Range("U" & r).Value
            solde = 0
        End If
        solde = solde + wshGL_Trans.Range("V" & r).Value - wshGL_Trans.Range("W" & r).Value
    Next r
    
    Call GL_Trans_Sub_Total(BreakGLNo, oldDesc, solde)
    
    r = wshGL_BV.Range("B2").Value + 2
    
    Call Display_TB_Totals(r, 6) 'D�bit and Cr�dit - 2024-03-05 @ 14:10
'    Call Display_TB_Totals(r, 7) 'Cr�dit
    
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

    wshGL_BV.Range("B2").Value = r - 2
  
    Application.EnableEvents = True
  
End Sub

Sub Display_TB_Totals(r As Long, c As Long) '2024-03-05 @ 14:03

    'Dt and Ct columns at the same time
    Dim sumDtRange As Range, sumCtRange As Range
    Dim sumDt As Double, sumCT As Double

    With wshGL_BV
        With .Range(.Cells(r, c), .Cells(r, c + 1)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Range(.Cells(r, c), .Cells(r, c + 1)).Borders(xlBottom)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    
        .Range(.Cells(r, c), .Cells(r, c + 1)).Font.Bold = True
        .Range(.Cells(r, c), .Cells(r, c + 1)).NumberFormat = "#,##0.00 $"
        Set sumDtRange = Range(Cells(4, c), Cells(r - 1, c))
        Set sumCtRange = Range(Cells(4, c + 1), Cells(r - 1, c + 1))
        
        .Cells(r, c).Value = Application.WorksheetFunction.Sum(sumDtRange)
        .Cells(r, c + 1).Value = Application.WorksheetFunction.Sum(sumCtRange)
        If .Cells(r, c).Value <> .Cells(r, c + 1).Value Then
            Call Erreur_Totaux_DT_CT
        End If
    End With
    
    Set sumDtRange = Nothing
    Set sumCtRange = Nothing
    
End Sub

Sub GL_Display_Trans_Selected_Account(GLAcct As String, GLDesc As String, minDate As Date, maxDate As Date) 'Display GL Trans for a specific account

    'Clear the display area & display the account number & description
    wshGL_BV.Range("M4:T99999").Clearcontents
    wshGL_BV.Range("L2").Value = "Du " & minDate & " au " & maxDate
    
    wshGL_BV.Range("L4").Font.Bold = True
    wshGL_BV.Range("L4").Value = GLAcct & " - " & GLDesc
    wshGL_BV.Range("B6").Value = GLAcct
    wshGL_BV.Range("B7").Value = GLDesc
    
    'Use the Advanced Filter Result already prepared for TB
    Dim row As Range, foundRow As Long, lastResultUsedRow As Long
    lastResultUsedRow = wshGL_Trans.Range("T99999").End(xlUp).row
    foundRow = 0
    
    'Find the first occurence of GlACct in AdvancedFilter Results on GL_Trans
    Dim foundCell As Range, searchRange As Range
    Set searchRange = wshGL_Trans.Range("T2:T" & lastResultUsedRow)
    Set foundCell = searchRange.Find(What:=GLAcct, LookIn:=xlValues, LookAt:=xlWhole)
    foundRow = foundCell.row
    
    ' Check if the target value was found
    If foundRow = 0 Then
        MsgBox "Il n'existe aucune transaction pour ce compte (p�riode choisie)."
        Exit Sub
    End If
    
    Dim rowGLDetail As Long
    rowGLDetail = 5
    wshGL_BV.Range("S4").Value = 0
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
    Do Until wshGL_Trans.Range("T" & foundRow).Value <> GLAcct
        'Traitement des transactions d�taill�es
        d = Format(wshGL_Trans.Range("Q" & foundRow).Value2, "dd-mm-yyyy")
        If d >= minDate And d <= maxDate Then
            .Range("M" & rowGLDetail).Value = wshGL_Trans.Range("Q" & foundRow).Value
            .Range("N" & rowGLDetail).Value = wshGL_Trans.Range("P" & foundRow).Value
            .Range("N" & rowGLDetail).HorizontalAlignment = xlCenter
            .Range("O" & rowGLDetail).Value = wshGL_Trans.Range("R" & foundRow).Value
            .Range("P" & rowGLDetail).Value = wshGL_Trans.Range("S" & foundRow).Value
            .Range("Q" & rowGLDetail).Value = wshGL_Trans.Range("V" & foundRow).Value
            .Range("R" & rowGLDetail).Value = wshGL_Trans.Range("W" & foundRow).Value
            .Range("S" & rowGLDetail).Value = wshGL_BV.Range("S" & rowGLDetail - 1).Value + _
                wshGL_Trans.Range("V" & foundRow).Value - wshGL_Trans.Range("W" & foundRow).Value
'            With .Range("S" & rowGLDetail).Font
'                .Name = "Aptos Narrow"
'                .Size = 11
'            End With
            .Range("T" & rowGLDetail).Value2 = wshGL_Trans.Range("X" & foundRow).Value
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

    Dim timerStart As Double: timerStart = Timer

    With wshGL_Trans
        Dim rgResult As Range, rgData As Range, rgCriteria As Range, rgCopyToRange As Range
        Set rgResult = .Range("P2").CurrentRegion
        Call ClearRangeBorders(rgResult)
        rgResult.Clearcontents
        Set rgData = .Range("A1").CurrentRegion
        .Range("L3").Value = GLNo
        .Range("M3").Value = ">=" & Format(minDate, "mm-dd-yyyy")
        .Range("N3").Value = "<=" & Format(maxDate, "mm-dd-yyyy")
        
        Set rgCriteria = .Range("L2:N3")
        Set rgCopyToRange = .Range("P1")
        
        rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgCopyToRange
        
        Dim lastResultUsedRow
        lastResultUsedRow = .Range("P99999").End(xlUp).row
        If lastResultUsedRow < 3 Then GoTo NoSort
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
            .SetRange wshGL_Trans.Range("P2:Y" & lastResultUsedRow) 'Set Range
            .Apply 'Apply Sort
         End With
    End With

NoSort:

    Call Output_Timer_Results("GL_Trans_Advanced_Filter()", timerStart)

End Sub

Sub GL_Trans_Sub_Total(GLNo As String, GLDesc As String, s As Currency)

    Dim timerStart As Double: timerStart = Timer

    Dim r As Long
    With wshGL_BV
        r = .Range("B2").Value + 1
        .Range("D" & r).HorizontalAlignment = xlCenter
        .Range("D" & r).Value = GLNo
        .Range("E" & r).Value = GLDesc
        If s > 0 Then
            .Range("F" & r).Value = s
        ElseIf s < 0 Then
            .Range("G" & r).Value = -s
        End If
        .Range("B2").Value = wshGL_BV.Range("B2").Value + 1
    End With
    
    Call Output_Timer_Results("GL_Trans_Sub_Total()", timerStart)

End Sub

Sub DetermineFromAndToDate(period As String)

    Select Case period
        Case "Mois"
            wshGL_BV.Range("B8").Value = wshAdmin.Range("MoisDe").Value
            wshGL_BV.Range("B9").Value = wshAdmin.Range("MoisA").Value
        Case "Mois dernier"
            wshGL_BV.Range("B8").Value = wshAdmin.Range("MoisPrecDe").Value
            wshGL_BV.Range("B9").Value = wshAdmin.Range("MoisPrecA").Value
        Case "Trimestre"
            wshGL_BV.Range("B8").Value = wshAdmin.Range("TrimDe").Value
            wshGL_BV.Range("B9").Value = wshAdmin.Range("TrimA").Value
        Case "Trimestre dernier"
            wshGL_BV.Range("B8").Value = wshAdmin.Range("TrimPrecDe").Value
            wshGL_BV.Range("B9").Value = wshAdmin.Range("TrimPrecA").Value
        Case "Ann�e"
            wshGL_BV.Range("B8").Value = wshAdmin.Range("AnneeDe").Value
            wshGL_BV.Range("B9").Value = wshAdmin.Range("AnneeA").Value
        Case "Ann�e derni�re"
            wshGL_BV.Range("B8").Value = wshAdmin.Range("AnneePrecDe").Value
            wshGL_BV.Range("B9").Value = wshAdmin.Range("AnneePrecA").Value
        Case "Dates Manuelles"
            wshGL_BV.Range("B8").Value = CDate(Format("01-01-2023", "dd-mm-yyyy"))
            wshGL_BV.Range("B9").Value = CDate(Format("12-31-2023", "dd-mm-yyyy"))
        Case "Toutes les dates"
            wshGL_BV.Range("B8").Value = CDate(Format(wshGL_BV.Range("B3").Value, "dd-mm-yyyy"))
            wshGL_BV.Range("B9").Value = CDate(Format(wshGL_BV.Range("B4").Value, "dd-mm-yyyy"))
    End Select
'            Debug.Print "Period is '" & period & "' so MinDate = " & wshGL_BV.Range("B8").Value & _
'                "  maxDate = " & wshGL_BV.Range("B9").Value
End Sub

Sub GL_TB_Setup_And_Print()
    
    Dim lastRow As Long, printRange As Range, shp As Shape
    lastRow = Range("D999").End(xlUp).row + 2
    If lastRow < 4 Then Exit Sub
    Set printRange = wshGL_BV.Range("D1:G" & lastRow)
    
    Dim pagesRequired As Integer
    pagesRequired = Int((lastRow - 1) / 60) + 1
    
    Set shp = ActiveSheet.Shapes("GL_BV_Print")
    shp.Visible = msoFalse
    
    Call GL_SetUp_And_Print_Document(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    Set printRange = Nothing
    Set shp = Nothing
    
End Sub

Sub GL_TB_Setup_And_Print_Trans()
    
    Dim lastRow As Long, printRange As Range, shp As Shape
    lastRow = Range("M9999").End(xlUp).row
    If lastRow < 4 Then Exit Sub
    Set printRange = wshGL_BV.Range("L1:T" & lastRow)
    
    Dim pagesRequired As Integer
    pagesRequired = Int((lastRow - 1) / 80) + 1
    
    Set shp = ActiveSheet.Shapes("GL_BV_Print_Trans")
    shp.Visible = msoFalse
    
    Call GL_SetUp_And_Print_Document(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
End Sub

Sub GL_SetUp_And_Print_Document(myPrintRange As Range, pagesTall As Integer)
    
    Dim timerStart As Double: timerStart = Timer
    
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
        .CenterHeader = "&""Aptos Narrow,Gras""&20 " & wshAdmin.Range("NomEntreprise").Value
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
 
    Call Output_Timer_Results("GL_SetUp_And_Print_Document()", timerStart)
 
End Sub

Public Sub ClearDynamicShape()
    'Hide the shape if it's visible
    If dynamicShape.Visible Then
        dynamicShape.Visible = False
    End If
    
    'Set dynamicShape to Nothing to release memory
    Set dynamicShape = Nothing
    
End Sub

Sub Back_To_GL_Menu()
    
    wshMenuCOMPTA.Activate
    wshMenuCOMPTA.Range("A1").Select
    
End Sub


