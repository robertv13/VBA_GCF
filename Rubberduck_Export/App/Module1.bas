Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$8:$8"
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$B$9:$I$51"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = _
        "&""-,Gras""&12&K0070C0Liste âgée des comptes clients" & Chr(10) & "&11Par ordre alphabétique - 1 ligne par Facture"
        .RightHeader = ""
        .LeftFooter = "&9&D - &T"
        .CenterFooter = "&9&KFF0000&A"
        .RightFooter = "&""Segoe UI,Normal""&9Page &P of &N"
        .LeftMargin = Application.InchesToPoints(0.15748031496063)
        .RightMargin = Application.InchesToPoints(0.15748031496063)
        .TopMargin = Application.InchesToPoints(0.748031496062992)
        .BottomMargin = Application.InchesToPoints(0.551181102362205)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = -3
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 10
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.text = ""
        .EvenPage.CenterHeader.text = ""
        .EvenPage.RightHeader.text = ""
        .EvenPage.LeftFooter.text = ""
        .EvenPage.CenterFooter.text = ""
        .EvenPage.RightFooter.text = ""
        .FirstPage.LeftHeader.text = ""
        .FirstPage.CenterHeader.text = ""
        .FirstPage.RightHeader.text = ""
        .FirstPage.LeftFooter.text = ""
        .FirstPage.CenterFooter.text = ""
        .FirstPage.RightFooter.text = ""
    End With
    Application.PrintCommunication = True
End Sub
