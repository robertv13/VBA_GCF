Attribute VB_Name = "Module1"
Option Explicit

Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$B$2:$G$3"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = _
        "&""-,Gras""&14Recherche&""-,Normal""&11" & Chr(10) & "&""-,Gras""&12Chaine"
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.708661417322835)
        .RightMargin = Application.InchesToPoints(0.708661417322835)
        .TopMargin = Application.InchesToPoints(0.748031496062992)
        .BottomMargin = Application.InchesToPoints(0.748031496062992)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
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
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$B$2:$G$12"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""-,Gras""&14&K0070C0Recherche"
        .RightHeader = ""
        .LeftFooter = "&11&D - &T"
        .CenterFooter = "&11&KFF0000&A"
        .RightFooter = "&11Page &P de &N"
        .LeftMargin = Application.InchesToPoints(0.15748031496063)
        .RightMargin = Application.InchesToPoints(0.15748031496063)
        .TopMargin = Application.InchesToPoints(0.551181102362205)
        .BottomMargin = Application.InchesToPoints(0.551181102362205)
        .HeaderMargin = Application.InchesToPoints(0.15748031496063)
        .FooterMargin = Application.InchesToPoints(0.15748031496063)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 4
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
