Attribute VB_Name = "modCopyCutPaste"
Option Explicit

Sub CopyPasteOneCell()
    'Copy and Paste Single Cell
    Range("A1").Copy Range("B1")
    Application.CutCopyMode = False
End Sub

Sub CutPasteOneCell()
    'Cut and Paste Single Cell
    Range("A1").Cut Range("B1")
    Application.CutCopyMode = False
End Sub

Sub CopyPasteOneRow()
    'Copy and Paste 1 Row, row '5' to row '10'
    Range("5:5").Copy Range("10:10")
    Application.CutCopyMode = False
End Sub

Sub CutPasteOneRow()
    'Cut and Paste 1 Row, row '5' to row '10'
    Range("5:5").Cut Range("20:20")
    Application.CutCopyMode = False
End Sub

Sub CopyPasteCellOtherWorksheet()
    'Copy and Paste to another worksheet, within the same Workbook
    Worksheets("feuil1").Range("A1").Copy Worksheets("feuil2").Range("B1")
    Application.CutCopyMode = False
End Sub

Sub CutPasteCellOtherWorksheet()
    'Cut and Paste to another worksheet, within the same Workbook
    Worksheets("feuil1").Range("A1").Cut Worksheets("feuil2").Range("B1")
    Application.CutCopyMode = False
End Sub

Sub CopyPasteCellOtherWorkbook()                 'BUG !!!
    'Copy and Paste to another Workbook
    Workbooks("VBA_Code_Example.xlsb").Worksheets("Feuil1").Range("A1").Copy _
        Workbooks("book2.xlsm").Worksheets("Feuil1").Range("B1")
    Application.CutCopyMode = False
End Sub

Sub CutyPasteCellOtherWorkbook()                 'BUG !!!
    'Copy and Paste to another Workbook
    Workbooks("VBA_Code_Example.xlsb").Worksheets("Feuil1").Range("A1").Cut _
        Workbooks("book2.xlsm").Worksheets("Feuil1").Range("B1")
    Application.CutCopyMode = False
End Sub

Sub CopySelection()                              'BUG !!!
    'Copy and Paste the selection to a Defined Range 'B1'
    Selection.Copy Range("b1")
    'Copy and Paste the selection to a Range using Offset, 2 rows down, 1 column right
    Selection.Copy
    Selection.Offset(2, 1).Paste
End Sub

Sub PasteSpecial()
    'Perform one Paste Special Operation a the time for cell 'A1'
    Range("A1").Copy
    'Paste Formats
    Range("B1").PasteSpecial Paste:=xlPasteFormats
    'Paste Column Widths
    Range("B1").PasteSpecial Paste:=xlPasteColumnWidths
    'Paste Formulas or contents
    Range("B1").PasteSpecial Paste:=xlPasteFormulas
    'Perform Multiple Paste Special Operations at Once:
    Range("A1").Copy
    'Paste Formats, ColumnWidths and Contents
    Range("B3").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
End Sub

Sub ValuePaste()
    'Value Paste Cells
    Range("B1").Value = Range("A1").Value
    Range("B1:B3").Value = Range("A1:A3").Value
    'Set Values Between Worksheets
    Worksheets("Feuil2").Range("A1").Value = Worksheets("Feuil1").Range("A1").Value
    'Set Values Between Workbooks 'BUG !!!
    Workbooks("C:\VBA\Reference\book2.xlsm").Worksheets("Feuil1").Range("A1").Value = _
                                                                                    Workbooks("VBA_Code_Examples.xlsm").Worksheets("Feuil1").Range("A1").Value
    Application.CutCopyMode = False
End Sub

Sub FormatPainter()
    'Copy format from cell 'A1' and paste it to cell 'B1'
    Range("A1").Copy
    Range("B1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub


