Attribute VB_Name = "modDebug"
Option Explicit

Sub getShapeNames()

    Dim sh As Shape
    Dim x As Double
    x = 1
    ActiveSheet.Select

    For Each sh In ActiveSheet.Shapes
        ActiveSheet.Range("C" + Trim(Str(x + 1))).Value = x & _
            " - " & sh.Name
        x = x + 1
    Next

End Sub

Sub ARoseByAnyOtherName()
    ActiveSheet.Shapes(10).Select
    Selection.Name = "imgIconeSchedule"
'    ActiveSheet.Shapes(15).Select
'    Selection.Name = "imgIconeGraphs"
'    ActiveSheet.Shapes(16).Select
'    Selection.Name = "imgIconeReports"

End Sub

