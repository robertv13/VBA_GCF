Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    columns("K:P").Select
    Selection.ColumnWidth = 15
    Range("P2").Select
    With ActiveSheet.PivotTables("Tableau croisé dynamique1").PivotFields( _
        "Hres/TEC")
        .NumberFormat = "# ##0,00"
    End With
    Range("N2").Select
    ActiveWindow.SmallScroll Down:=190
    Range("N193").Select
    With ActiveSheet.PivotTables("Tableau croisé dynamique1").PivotFields( _
        "Hres/Nfact")
        .NumberFormat = "# ##0,00"
    End With
    ActiveWindow.SmallScroll Down:=-80
    Range("G136").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G137").Select
    ActiveWindow.SmallScroll Down:=-5
    Range("L123").Select
    Calculate
    Range("G136").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("G137").Select
    ActiveWindow.SmallScroll Down:=-105
    Range("N7").Select
End Sub
