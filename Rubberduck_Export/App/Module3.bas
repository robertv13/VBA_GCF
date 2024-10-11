Attribute VB_Name = "Module3"
Option Explicit

Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("X3").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "TEC_TDB_Data!R1C23:R107C30", Version:=8).CreatePivotTable TableDestination _
        :="PivotSheet!R3C1", TableName:="Tableau croisé dynamique1", DefaultVersion _
        :=8
    Sheets("PivotSheet").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("Tableau croisé dynamique1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("Tableau croisé dynamique1").pivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("Tableau croisé dynamique1").RepeatAllLabels _
        xlRepeatLabels
    ActiveSheet.PivotTables("Tableau croisé dynamique1").AddDataField ActiveSheet. _
        PivotTables("Tableau croisé dynamique1").PivotFields("ProfID"), _
        "Somme de ProfID", xlSum
    With ActiveSheet.PivotTables("Tableau croisé dynamique1").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tableau croisé dynamique1").AddDataField ActiveSheet. _
        PivotTables("Tableau croisé dynamique1").PivotFields("H_N_D"), "Somme de H_N_D" _
        , xlSum
    ActiveSheet.PivotTables("Tableau croisé dynamique1").AddDataField ActiveSheet. _
        PivotTables("Tableau croisé dynamique1").PivotFields("H_Facturables"), _
        "Somme de H_Facturables", xlSum
    ActiveSheet.PivotTables("Tableau croisé dynamique1").AddDataField ActiveSheet. _
        PivotTables("Tableau croisé dynamique1").PivotFields("H_NonFact"), _
        "Somme de H_NonFact", xlSum
    With ActiveSheet.PivotTables("Tableau croisé dynamique1").PivotFields( _
        "Somme de ProfID")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tableau croisé dynamique1").PivotFields("ProfID"). _
        Orientation = xlHidden
    With ActiveSheet.PivotTables("Tableau croisé dynamique1").PivotFields("Prof")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tableau croisé dynamique1").PivotFields("Prof")
        .Orientation = xlRowField
        .Position = 1
    End With
    Range("B4:D21").Select
    With ActiveSheet.PivotTables("Tableau croisé dynamique1").PivotFields( _
        "Somme de H_N_D")
        .NumberFormat = "# ##0,00"
    End With
    Range("B3").Select
    ActiveSheet.PivotTables("Tableau croisé dynamique1").DataPivotField.PivotItems( _
        "Somme de H_N_D").Caption = "Hres/Nettes"
    Range("C3").Select
    ActiveSheet.PivotTables("Tableau croisé dynamique1").DataPivotField.PivotItems( _
        "Somme de H_Facturables").Caption = "Hres/FACT"
    Range("D3").Select
    ActiveSheet.PivotTables("Tableau croisé dynamique1").DataPivotField.PivotItems( _
        "Somme de H_NonFact").Caption = "Hres/NonFact"
    columns("B:D").Select
    Selection.ColumnWidth = 12
    Range("B3:D3").Select
    Selection.Font.size = 10
    Selection.Font.size = 9
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A3").Select
    ActiveSheet.PivotTables("Tableau croisé dynamique1").CompactLayoutRowHeader = _
        "Professionnel"
    Range("A4").Select
End Sub
