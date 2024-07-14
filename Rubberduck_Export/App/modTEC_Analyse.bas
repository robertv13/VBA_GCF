Attribute VB_Name = "modTEC_Analyse"
Option Explicit

Sub TEC_Sort_Group_And_Subtotal()

    Dim lastUsedRow As Long
    Dim firstEmptyCol As Long
    
    'Set the source worksheet, lastUsedRow and lastUsedCol
    Dim wsSource As Worksheet: Set wsSource = wshTEC_Local
    'Find the last row with data in the source worksheet
    lastUsedRow = wsSource.Cells(wsSource.rows.count, "A").End(xlUp).row
    'Find the first empty column from the left in the source worksheet
    firstEmptyCol = 1
    Do Until IsEmpty(wsSource.Cells(2, firstEmptyCol))
        firstEmptyCol = firstEmptyCol + 1
    Loop
    Dim lastUsedCol As Long
    lastUsedCol = firstEmptyCol - 1
    
    'Set the current worksheet as the result
    Dim wsDest As Worksheet: Set wsDest = wshTEC_Analyse
    Dim destLastUsedRow As Long
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row
    'Remove existing subtotals in the destination worksheet
    wsDest.Cells.RemoveSubtotal
    wsDest.Range("A2:H" & destLastUsedRow).ClearContents
    
    Dim i As Long, r As Long
    r = 1
    Application.EnableEvents = False
    For i = 3 To lastUsedRow
        'Conditions for exclusion (adjust as needed)
        If wsSource.Cells(i, 14).value <> "VRAI" And _
            wsSource.Cells(i, 12).value <> "VRAI" And _
            wsSource.Cells(i, 10).value = "VRAI" Then
                r = r + 1
                wsDest.Cells(r, 1).value = wsSource.Cells(i, ftecProf_ID).value
                wsDest.Cells(r, 2).value = wsSource.Cells(i, ftecClientNom).value
                wsDest.Cells(r, 4).value = wsSource.Cells(i, ftecDate).value
                wsDest.Cells(r, 5).value = wsSource.Cells(i, ftecProf).value
                wsDest.Cells(r, 6).value = wsSource.Cells(i, ftecDescription).value
                wsDest.Cells(r, 7).value = wsSource.Cells(i, ftecHeures).value
                wsDest.Cells(r, 8).value = wsSource.Cells(i, ftecCommentaireNote).value
        End If
    Next i
    Application.EnableEvents = False
   
    'Find the last row in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row

    'Sort by Client_ID (column E) and Date (column D) in the destination worksheet
    wsDest.Sort.SortFields.clear
    wsDest.Sort.SortFields.add key:=wsDest.Range("B2:B" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("D2:D" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("A2:A" & destLastUsedRow), Order:=xlAscending
    
    With wsDest.Sort
        .SetRange wsDest.Range("A1:P" & destLastUsedRow)
        .header = xlYes
        .MatchCase = False
        .orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Add subtotals for hours (column F) at each change in ClientNom_ID (column B) in the destination worksheet
    wsDest.Range("A1:P" & destLastUsedRow).Subtotal GroupBy:=2, Function:=xlSum, _
        TotalList:=Array(7), Replace:=True, PageBreaks:=False, SummaryBelowData:=False

    'Group the data to show subtotals in the destination worksheet
    wsDest.Outline.ShowLevels RowLevels:=2

End Sub

Sub Build_Hours_Summary(client As String, r As Long)

    If r < 3 Then Exit Sub
    
    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Analyse.Cells(wshTEC_Analyse.rows.count, "A").End(xlUp).row
    wshTEC_Analyse.Range("J:M").clear
    
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
    Dim t As Double
    Dim i As Long, saveR As Long
    saveR = r
    i = r + 1
    Do Until Cells(i, 4) = ""
        If Cells(i, 5).value <> "" Then
            t = t + Cells(i, 7).value
            If dictHours.Exists(Cells(i, 5).value) Then
                dictHours(Cells(i, 5).value) = dictHours(Cells(i, 5).value) + Cells(i, 7).value
            Else
                dictHours.add Cells(i, 5).value, Cells(i, 7).value
            End If
        End If
        i = i + 1
    Loop

    Dim Prof As Variant
    For Each Prof In Fn_Sort_Dictionary_By_Value(dictHours, True) ' Sort dictionary by hours in descending order
        Cells(r, 10).value = Prof
        Cells(r, 11).HorizontalAlignment = xlRight
        Cells(r, 11).value = Format(dictHours(Prof), "#,##0.00")
        Cells(r, 13).formula = "=cells(r,11)*cells(r,12)"
        r = r + 1
    Next Prof
    
    With Cells(r, 11)
        .HorizontalAlignment = xlRight
        .value = Format(t, "#,##0.00")
        .Font.Bold = True
    End With
    
    With Range("J" & saveR & ":M" & r).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    With Range("K" & r & ", M" & r)
'        .Borders(xlDiagonalDown).LineStyle = xlNone
'        .Borders(xlDiagonalUp).LineStyle = xlNone
'        .Borders(xlEdgeLeft).LineStyle = xlNone
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
'    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
'    Selection.Borders(xlEdgeRight).LineStyle = xlNone
'    Selection.Borders(xlInsideVertical).LineStyle = xlNone
'    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    'Clean up - 2024-07-11 @ 15:20
    Set dictHours = Nothing
    
End Sub
    
