Attribute VB_Name = "modTEC_Analyse"
Option Explicit

Sub TEC_Sort_Group_And_Subtotal()

    Application.ScreenUpdating = False
    
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
    'Remove existing subtotals in the destination worksheet
    wsDest.Cells.RemoveSubtotal
    
    Dim destLastUsedRow As Long
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "B").End(xlUp).row
    wsDest.Range("A6:H" & destLastUsedRow).ClearContents
    
    Dim i As Long, r As Long
    r = 6
    Application.EnableEvents = False
    For i = 3 To lastUsedRow
        'Conditions for exclusion (adjust as needed)
        If wsSource.Cells(i, 14).value <> "VRAI" And _
            wsSource.Cells(i, 12).value <> "VRAI" And _
            wsSource.Cells(i, 10).value = "VRAI" Then
                wsDest.Cells(r, 1).value = wsSource.Cells(i, ftecTEC_ID).value
                wsDest.Cells(r, 2).value = wsSource.Cells(i, ftecProf_ID).value
                wsDest.Cells(r, 3).value = wsSource.Cells(i, ftecClientNom).value
                wsDest.Cells(r, 5).value = wsSource.Cells(i, ftecDate).value
                wsDest.Cells(r, 6).value = wsSource.Cells(i, ftecProf).value
                wsDest.Cells(r, 7).value = wsSource.Cells(i, ftecDescription).value
                wsDest.Cells(r, 8).value = wsSource.Cells(i, ftecHeures).value
                wsDest.Cells(r, 9).value = wsSource.Cells(i, ftecCommentaireNote).value
                r = r + 1
        End If
    Next i
    Application.EnableEvents = False
   
    'Find the last row in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row

    'Sort by Client_ID (column E) and Date (column D) in the destination worksheet
    wsDest.Sort.SortFields.clear
    wsDest.Sort.SortFields.add key:=wsDest.Range("C6:C" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("E6:E" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("B6:B" & destLastUsedRow), Order:=xlAscending
    
    With wsDest.Sort
        .SetRange wsDest.Range("A5:H" & destLastUsedRow)
        .header = xlYes
        .MatchCase = False
        .orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Add subtotals for hours (column G) at each change in ClientNom_ID (column B) in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row
    Application.DisplayAlerts = False
    wsDest.Range("A6:I" & destLastUsedRow).Subtotal GroupBy:=3, Function:=xlSum, _
        TotalList:=Array(8), Replace:=True, PageBreaks:=False, SummaryBelowData:=False
    Application.DisplayAlerts = True
    wsDest.Range("A:B").EntireColumn.Hidden = True

    'Group the data to show subtotals in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row
    wsDest.Outline.ShowLevels RowLevels:=2
    
    'Add a formula to sum the billed amounts
    wshTEC_Analyse.Range("D6").formula = "=SUM(D7:D" & destLastUsedRow & ")"
    
    'Change the format of 'Total general' row
    With wshTEC_Analyse.Range("D6")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .Color = -16776961
            .TintAndShade = 0
            .Bold = True
            .Size = 12
        End With
    End With
    With wshTEC_Analyse.Range("H6")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .Color = -16776961
            .TintAndShade = 0
            .Bold = True
            .Size = 12
        End With
    End With
    
    'Change the format of Group Totals rows
    For r = 7 To destLastUsedRow
        If wsDest.Range("A" & r).value = "" Then
            With wsDest.Range("C" & r).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With wsDest.Range("C" & r).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            With wsDest.Range("H" & r).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With wsDest.Range("H" & r).Font
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
    Next r
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub Build_Hours_Summary(client As String, r As Long)

    If r < 7 Then Exit Sub
    
    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Analyse.Cells(wshTEC_Analyse.rows.count, "A").End(xlUp).row
    wshTEC_Analyse.Range("K:Q").clear
    Call Delete_CheckBox
    
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")
    Dim i As Long, saveR As Long
    r = r + 1 'Summary starts on the next line (first line of expanded lines)
    saveR = r
    i = r
    Do Until Cells(i, 5) = ""
        If Cells(i, 6).value <> "" Then
'            t = t + Cells(i, 7).value
            If dictHours.Exists(Cells(i, 6).value) Then
                dictHours(Cells(i, 6).value) = dictHours(Cells(i, 6).value) + Cells(i, 8).value
            Else
                dictHours.add Cells(i, 6).value, Cells(i, 8).value
            End If
        End If
        i = i + 1
    Loop

    Dim Prof As Variant
    Dim ProfID As Integer
    wshTEC_Analyse.Range("Q" & r).value = 0 'Reset the total WIP value
    For Each Prof In Fn_Sort_Dictionary_By_Value(dictHours, True) ' Sort dictionary by hours in descending order
        Cells(r, 11).value = Prof
        Dim strProf As String
        strProf = Prof
        ProfID = Fn_GetID_From_Initials(strProf)
        Cells(r, 12).HorizontalAlignment = xlRight
        Cells(r, 12).NumberFormat = "#,##0.00"
        Cells(r, 12).value = dictHours(Prof)
        Dim tauxHoraire As Currency
        tauxHoraire = Fn_Get_Hourly_Rate(ProfID, "2024-07-15")
        Cells(r, 13).value = tauxHoraire
        Cells(r, 14).NumberFormat = "#,##0.00$"
        Cells(r, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
        Cells(r, 14).HorizontalAlignment = xlRight
        r = r + 1
    Next Prof
    
    'Sort the summary by rate (descending value) if required
    If r - 1 > saveR Then
        Dim rngSort As Range
        Set rngSort = wshTEC_Analyse.Range(wshTEC_Analyse.Cells(saveR, 11), wshTEC_Analyse.Cells(r - 1, 14))
        rngSort.Sort Key1:=wshTEC_Analyse.Cells(saveR, 13), Order1:=xlDescending, header:=xlNo
    End If
    
    'Add totals to the summary
    Dim rTotal As Long
    rTotal = r
    With Cells(rTotal, 12)
        .HorizontalAlignment = xlRight
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
'        .value = Format(t, "#,##0.00")
        .Font.Bold = True
    End With
    
    With Cells(r, 14)
        .HorizontalAlignment = xlRight
'        .value = Format(tdollars, "#,##0.00$")
        .FormulaR1C1 = "=SUM(R" & saveR & "C:R[-1]C)"
        .Font.Bold = True
    End With
    
    With Range("K" & saveR & ":N" & r).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    With Range("L" & r & ", N" & r)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With

    'Save the TOTAL WIP value
    With wshTEC_Analyse.Range("P" & saveR)
        .value = "Valeur TEC:"
        .Font.Italic = True
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With wshTEC_Analyse.Range("Q" & saveR)
        .NumberFormat = "#,##0.00$"
        .value = wshTEC_Analyse.Range("N" & r).value
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    'Create a visual clue is amounts are different
    With wshTEC_Analyse.Range("Q" & r)
        Dim formula As String
        formula = "=IF(N" & r & "<>Q" & saveR & ", N" & r & "-Q" & saveR & ",""""" & ")"
        .formula = formula
        .NumberFormat = "#,##0.00$"
    End With
    
    Call Add_And_Modify_Checkbox(saveR, r)
    
    'Clean up - 2024-07-11 @ 15:20
    Set dictHours = Nothing
    Set rngSort = Nothing
    
End Sub
    
Sub Add_And_Modify_Checkbox(startRow As Long, lastRow As Long)
    
    'Set your worksheet (adjust this to match your worksheet name)
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Define the range for the summary
    Dim summaryRange As Range
    Set summaryRange = ws.Range(ws.Cells(startRow, 11), ws.Cells(lastRow, 14)) 'Columns K to N
    
    'Add an ActiveX checkbox next to the summary in column O
    Dim checkBox As OLEObject
    With ws
        Set checkBox = .OLEObjects.add(ClassType:="Forms.CheckBox.1", _
                    Left:=.Cells(lastRow, 15).Left + 5, _
                    Top:=.Cells(lastRow, 15).Top, width:=80, Height:=16)
        
        'Modify checkbox properties
        With checkBox.Object
            .Caption = "On facture"
            .Font.Size = 11  ' Set font size
            .Font.Bold = True  ' Set font bold
            .ForeColor = RGB(0, 0, 255)  ' Set font color (Blue)
            .BackColor = RGB(200, 255, 200)  ' Set background color (Light Green)
        End With
    End With
    
End Sub

Sub Delete_CheckBox()

    'Set your worksheet (adjust this to match your worksheet name)
    Dim ws As Worksheet: Set ws = wshTEC_Analyse
    
    'Check if CheckBox1 exists and then delete it
    Dim checkBox As OLEObject
    Dim i As Integer
    For i = 1 To 5
        On Error Resume Next
        Set checkBox = ws.OLEObjects("CheckBox" & i)
        If Not checkBox Is Nothing Then
            checkBox.delete
        End If
        On Error GoTo 0
    Next i
End Sub

'Sub ExpandAllGroups()
'
'    Dim ws As Worksheet
'    Dim r As Range
'
'    ' Set the worksheet you want to work on
'    Set ws = ThisWorkbook.Sheets("Sheet1")
'
'    ' Loop through each row in the used range of the worksheet
'    For Each r In ws.usedRange.rows
'        ' If the row is part of a group and can be expanded
'        If r.OutlineLevel > 1 Then
'            r.ShowDetail = True
'        End If
'    Next r
'End Sub
'
'Sub Groups_Collapse_All()
'
'    'Set the worksheet you want to work on
'    Dim ws As Worksheet: Set ws = wshTEC_Analyse
'
'    'Loop through each row in the used range of the worksheet
'    Dim r As Range
'    Debug.Print ws.usedRange.rows.count
'    For Each r In ws.usedRange.rows
'        'If the row is part of a group and can be collapsed
'        If r.OutlineLevel = 3 Then
'            Debug.Print r.Address
'            r.ShowDetail = False
'        End If
'    Next r
'End Sub
'
'Sub Fix_Format(rng As Range, value As Variant)
'
'    Debug.Print rng.Address & " - " & value
'
'End Sub
