Attribute VB_Name = "modCC_Analyse"
Option Explicit

Dim previousCellAddress As Variant

Sub CC_Sort_Group_And_Subtotal()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modCC_Analyse:CC_Sort_Group_And_Subtotal()")
    
    Application.ScreenUpdating = False
    
    'Calculate the center of the used range
    Dim centerX As Double, centerY As Double
    centerX = 482
    centerY = 58

    'Set the dimensions of the progress bar
    Dim barWidth As Double, barHeight As Double
    barWidth = 300
    barHeight = 25  'Height of the progress bar

    'Create the background shape of the progress bar positioned in the center of the visible range
    Dim progressBarBg As Shape
    Set progressBarBg = ActiveSheet.Shapes.AddShape(msoShapeRectangle, centerX - barWidth / 3, centerY - barHeight / 2, barWidth, barHeight)
    progressBarBg.Fill.ForeColor.RGB = RGB(255, 255, 255)  ' White background
    progressBarBg.Line.Visible = msoTrue  'Show the border of the progress bar
    progressBarBg.TextFrame.HorizontalAlignment = xlHAlignCenter
    progressBarBg.TextFrame.VerticalAlignment = xlVAlignCenter
    progressBarBg.TextFrame.Characters.Font.Size = 14
    progressBarBg.TextFrame.Characters.Font.Color = RGB(0, 0, 0) 'Black font
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à 0 %"
    
    'Create the fill shape of the progress bar
    Dim progressBarFill As Shape
    Set progressBarFill = ActiveSheet.Shapes.AddShape(msoShapeRectangle, centerX - barWidth / 3, centerY - barHeight / 2, 0, barHeight)
    progressBarFill.Fill.ForeColor.RGB = RGB(0, 255, 0)  ' Green fill color
    progressBarFill.Fill.Transparency = 0.6  'Set transparency to 60%
    progressBarFill.Line.Visible = msoFalse  'Hide the border of the fill
    
    'Update the progress bar fill
    progressBarFill.width = 0.1 * barWidth   '10 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.1, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    Dim lastUsedRow As Long, firstEmptyCol As Long
    
    'Set the source worksheet, lastUsedRow and lastUsedCol
    Dim wsSource As Worksheet: Set wsSource = wshFAC_Comptes_Clients
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
    Dim wsDest As Worksheet: Set wsDest = wshCC_Analyse
    wsDest.Range("J3").value = #7/24/2025#
    'Remove existing subtotals in the destination worksheet
    On Error Resume Next
    wsDest.Cells.RemoveSubtotal
    On Error GoTo 0
    
    Dim destLastUsedRow As Long
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row
    wsDest.Range("A6:K" & destLastUsedRow).clear
    
    'Update the progress bar fill
    progressBarFill.width = 0.25 * barWidth  '25 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.25, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    Dim i As Long, r As Long
    r = 6
    Dim b As Long
    
    Application.EnableEvents = False
    Dim ageJours As Long
    For i = 3 To lastUsedRow
        'Conditions for exclusion (adjust as needed)
        If wsSource.Cells(i, 9).value <> 0 Then
            If wsSource.Cells(i, 2).value <= wsDest.Range("J3").value Then
'                wsDest.Cells(r, 1).value = wsSource.Cells(i, ftecTEC_ID).value
                wsDest.Cells(r, 1).value = wsSource.Cells(i, 3).value
                wsDest.Cells(r, 2).value = wsSource.Cells(i, 1).value
                wsDest.Cells(r, 3).value = wsSource.Cells(i, 2).value
                wsDest.Cells(r, 4).value = wsSource.Cells(i, 6).value
                ageJours = Round(Now() - wsSource.Cells(i, 6).value, 0)
                wsDest.Cells(r, 5).formula = ageJours
                wsDest.Cells(r, 6).value = wsSource.Cells(i, 9).value
                b = Fn_Get_Bucket_For_Aging(ageJours, _
                                            wsDest.Range("M3").value, _
                                            wsDest.Range("N3").value, _
                                            wsDest.Range("O3").value, _
                                            wsDest.Range("P3").value)
                If b < 0 Or b > 4 Then Stop
                wsDest.Cells(r, 7 + b).value = wsSource.Cells(i, 9).value
                r = r + 1
            End If
        End If
    Next i
    Application.EnableEvents = False
    
    'Update the progress bar fill
    progressBarFill.width = 0.5 * barWidth   '50 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.5, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
   
    'Find the last row in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row

    'Sort by Client_ID (column E) and Date (column D) in the destination worksheet
    wsDest.Sort.SortFields.clear
    wsDest.Sort.SortFields.add key:=wsDest.Range("A5:A" & destLastUsedRow), Order:=xlAscending
    wsDest.Sort.SortFields.add key:=wsDest.Range("B5:B" & destLastUsedRow), Order:=xlAscending
    
    With wsDest.Sort
        .SetRange wsDest.Range("A5:K" & destLastUsedRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Update the progress bar fill
    progressBarFill.width = 0.65 * barWidth  '65 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.65, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    'Add subtotals for amount columns at each change in ClientNom_ID (column C) in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row
    Application.DisplayAlerts = False
    wsDest.Range("A6:K" & destLastUsedRow).Subtotal GroupBy:=1, Function:=xlSum, _
        TotalList:=Array(6, 7, 8, 9, 10, 11), Replace:=True, PageBreaks:=False, SummaryBelowData:=False
    Application.DisplayAlerts = True
'    wsDest.Range("A:B").EntireColumn.Hidden = True

    'Group the data to show subtotals in the destination worksheet
    destLastUsedRow = wsDest.Cells(wsDest.rows.count, "A").End(xlUp).row
    wsDest.Range("F6:K" & destLastUsedRow).NumberFormat = "#,##0.00 $"
    wsDest.Outline.ShowLevels RowLevels:=2
    
    'Update the progress bar fill
    progressBarFill.width = 0.75 * barWidth  '75 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.75, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    'Change the format of the top row (Total General)
    With wsDest.Range("A6:K6")
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
    
    'Change the format of all Client's Total rows
    For r = 7 To destLastUsedRow
        If InStr(wsDest.Range("A" & r).value, "Total ") = 1 Then
            With wsDest.Range("A" & r & ":K" & r).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With wsDest.Range("A" & r & ":K" & r).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
    Next r
    
    'Update the progress bar fill
    progressBarFill.width = 0.85 * barWidth  '85 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.85, "0%")
    
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
'    'Set conditional formats for total hours (Client's total)
'    Dim rngTotals As Range: Set rngTotals = wsDest.Range("C7:C" & destLastUsedRow)
'    Call Apply_Conditional_Formatting_Alternate_On_Column_H(rngTotals, destLastUsedRow)
'
'    'Bring in all the invoice requests
'    Call Bring_In_Existing_Invoice_Requests(destLastUsedRow)
'
'    'Clean up the summary aera of the worksheet
'    Call Clean_Up_Summary_Area(wsDest)
'
    'Update the progress bar fill
    progressBarFill.width = 0.95 * barWidth   '95 %
    'Update the caption on the background shape
    progressBarBg.TextFrame.Characters.text = "Préparation complétée à " & Format$(0.95, "0%")
    
    'Introduce a small delay to ensure the worksheet is fully updated
    Application.Wait (Now + TimeValue("0:00:01")) '2024-07-23 @ 16:13 - Slowdown the application
        
    'Temporarily enable screen updating to show the progress bar
    Application.ScreenUpdating = True
    DoEvents  'Allow Excel to process other events
    Application.ScreenUpdating = False
    
    progressBarBg.delete
    progressBarFill.delete
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

'    Application.StatusBar = ""

    Call Output_Timer_Results("modCC_Analyse:CC_Sort_Group_And_Subtotal()", timerStart)

End Sub

Sub CC_Analyse_Back_To_FAC_Menu()

    wshCC_Analyse.Visible = xlSheetHidden
    
    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select

End Sub

