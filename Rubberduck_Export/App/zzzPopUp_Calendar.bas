Attribute VB_Name = "zzzPopUp_Calendar"
Dim SelCell As Range
Dim DayName As String

Sub ShowSettings()
    'Show or Hide Calendar Settings Panel
    If ActiveSheet.Shapes.Range(Array("Settings")).Visible = True Then
        ActiveSheet.Shapes.Range(Array("Settings", "CalCol1", "CalCol2", "CalCol3", "CalCol4", "CalCol5", "CalCol6", "CalCol7", "CalCol8", "CalCol9")).Visible = False
    Else:
        ActiveSheet.Shapes.Range(Array("Settings", "CalCol1", "CalCol2", "CalCol3", "CalCol4", "CalCol5", "CalCol6", "CalCol7", "CalCol8", "CalCol9")).Visible = True
    End If
End Sub

Sub ReplaceCalendar()                            'Shape Deleted
    CalPopUp.Shapes("Calendar").Copy             'Copy From Developers sheet
    ActiveCell.Select                            'Select the active Cell
    ActiveSheet.Paste                            'Paste in Calendar
End Sub

Sub CalendarHide()
    Dim DayNum As Long
    'Hide Calendar, Reset Day Colors
    On Error GoTo NoCal
    ActiveSheet.Shapes("Calendar").Visible = msoFalse
    Sheets("CalPopUp").Range("A7").value = ""
    If Sheets("CalPopUp").Range("A20").value <> Empty Then
        For DayNum = 1 To 42
            DayName = DayNum & "Day"
            With ActiveSheet.Shapes(DayName)
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextFrame2.TextRange.Font.Bold = msoFalse
            End With
        Next DayNum
    End If
    Exit Sub
NoCal:                                           'If calendar has been removed by accident, paste in backup calendar from  CalPopUp Sheet
End Sub

Sub CalendarShow()
    With ActiveSheet
        Set SelCell = Selection
        'Check if active cell is a valid date
        If IsDate(SelCell.value) = True Then
            Sheets("CalPopUp").Range("A1").value = SelCell.value
        Else:                                    'If No Date or incorrect Date user current date
            Sheets("CalPopUp").Range("A1").value = "=Today()"
        End If
        'Clear all shapes to white (if calendar is visible)
        If ActiveSheet.Shapes("Calendar").Visible = True Then
            For DayNum = 1 To 42
                DayName = DayNum & "Day"
                With ActiveSheet.Shapes(DayName)
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .TextFrame2.TextRange.Font.Bold = msoFalse
                End With
            Next DayNum
        End If
    
        Sheets("CalPopUp").Range("A3").value = Month(Sheets("CalPopUp").Range("A1").value) 'Set Month
        Sheets("CalPopUp").Range("A2").value = Year(Sheets("CalPopUp").Range("A1").value) 'Set Year
        DayName = Sheets("CalPopUp").Range("A20").value & "Day"
        ' UnGroupCal
        If InStr(.Shapes("Calendar").GroupItems("NextTri").OnAction, "!") <> 0 Or InStr(.Shapes("1Day").DrawingObject.formula, "]") <> 0 Then 'Run Workbook Link Remover and Cell Link Replacement
            MacroLinkRemover
            CalFormulaReplacement
        End If
        'GroupCal
        On Error GoTo NoCal
        With ActiveSheet.Shapes(DayName)
            .Fill.ForeColor.RGB = RGB(252, 213, 180)
            .TextFrame2.TextRange.Font.Bold = msoTrue
        End With
        On Error GoTo NoCal
        .Shapes("Calendar").Visible = msoCTrue
        .Shapes.Range(Array("Settings", "CalCol1", "CalCol2", "CalCol3", "CalCol4", "CalCol5", "CalCol6", "CalCol7", "CalCol8", "CalCol9")).Visible = False '
        .Shapes("Calendar").Left = SelCell.Left
        .Shapes("Calendar").Placement = xlMove
        .Shapes("Calendar").Top = SelCell.Offset(1, 0).Top
        If Sheets("CalPopUp").Range("A6").value > 0 Then
            .Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = True
        Else:
            .Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = False
        End If
        Sheets("CalPopUp").Range("A7").value = SelCell.Address
        ActiveCell.Select
    End With
    Exit Sub
NoCal:
    MsgBox "The Pop-up Calendar does not exist on this worksheet. Please copy the calendar over from another sheet and paste into this sheet"
End Sub

Sub PrevMonth()
    'Previous Month Button
    If Sheets("CalPopUp").Range("A20").value <> Empty Then
        DayName = Sheets("CalPopUp").Range("A20").value & "Day"
        With ActiveSheet.Shapes(DayName)
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Bold = msoFalse
        End With
    End If
    With Sheets("CalPopUp")
        If .Range("A3").value = 1 Then
            .Range("A3").value = 12
            .Range("A2").value = .Range("A2").value - 1
        Else:
            .Range("A3").value = .Range("A3").value - 1
        End If
        If .Range("A6").value > 0 Then
            ActiveSheet.Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = True
        Else:
            ActiveSheet.Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = False
        End If
    End With
End Sub

Sub NextMonth()
    'Next Month button
    If Sheets("CalPopUp").Range("A20").value <> Empty Then
        DayName = Sheets("CalPopUp").Range("A20").value & "Day"
        With ActiveSheet.Shapes(DayName)
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Bold = msoFalse
        End With
    End If
    With Sheets("CalPopUp")
        If .Range("A3").value = 12 Then
            .Range("A3").value = 1
            .Range("A2").value = .Range("A2").value + 1
        Else:
            .Range("A3").value = .Range("A3").value + 1
        End If
        If .Range("A6").value > 0 Then
            ActiveSheet.Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = True
        Else:
            ActiveSheet.Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = False
        End If
    End With
End Sub

Sub PrevYear()
    ThisWorkbook.Sheets("CalPopUp").Range("A2").value = ThisWorkbook.Sheets("CalPopUp").Range("A2").value - 1
End Sub

Sub NextYear()
    ThisWorkbook.Sheets("CalPopUp").Range("A2").value = ThisWorkbook.Sheets("CalPopUp").Range("A2").value + 1
End Sub

''''''''''''''''''''''''''''''''''''''
'''''Select Day Of The Month
''''''''''''''''''''''''''''''''''''''
Sub SelectDay()
    Dim DayNumb As Long, RowNumb As Long, ColNumb As Long
    DayNumb = Replace(Application.Caller, "Day", "")
    RowNumb = Application.WorksheetFunction.RoundUp(DayNumb / 7, 0)
    ColNumb = DayNumb Mod 7 + 1
    If ColNumb = 1 Then ColNumb = 8
    'On Error Resume Next
    If ThisWorkbook.Sheets("CalPopUp").Range("A7").value = Empty Then Exit Sub
    ActiveSheet.Range(ThisWorkbook.Sheets("CalPopUp").Range("A7").value).value = ThisWorkbook.Sheets("CalPopUp").Cells(RowNumb, ColNumb).value
    ActiveSheet.Shapes("Calendar").Visible = msoFalse
    ActiveCell.Offset(0, 1).Select
End Sub

'''''''''''''''''Color Calendar Background''''''''''''''''''''''
Sub CalCol()
    With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
        With Selection.ShapeRange.Fill
            .ForeColor.RGB = ActiveSheet.Shapes(Application.Caller).Fill.ForeColor.RGB
        End With
        ActiveSheet.Range(Sheets("CalPopUp").Range("A7").value).Select
    End With
End Sub

'Create Calendar Sheet on First Run of Calendar
Sub CreateCalSht()
    Dim ColCnt, RowCnt, DayCnt, CalCol As Long
    Dim ws, ActSht As Worksheet
    Set ActSht = ActiveSheet
    'On Error GoTo NoCal
    ActiveSheet.Shapes("Calendar").Copy
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = "CalPopUp"
    ActSht.Activate

    'Reassign Shape Links & Macros
    With ActiveSheet
        UnGroupCal
        '.Unprotect
        .Shapes("PrevYr").OnAction = "'" & ActiveWorkbook.Name & "'!PrevYear"
        .Shapes("NextYr").OnAction = "'" & ActiveWorkbook.Name & "'!NextYear"
        .Shapes("NextRec").OnAction = "'" & ActiveWorkbook.Name & "'!NextMonth"
        .Shapes("NextTri").OnAction = "'" & ActiveWorkbook.Name & "'!NextMonth"
        .Shapes("PrevRec").OnAction = "'" & ActiveWorkbook.Name & "'!PrevMonth"
        .Shapes("PrevTri").OnAction = "'" & ActiveWorkbook.Name & "'!PrevMonth"
        .Shapes("SetBtn").OnAction = "'" & ActiveWorkbook.Name & "'!ShowSettings"
        .Shapes("Month").DrawingObject.formula = "=CalPopUp!A4"
        .Shapes("Year").DrawingObject.formula = "=CalPopUp!A2"
        DayCnt = 1
        For RowCnt = 1 To 6
            For ColCnt = 2 To 8
                .Shapes(DayCnt & "Day").DrawingObject.formula = "=CalPopUp!" & .Cells(RowCnt, ColCnt).Address 'Assigned Linked Cell
                .Shapes(DayCnt & "Day").OnAction = "'" & ActiveWorkbook.Name & "'!SelectDay" 'Assign Macro
                DayCnt = DayCnt + 1
            Next ColCnt
        Next RowCnt
    
        'Assign Color Macros
        For CalCol = 1 To 9
            .Shapes("CalCol" & CalCol).OnAction = "'" & ActiveWorkbook.Name & "'!CalCol" 'Assign Color Macro
        Next CalCol
    End With

    With Sheets("CalPopUp")
        .Paste
        .Visible = xlSheetHidden
    
        'Add in Formulas and Details
        .Range("A1").value = Date                'Set Current Date
        .Range("A2").value = Year(Date)          'Set Current Year
        .Range("A3").value = Month(Date)         'Set Current Month #
        .Range("A4").value = "=INDEX(CalMonths,A3,)"
        .Range("A5").value = "=A4&" & Chr(34) & " " & Chr(34) & "&CalYear"
        .Range("A6").value = "=SUM(B6:H6)"
        .Range("A8").value = "January"
        .Range("A8").AutoFill Destination:=.Range("A8:A19"), Type:=xlFillDefault
        .Range("A20").value = "=IFERROR(INDIRECT(ADDRESS(SUMPRODUCT((B1:H6=A1)*ROW(B1:H6))+6,SUMPRODUCT((B1:H6=A1)*COLUMN(B1:H6)),1,1))," & Chr(34) & Chr(34) & ")"
        'Set Defined Names
        ActiveWorkbook.Names.Add Name:="CalMonths", RefersTo:="=CalPopUp!$A$8:$A$19"
        ActiveWorkbook.Names.Add Name:="CalYear", RefersTo:="=CalPopUp!$A$2"
    
        'Add in Calendar Formulas
      
        .Range("B1").value = "=IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=1,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & ")"
        .Range("C1").value = "=IF(B1<>" & Chr(34) & Chr(34) & ",B1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=2,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("D1").value = "=IF(C1<>" & Chr(34) & Chr(34) & ",C1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=3,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("E1").value = "=IF(D1<>" & Chr(34) & Chr(34) & ",D1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=4,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("F1").value = "=IF(E1<>" & Chr(34) & Chr(34) & ",E1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=5,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("G1").value = "=IF(F1<>" & Chr(34) & Chr(34) & ",F1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=6,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("H1").value = "=IF(G1<>" & Chr(34) & Chr(34) & ",G1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=7,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("B2").value = "=H1+1"
        .Range("C2").value = "=B2+1"
        .Range("C2").AutoFill Destination:=.Range("C2:H2"), Type:=xlFillDefault
        .Range("B2:H2").AutoFill Destination:=.Range("B2:H4"), Type:=xlFillDefault
        .Range("B5").value = "=IF(OR(H4=" & Chr(34) & Chr(34) & ",MONTH(H4+1)<>$A$3)," & Chr(34) & Chr(34) & ",H4+1)"
        .Range("C5").value = "=IFERROR(IF(MONTH(B5+1)<>$A$3," & Chr(34) & Chr(34) & ",B5+1)," & Chr(34) & Chr(34) & ")"
        .Range("B6").value = "=IFERROR(IF(OR(H5=" & Chr(34) & Chr(34) & ",MONTH(H5+1)<>$A$3)," & Chr(34) & Chr(34) & ",H5+1)," & Chr(34) & Chr(34) & ")"
        .Range("C6").value = "=IFERROR(F(MONTH(I5+1)<>$A$3," & Chr(34) & Chr(34) & ",I5+1)," & Chr(34) & Chr(34) & ")"
        .Range("C5:C6").AutoFill Destination:=.Range("C5:H6"), Type:=xlFillDefault
        
        'Set format to Single Day
        .Range("B1:H6").NumberFormat = "d"
        
        'Add in relative Day #'s
        .Range("B7").value = "1"
        .Range("C7").value = "2"
        .Range("B8").value = "8"
        .Range("C8").value = "9"
        .Range("B7:C8").AutoFill Destination:=.Range("B7:H8"), Type:=xlFillDefault
        .Range("B7:H8").AutoFill Destination:=.Range("B7:H12"), Type:=xlFillDefault
        GroupCal
    End With
    Exit Sub
NoCal:
    MsgBox "The Pop-up Calendar does not exist on this worksheet. Please copy the calendar over from another sheet and paste into this sheet"
End Sub

Sub CheckForSheet()
    'Checks for existance of Calendar Pop-up Worksheet
    Dim ws As Worksheet
    On Error GoTo CreateWS
    Set ws = ActiveWorkbook.Sheets("CalPopUp")
    Exit Sub
CreateWS:
    CreateCalSht
End Sub

Sub MacroLinkRemover()
    'PURPOSE: Remove an external workbook reference from all shapes triggering macros
    'Source: www.ExcelForFreelancers.com
    Dim Shp As Shape
    Dim MacroLink, NewLink As String
    Dim SplitLink As Variant

    For Each Shp In ActiveSheet.Shapes           'Loop through each shape in worksheet
  
        'Grab current macro link (if available)
        On Error GoTo NextShp
        MacroLink = Shp.OnAction
    
        'Determine if shape was linking to a macro
        If MacroLink <> "" And InStr(MacroLink, "!") <> 0 Then
            'Split Macro Link at the exclaimation mark (store in Array)
            SplitLink = Split(MacroLink, "!")
        
            'Pull text occurring after exclaimation mark
            NewLink = SplitLink(1)
        
            'Remove any straggling apostrophes from workbook name
            If Right(NewLink, 1) = "'" Then
                NewLink = Left(NewLink, Len(NewLink) - 1)
            End If
        
            'Apply New Link
            Shp.OnAction = NewLink
        End If
NextShp:
    Next Shp
End Sub

Sub CalFormulaReplacement()
    With ActiveSheet
        Dim DayNum, ColNum, RowNum As Long
        Dim Shp As Shape
        ColNum = 2
        RowNum = 1
        For DayNum = 1 To 42
            .Shapes(DayNum & "Day").DrawingObject.formula = "=CalPopUp!" & .Cells(RowNum, ColNum).Address
            ColNum = ColNum + 1
            If ColNum = 9 Then
                ColNum = 2
                RowNum = RowNum + 1
            End If
        Next DayNum
        .Shapes("Month").DrawingObject.formula = "=CalPopUp!$A$4"
        .Shapes("Year").DrawingObject.formula = "=CalPopUp!$A$2"
    End With
End Sub

Sub UnGroupCal()
    On Error Resume Next
    ActiveSheet.Shapes("Calendar").Ungroup
    ActiveSheet.Shapes("NextMonth").Ungroup
    ActiveSheet.Shapes("PrevMonth").Ungroup
    On Error GoTo 0
End Sub

Sub GroupCal()
    ActiveSheet.Shapes.Range(Array("NextTri", "NextRec")).Group.Select
    Selection.ShapeRange.Name = "NextMonth"
    ActiveSheet.Shapes.Range(Array("PrevTri", "PrevRec")).Group.Select
    Selection.ShapeRange.Name = "PrevMonth"
    ActiveSheet.Shapes.Range(Array("Settings", "40Day", "41Day", "39Day", "38Day" _
                                                                         , "42Day", "37Day", "36Day", "CalBack", "Month", "Year", "CalBorder", "1Day", _
                                   "3Day", "14Day", "7Day", "4Day", "2Day", "5Day", "8Day", "10Day", "6Day", _
                                   "13Day", "11Day", "9Day", "12Day", "15Day", "17Day", "20Day", "21Day", "18Day" _
                                                                                                         , "16Day", "19Day", "22Day", "24Day", "26Day", "27Day", "25Day", "23Day", _
                                   "28Day", "29Day", "31Day", "34Day", "35Day", "32Day", "30Day", "33Day", "Sa", _
                                   "Fr", "Th", "We", "Tu", "Mo", "Su", "SetBtn", "CalCol1", "CalCol2", "CalCol3", _
                                   "CalCol4", "CalCol5", "CalCol6", "CalCol7", "CalCol8", "CalCol9", "PrevMonth", _
                                   "NextMonth", "NextYr", "PrevYr")).Visible = msoCTrue
    ActiveSheet.Shapes.Range(Array("Settings", "40Day", "41Day", "39Day", "38Day" _
                                                                         , "42Day", "37Day", "36Day", "CalBack", "Month", "Year", "CalBorder", "1Day", _
                                   "3Day", "14Day", "7Day", "4Day", "2Day", "5Day", "8Day", "10Day", "6Day", _
                                   "13Day", "11Day", "9Day", "12Day", "15Day", "17Day", "20Day", "21Day", "18Day" _
                                                                                                         , "16Day", "19Day", "22Day", "24Day", "26Day", "27Day", "25Day", "23Day", _
                                   "28Day", "29Day", "31Day", "34Day", "35Day", "32Day", "30Day", "33Day", "Sa", _
                                   "Fr", "Th", "We", "Tu", "Mo", "Su", "SetBtn", "CalCol1", "CalCol2", "CalCol3", _
                                   "CalCol4", "CalCol5", "CalCol6", "CalCol7", "CalCol8", "CalCol9", "PrevMonth", _
                                   "NextMonth", "NextYr", "PrevYr")).Select
    Selection.ShapeRange.Group.Select
    Selection.ShapeRange.Name = "Calendar"
    Selection.Name = "Calendar"
    Selection.Placement = xlMove
    ActiveSheet.Shapes("Calendar").Placement = 2
End Sub


