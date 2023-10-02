Attribute VB_Name = "PopUp_Calendar"
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

Sub CalendarHide()

    Dim DayNum As Long
    'Hide Calendar, Reset Day Colors
    On Error GoTo NoCal
    ActiveSheet.Shapes("Calendar").Visible = msoFalse
    Sheets("CalPopUp").Range("A7").Value = ""
    If Sheets("CalPopUp").Range("A20").Value <> Empty Then
        For DayNum = 1 To 42
        DayName = DayNum & "Day"
        With ActiveSheet.Shapes(DayName)
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextFrame2.TextRange.Font.Bold = msoFalse
        End With
        Next DayNum
    End If
    
    Exit Sub
    
NoCal:         'If calendar has been removed by accident, paste in backup calendar from  CalPopUp Sheet

End Sub

Sub CalendarShow()

With ActiveSheet
    On Error GoTo ExitSub
    Set SelCell = Selection
    On Error GoTo 0
    'Check if active cell is a valid date
    If IsDate(SelCell.Value) = True Then
        Sheets("CalPopUp").Range("A1").Value = SelCell.Value
        Else: 'If No Date or incorrect Date user current date
        Sheets("CalPopUp").Range("A1").Value = "=Today()"
    End If
        Sheets("CalPopUp").Range("A3").Value = Month(Sheets("CalPopUp").Range("A1").Value) 'Set Month
        Sheets("CalPopUp").Range("A2").Value = Year(Sheets("CalPopUp").Range("A1").Value) 'Set Year
        DayName = Sheets("CalPopUp").Range("A20").Value & "Day"
        UnGroupCal
        If InStr(.Shapes("NextTri").OnAction, "!") <> 0 Or InStr(.Shapes("1Day").DrawingObject.Formula, "]") <> 0 Then   'Run Workbook Link Remover and Cell Link Replacement
            MacroLinkRemover
            CalFormulaReplacement
        End If
        GroupCal
        On Error GoTo NoCal
        With ActiveSheet.Shapes(DayName)
            .Fill.ForeColor.RGB = RGB(252, 213, 180)
            .TextFrame2.TextRange.Font.Bold = msoTrue
        End With
        On Error GoTo NoCal
        .Shapes("Calendar").Visible = msoCTrue
        .Shapes.Range(Array("Settings", "CalCol1", "CalCol2", "CalCol3", "CalCol4", "CalCol5", "CalCol6", "CalCol7", "CalCol8", "CalCol9")).Visible = False '
        .Shapes("Calendar").Left = SelCell.Left
        If SelCell.Row < 5 And ActiveWindow.ScrollRow > 6 Then .Shapes("Calendar").Top = SelCell.Offset(ActiveWindow.ScrollRow - 4, 0).Top Else: .Shapes("Calendar").Top = SelCell.Offset(1, 0).Top
            If Sheets("CalPopUp").Range("A6").Value > 0 Then
                    .Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = True
                Else:
                    .Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = False
            End If
        Sheets("CalPopUp").Range("A7").Value = SelCell.Address
        ActiveCell.Select
End With

Exit Sub

NoCal:
    MsgBox "Le calendrier n'existe pas dans ce classeur. Veuillez copier le calendrier à partir d'un autre classeur."
ExitSub:
End Sub

Sub PrevMonth()

    'Previous Month Button
    If Sheets("CalPopUp").Range("A20").Value <> Empty Then
    DayName = Sheets("CalPopUp").Range("A20").Value & "Day"
        With ActiveSheet.Shapes(DayName)
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextFrame2.TextRange.Font.Bold = msoFalse
        End With
    End If
    With Sheets("CalPopUp")
        If .Range("A3").Value = 1 Then
            .Range("A3").Value = 12
            .Range("A2").Value = .Range("A2").Value - 1
        Else:
            .Range("A3").Value = .Range("A3").Value - 1
        End If
        If .Range("A6").Value > 0 Then
            ActiveSheet.Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = True
            Else:
            ActiveSheet.Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = False
        End If
    End With
    
End Sub

Sub NextMonth()

    'Next Month button
    If Sheets("CalPopUp").Range("A20").Value <> Empty Then
    DayName = Sheets("CalPopUp").Range("A20").Value & "Day"
        With ActiveSheet.Shapes(DayName)
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Bold = msoFalse
        End With
    End If
    With Sheets("CalPopUp")
        If .Range("A3").Value = 12 Then
            .Range("A3").Value = 1
            .Range("A2").Value = .Range("A2").Value + 1
        Else:
            .Range("A3").Value = .Range("A3").Value + 1
        End If
        If .Range("A6").Value > 0 Then
            ActiveSheet.Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = True
            Else:
            ActiveSheet.Shapes.Range(Array("36Day", "37Day", "38Day", "39Day", "40Day", "41Day", "42Day")).Visible = False
        End If
    End With
End Sub

''''''''''''''''''''''''''''''''''''''
'''''Individual Day Selection Buttons
''''''''''''''''''''''''''''''''''''''

Sub DayBtn1()

    With Sheets("CalPopUp")
        If .Range("B1").Value <> Empty And .Range("A7").Value <> Empty Then
            ActiveSheet.Range(.Range("A7").Value).Value = .Range("B1").Value
        End If
        ActiveSheet.Shapes("Calendar").Visible = msoFalse
        ActiveCell.Offset(0, 1).Select
    End With
End Sub

Sub DayBtn2()

    With Sheets("CalPopUp")
        If .Range("C1").Value <> Empty And .Range("A7").Value <> Empty Then
            ActiveSheet.Range(.Range("A7").Value).Value = .Range("C1").Value
        End If
        ActiveSheet.Shapes("Calendar").Visible = msoFalse
        ActiveCell.Offset(0, 1).Select
    End With
End Sub

Sub DayBtn3()

    With Sheets("CalPopUp")
        If .Range("D1").Value <> Empty And .Range("A7").Value <> Empty Then
            ActiveSheet.Range(.Range("A7").Value).Value = .Range("D1").Value
        End If
        ActiveSheet.Shapes("Calendar").Visible = msoFalse
        ActiveCell.Offset(0, 1).Select
    End With
End Sub
Sub DayBtn4()
With Sheets("CalPopUp")
If .Range("E1").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("E1").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn5()
With Sheets("CalPopUp")
If .Range("F1").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("F1").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn6()
With Sheets("CalPopUp")
If .Range("G1").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("G1").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn7()
With Sheets("CalPopUp")
If .Range("H1").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("H1").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn8()
With Sheets("CalPopUp")
If .Range("B2").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("B2").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn9()
With Sheets("CalPopUp")
If .Range("C2").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("C2").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn10()
With Sheets("CalPopUp")
If .Range("D2").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("D2").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn11()
With Sheets("CalPopUp")
If .Range("E2").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("E2").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn12()
With Sheets("CalPopUp")
If .Range("F2").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("F2").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn13()
With Sheets("CalPopUp")
If .Range("G2").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("G2").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn14()
With Sheets("CalPopUp")
If .Range("H2").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("H2").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn15()
With Sheets("CalPopUp")
If .Range("B3").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("B3").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn16()
With Sheets("CalPopUp")
If .Range("C3").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("C3").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn17()
With Sheets("CalPopUp")
If .Range("D3").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("D3").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn18()
With Sheets("CalPopUp")
If .Range("E3").Value <> Empty And .Range("A7").Value <> Empty Then
   ActiveSheet.Range(.Range("A7").Value).Value = .Range("E3").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn19()
With Sheets("CalPopUp")
If .Range("F3").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("F3").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn20()
With Sheets("CalPopUp")
If .Range("G3").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("G3").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn21()
With Sheets("CalPopUp")
If .Range("H3").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("H3").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn22()
With Sheets("CalPopUp")
If .Range("B4").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("B4").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn23()
With Sheets("CalPopUp")
If .Range("C4").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("C4").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn24()
With Sheets("CalPopUp")
If .Range("D4").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("D4").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn25()
With Sheets("CalPopUp")
If .Range("E4").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("E4").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn26()
With Sheets("CalPopUp")
If .Range("F4").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("F4").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn27()
With Sheets("CalPopUp")
If .Range("G4").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("G4").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn28()
With Sheets("CalPopUp")
If .Range("H4").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("H4").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn29()
With Sheets("CalPopUp")
If .Range("B5").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("B5").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn30()
With Sheets("CalPopUp")
If .Range("C5").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("C5").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn31()
With Sheets("CalPopUp")
If .Range("D5").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("D5").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn32()
With Sheets("CalPopUp")
If .Range("E5").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("E5").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn33()
With Sheets("CalPopUp")
If .Range("F5").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("F5").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn34()
With Sheets("CalPopUp")
If .Range("G5").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("G5").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn35()
With Sheets("CalPopUp")
If .Range("H5").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("H5").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn36()
With Sheets("CalPopUp")
If .Range("B6").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("B6").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn37()
With Sheets("CalPopUp")
If .Range("C6").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("C6").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn38()
With Sheets("CalPopUp")
If .Range("D6").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("D6").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn39()
With Sheets("CalPopUp")
If .Range("E6").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("E6").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn40()
With Sheets("CalPopUp")
If .Range("F6").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("F6").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn41()
With Sheets("CalPopUp")
If .Range("G6").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("G6").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub
Sub DayBtn42()
With Sheets("CalPopUp")
If .Range("H6").Value <> Empty And .Range("A7").Value <> Empty Then
    ActiveSheet.Range(.Range("A7").Value).Value = .Range("H6").Value
End If
ActiveSheet.Shapes("Calendar").Visible = msoFalse
ActiveCell.Offset(0, 1).Select
End With
End Sub

'''''''''''''''''Color Calendar Background''''''''''''''''''''''
Sub CalCol1()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(234, 234, 234)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub
Sub CalCol2()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(197, 190, 151)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub
Sub CalCol3()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(141, 180, 227)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub
Sub CalCol4()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(184, 204, 228)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub
Sub CalCol5()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(230, 185, 184)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub
Sub CalCol6()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(215, 228, 188)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub
Sub CalCol7()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(204, 192, 218)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub
Sub CalCol8()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(182, 221, 232)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub
Sub CalCol9()
With ActiveSheet.Shapes.Range(Array("CalBack", "Settings")).Select
    With Selection.ShapeRange.Fill
    .ForeColor.RGB = RGB(252, 213, 180)
    End With
    ActiveSheet.Range(Sheets("CalPopUp").Range("A7").Value).Select
End With
End Sub

'Create Calendar Sheet on First Run of Calendar
Sub CreateCalSht()

Dim ColCnt, RowCnt, DayCnt, CalCol As Long
Dim ws, ActSht As Worksheet
Set ActSht = ActiveSheet
'On Error GoTo NoCal
ActiveSheet.Shapes("Calendar").Copy
Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
ws.Name = "CalPopUp"
ActSht.Activate

'Reassign Shape Links & Macros
With ActiveSheet
UnGroupCal
    '.Unprotect
    .Shapes("NextRec").OnAction = "'" & ActiveWorkbook.Name & "'!NextMonth"
    .Shapes("NextTri").OnAction = "'" & ActiveWorkbook.Name & "'!NextMonth"
    .Shapes("PrevRec").OnAction = "'" & ActiveWorkbook.Name & "'!PrevMonth"
    .Shapes("PrevTri").OnAction = "'" & ActiveWorkbook.Name & "'!PrevMonth"
    .Shapes("SetBtn").OnAction = "'" & ActiveWorkbook.Name & "'!ShowSettings"
    .Shapes("MonthYear").DrawingObject.Formula = "=CalPopUp!A5"
    DayCnt = 1
    For RowCnt = 1 To 6
        For ColCnt = 2 To 8
                .Shapes(DayCnt & "Day").DrawingObject.Formula = "=CalPopUp!" & .Cells(RowCnt, ColCnt).Address 'Assigned Linked Cell
                .Shapes(DayCnt & "Day").OnAction = "'" & ActiveWorkbook.Name & "'!DayBtn" & DayCnt 'Assign Macro
                DayCnt = DayCnt + 1
        Next ColCnt
    Next RowCnt
    
    'Assign Color Macros
    For CalCol = 1 To 9
     .Shapes("CalCol" & CalCol).OnAction = "'" & ActiveWorkbook.Name & "'!CalCol" & CalCol
    Next CalCol
End With

With Sheets("CalPopUp")
    .Paste
        .Visible = xlSheetHidden
    
    'Add in Formulas and Details
        .Range("A4").Value = "=INDEX(CalMonths,A3,)"
        .Range("A5").Value = "=A4&" & Chr(34) & " " & Chr(34) & "&CalYear"
        .Range("A6").Value = "=SUM(B6:H6)"
        .Range("A8").Value = "January"
        .Range("A8").AutoFill Destination:=.Range("A8:A19"), Type:=xlFillDefault
        .Range("A20").Value = "=IFERROR(INDIRECT(ADDRESS(SUMPRODUCT((B1:H6=A1)*ROW(B1:H6))+6,SUMPRODUCT((B1:H6=A1)*COLUMN(B1:H6)),1,1))," & Chr(34) & Chr(34) & ")"
    'Set Defined Names
    ActiveWorkbook.Names.Add Name:="CalMonths", RefersTo:="=CalPopUp!$A$8:$A$19"
    ActiveWorkbook.Names.Add Name:="CalYear", RefersTo:="=CalPopUp!$A$2"
    
    'Add in Calendar Formulas
        .Range("B1").Value = "=IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=1,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & ")"
        .Range("C1").Value = "=IF(B1<>" & Chr(34) & Chr(34) & ",B1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=2,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("D1").Value = "=IF(C1<>" & Chr(34) & Chr(34) & ",C1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=3,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("E1").Value = "=IF(D1<>" & Chr(34) & Chr(34) & ",D1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=4,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("F1").Value = "=IF(E1<>" & Chr(34) & Chr(34) & ",E1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=5,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("G1").Value = "=IF(F1<>" & Chr(34) & Chr(34) & ",F1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=6,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("H1").Value = "=IF(G1<>" & Chr(34) & Chr(34) & ",G1+1,IF(WEEKDAY(DATE(CalYear,MATCH($A$4,CalMonths,0),1))=7,DATE(CalYear,MATCH($A$4,CalMonths,0),1)," & Chr(34) & Chr(34) & "))"
        .Range("B2").Value = "=H1+1"
        .Range("C2").Value = "=B2+1"
        .Range("C2").AutoFill Destination:=.Range("C2:H2"), Type:=xlFillDefault
        .Range("B2:H2").AutoFill Destination:=.Range("B2:H4"), Type:=xlFillDefault
        .Range("B5").Value = "=IF(OR(H4=" & Chr(34) & Chr(34) & ",MONTH(H4+1)<>$A$3)," & Chr(34) & Chr(34) & ",H4+1)"
        .Range("C5").Value = "=IFERROR(IF(MONTH(B5+1)<>$A$3," & Chr(34) & Chr(34) & ",B5+1)," & Chr(34) & Chr(34) & ")"
        .Range("B6").Value = "=IFERROR(IF(OR(H5=" & Chr(34) & Chr(34) & ",MONTH(H5+1)<>$A$3)," & Chr(34) & Chr(34) & ",H5+1)," & Chr(34) & Chr(34) & ")"
        .Range("C6").Value = "=IFERROR(F(MONTH(I5+1)<>$A$3," & Chr(34) & Chr(34) & ",I5+1)," & Chr(34) & Chr(34) & ")"
        .Range("C5:C6").AutoFill Destination:=.Range("C5:H6"), Type:=xlFillDefault
        
        'Set format to Single Day
        .Range("B1:H6").NumberFormat = "d"
        
        'Add in relative Day #'s
        .Range("B7").Value = "1"
        .Range("C7").Value = "2"
        .Range("B8").Value = "8"
        .Range("C8").Value = "9"
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
    
    For Each Shp In ActiveSheet.Shapes 'Loop through each shape in worksheet
      
        'Grab current macro link (if available)
        On Error GoTo NextShp
        MacroLink = Shp.OnAction
      
        'Determine if shape was linking to a macro
        If MacroLink <> "" And InStr(MacroLink, "!") <> 0 Then
            'Split Macro Link at the exclamation mark (store in Array)
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
            .Shapes(DayNum & "Day").DrawingObject.Formula = "=CalPopUp!" & .Cells(RowNum, ColNum).Address
            ColNum = ColNum + 1
            If ColNum = 9 Then
               ColNum = 2
               RowNum = RowNum + 1
            End If
         Next DayNum
         .Shapes("MonthYear").DrawingObject.Formula = "=CalPopUp!$A$5"
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
            , "42Day", "37Day", "36Day", "CalBack", "MonthYear", "CalBorder", "1Day", _
            "3Day", "14Day", "7Day", "4Day", "2Day", "5Day", "8Day", "10Day", "6Day", _
            "13Day", "11Day", "9Day", "12Day", "15Day", "17Day", "20Day", "21Day", "18Day" _
            , "16Day", "19Day", "22Day", "24Day", "26Day", "27Day", "25Day", "23Day", _
            "28Day", "29Day", "31Day", "34Day", "35Day", "32Day", "30Day", "33Day", "Sa", _
            "Fr", "Th", "We", "Tu", "Mo", "Su", "SetBtn", "CalCol1", "CalCol2", "CalCol3", _
            "CalCol4", "CalCol5", "CalCol6", "CalCol7", "CalCol8", "CalCol9", "PrevMonth", _
            "NextMonth")).Visible = msoCTrue
            ActiveSheet.Shapes.Range(Array("Settings", "40Day", "41Day", "39Day", "38Day" _
            , "42Day", "37Day", "36Day", "CalBack", "MonthYear", "CalBorder", "1Day", _
            "3Day", "14Day", "7Day", "4Day", "2Day", "5Day", "8Day", "10Day", "6Day", _
            "13Day", "11Day", "9Day", "12Day", "15Day", "17Day", "20Day", "21Day", "18Day" _
            , "16Day", "19Day", "22Day", "24Day", "26Day", "27Day", "25Day", "23Day", _
            "28Day", "29Day", "31Day", "34Day", "35Day", "32Day", "30Day", "33Day", "Sa", _
            "Fr", "Th", "We", "Tu", "Mo", "Su", "SetBtn", "CalCol1", "CalCol2", "CalCol3", _
            "CalCol4", "CalCol5", "CalCol6", "CalCol7", "CalCol8", "CalCol9", "PrevMonth", _
            "NextMonth")).Select
        Selection.ShapeRange.Group.Select
        Selection.ShapeRange.Name = "Calendar"
        Selection.Name = "Calendar"
        Selection.Placement = xlMove
        ActiveSheet.Shapes("Calendar").Placement = 2

End Sub







