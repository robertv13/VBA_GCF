Attribute VB_Name = "Schedule_Macros"
Option Explicit
'Global Variables
Dim CalRow As Long, CalCol As Long, LastRow As Long, LastResultRow As Long, ResultRow As Long
Dim ApptNumb As Long, DayApptCount As Long, ApptTop As Long, NextDateRow As Long, ApptLimit As Long
Dim ApptColor As String, ApptID As String, ContName As String
Dim ApptShp As Shape
Dim ApptDate As Date, ApptTime As String
Dim ApptWidth As Double, ApptLeft As Double

Sub Schedule_Refresh()
    'Clear all existing Appt shapes from Schedule sheet
    For Each ApptShp In Schedule.Shapes
        If InStr(ApptShp.Name, "CalAppt") > 0 Then ApptShp.Delete
    Next ApptShp
    ApptColor = Admin.Range("F7").Interior.Color 'Set Appt Shape Color
    ApptNumb = 1                                 'Set default Appt # to 1
    ApptWidth = 1                                'Set Default to 1
    With ApptsDB
        LastRow = .Range("A99999").End(xlUp).Row 'Last Appt Row
        If LastRow < 4 Then Exit Sub
        Application.ScreenUpdating = False
        .Range("A3:F" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("I2:J3"), CopyToRange:=.Range("M2:Q2"), Unique:=True
        LastResultRow = .Range("M99999").End(xlUp).Row
        If LastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=ApptsDB.Range("O3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add Key:=ApptsDB.Range("P3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Sort Based On Time
            .SetRange ApptsDB.Range("M3:Q" & LastResultRow) 'Set Range
            .Apply                               'Apply Sort
        End With
NoSort:
        For ResultRow = 3 To LastResultRow
            ApptID = .Range("M" & ResultRow).Value 'Appt ID
            ContName = .Range("N" & ResultRow).Value 'Appt Name
            ApptDate = .Range("O" & ResultRow).Value 'Appt Date
            ApptTime = Format(.Range("P" & ResultRow).Value, "h:mma/p")
            DayApptCount = Application.WorksheetFunction.CountIf([ItemDate_Results], ApptDate) 'Get # of Appts in a single Date
            If DayApptCount > 1 And ApptNumb = 1 Then ApptLimit = DayApptCount
            If DayApptCount > 5 Then
                If DayApptCount > 10 And ApptNumb = 1 Then 'On First Instance of Multiple any day count greater than 10
                    NextDateRow = ResultRow + DayApptCount - 1 'Determine Next Day Result Row
                    ApptLimit = 10               'Set Appt Limit to 10
                End If
                ApptWidth = 0.5                  'Set Appt with to 1/2
            End If
            For CalRow = 4 To 34 Step 6
                For CalCol = 4 To 10
                    If Schedule.Cells(CalRow, CalCol).Value = ApptDate Then 'Day Found
                        Schedule.Shapes("SampleApptShp").Duplicate.Name = "CalAppt" & ApptID
                        With Schedule.Shapes("CalAppt" & ApptID)
                            .Left = Schedule.Cells(CalRow, CalCol).Left + ApptLeft 'Set Left Pos.
                            .Top = Schedule.Cells(CalRow + ApptNumb + ApptTop, CalCol).Top + 1
                            .Width = Schedule.Cells(CalRow + ApptNumb, CalCol).Width * ApptWidth
                            .Height = Schedule.Cells(CalRow + ApptNumb, CalCol).Height
                            .TextFrame2.TextRange.Text = ApptTime & ":" & ContName 'Text inside shape
                            .Fill.ForeColor.RGB = ApptColor 'Set Appt Color
                            .OnAction = "Schedule_Appt_Select" 'Macro to select Appt
                        End With
                        If ApptNumb >= ApptLimit Then
                            ApptNumb = 1         'Reset
                            ApptWidth = 1        'Reset Appt Width
                            ApptLimit = 1        'Reset Appt Limit
                            If NextDateRow <> 0 Then
                                ResultRow = NextDateRow
                                NextDateRow = 0  'Reset Next Date Row to 0
                            End If               '
                        Else
                            ApptNumb = ApptNumb + 1 'Increment by 1
                        End If
                        If ApptNumb <= 5 Then
                            ApptLeft = 0         'Set Left Position
                            ApptTop = 0
                        Else                     'Appt Number from 6 to 10
                            ApptLeft = Schedule.Cells(CalRow + ApptNumb, CalCol).Width / 2
                            ApptTop = -5
                        End If
                    End If
                Next CalCol
            Next CalRow
        Next ResultRow
    End With
    Application.ScreenUpdating = True
End Sub

Sub Schedule_Appt_Select()
    With Schedule
        ApptID = Replace(Application.Caller, Left(Application.Caller, 7), "") 'Extract Appt ID
        For Each ApptShp In .Shapes              'Reset All Colors back to default color
            If InStr(ApptShp.Name, "CalAppt") > 0 Then ApptShp.Fill.ForeColor.RGB = Admin.Range("F7").Interior.Color
        Next ApptShp
        .Shapes("CalAppt" & ApptID).Fill.ForeColor.RGB = Admin.Range("F9").Interior.Color 'Select Appt. Color
        .Range("B9").Value = ApptID
        'Drag & Drop Settings
        Schedule.Range("B9").Value = ApptID
        .Range("B15").Value = Schedule.Shapes(Application.Caller).Left 'set Selected Left Position
        .Range("B16").Value = Schedule.Shapes(Application.Caller).Top ' Set Selected Top Position
        .Range("B1").Value = False               'Set Schedule Move to False
        .Shapes("CalAppt" & ApptID).Select
        Appt_Load
        Schedule_CheckForMove
    End With
End Sub

Sub Schedule_PrevMonth()
    With Schedule
        If .Range("B2").Value = 1 And .Range("B5").Value = 1 Then
            MsgBox "You are at the first year and first month available"
            Exit Sub
        End If
        If .Range("B5").Value = 1 Then           'January Month
            .Range("B2").Value = .Range("B2").Value - 1 'Reduce Year by 1
            .Range("B5").Value = 12              'Set To December
        Else                                     'Not January
            .Range("B5").Value = .Range("B5").Value - 1 'Reduce month by 1
        End If
        Schedule_Refresh                         'Refresh Schedule
    End With
End Sub

Sub Schedule_NextMonth()
    With Schedule
        If .Range("B2").Value = 7 And .Range("B5").Value = 12 Then
            MsgBox "You are at the last year and first last available"
            Exit Sub
        End If
        If .Range("B5").Value = 12 Then          'December Month
            .Range("B2").Value = .Range("B2").Value + 1 'Increase Year by 1
            .Range("B5").Value = 1               'Set To January
        Else                                     'Not January
            .Range("B5").Value = .Range("B5").Value + 1 'Increase month by 1
        End If
        Schedule_Refresh                         'Refresh Schedule
    End With
End Sub

Sub Schedule_ThisMonth()
    With Schedule
        .Range("B5").Value = Month(Date)         'Set Current Month #
        .Range("B2").Value = Admin.Range("Years").Find(Year(Date), , xlFormulas, xlWhole).Row - 3
        Schedule_Refresh                         'Refresh Schedule
    End With
End Sub

Sub PrintSchedule()
    Schedule.PrintOut , , , False, True, , , , False
End Sub

Sub Schedule_CheckForMove()
    Dim DestRow As Long, DestCol As Long, ApptRow As Long, ApptCol As Long, CountDelay As Long
    With Schedule
        If .Range("B10").Value = Empty Then Exit Sub
        ApptID = .Range("B9").Value              'Appt ID
    
        For CountDelay = 1 To 100000
            DoEvents
            If .Range("B1").Value = True Then End 'Exit Loop on Move Appt Move = True
    
            With .Shapes("CalAppt" & ApptID)
                If .Left <> Schedule.Range("B15").Value Or .Top <> Schedule.Range("B16").Value Then 'Move Detect
                    'Check for incorrect move
                    If .Left < Schedule.Range("D1").Left Or .Left > Schedule.Range("K1").Left Or .Top > Schedule.Range("A39").Top Or .Top < Schedule.Range("A4").Top - 1 Then
                        MsgBox "Plase make sure to move the Schedule Appt to a correct Schedule date on the Schedule"
                        Schedule_Refresh
                        End
                        Exit Sub
                    End If
                    
                    DestRow = Schedule.Shapes("CalAppt" & ApptID).TopLeftCell.Row 'Row Destination
                    DestCol = Schedule.Shapes("CalAppt" & ApptID).TopLeftCell.Column 'Column Dest.
                    DestRow = DestRow - (DestRow + 2) Mod 6 'Date Row
                    ApptDate = Schedule.Cells(DestRow, DestCol).Value 'Schedule Date
                    If ApptDate = 0 Then
                        MsgBox "Please move the Appt to a day with an existing Date"
                        Schedule_Refresh
                        Exit Sub
                    End If
                    Schedule.Range("M6").Value = ApptDate 'new Appt dates
                    Appt_SaveUpdate              'Save Appt Changes
                    Schedule.Range("B1").Value = True 'Set Move Appt to true
                    End
                End If
            End With
        Next CountDelay
        .Range("B1").Value = True                'Set Move Appt Movement to True
    End With
End Sub

