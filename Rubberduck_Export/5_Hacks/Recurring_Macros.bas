Attribute VB_Name = "Recurring_Macros"
Option Explicit

Sub CreateRecurringAppt()
    Dim ApptDate As Date
    Dim Freq As String, FreqQty As Long, ApptQty As Long, ApptNumb As Long
    With Schedule
        If .Range("B10").Value = Empty Then
            MsgBox "Please make sure to save this Appt before creating recurring Schedule Appt"
            Exit Sub
        End If
        If Application.WorksheetFunction.CountA(.Range("M12:M15")) < 4 Then
            MsgBox "Please make sure to fill in all recurring fields"
            Exit Sub
        End If
        FreqQty = .Range("M12").Value            'Frequency Qty
        Freq = .Range("M13").Value               'Frequency
        ApptDate = .Range("M14").Value           'Start Date
        ApptQty = .Range("M15").Value            ' Number of Schedule Appts
        .Range("B12").Value = True               'Set Recurring to True
        For ApptNumb = 1 To ApptQty
            .Range("M6").Value = ApptDate        'Cal. Appt Date
            .Range("B9").ClearContents           'Clear Existing Appt ID
            Appt_SaveUpdate                      'Save Schedule Appt
            Select Case Freq
            Case Is = "Minute(s)"
                ApptDate = DateAdd("n", FreqQty, ApptDate)
            Case Is = "Hour(s)"
                ApptDate = DateAdd("h", FreqQty, ApptDate)
            Case Is = "Day(s)"
                ApptDate = DateAdd("d", FreqQty, ApptDate)
            Case Is = "Week(s)"
                ApptDate = DateAdd("ww", FreqQty, ApptDate)
            Case Is = "Month(s)"
                ApptDate = DateAdd("m", FreqQty, Now)
            End Select
        Next ApptNumb
        .Range("B12").Value = False              'Set Recurring to false
        Schedule_Refresh                         'Refresh Schedule
        MsgBox ApptQty & " Schedule Appts have been created"
    End With
End Sub

