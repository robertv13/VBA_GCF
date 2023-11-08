Attribute VB_Name = "Appt_Macros"
Option Explicit

Dim ApptRow As Long, ApptCol As Long

Sub Appt_Load()
    With Schedule
        If .Range("B10").Value = Empty Then
            MsgBox "Please select a correct Schedule Appt to load"
            Exit Sub
        End If
        ApptRow = .Range("B10").Value            'Appt Row
        For ApptCol = 2 To 7
            .Range("M" & ApptCol + 2).Value = ApptsDB.Cells(ApptRow, ApptCol).Value 'Bring Over data
        Next ApptCol
    End With
End Sub

Sub Appt_New()
    Schedule.Range("B9,M4:M9").ClearContents     'Clear Out
End Sub

Sub Appt_SaveUpdate()
    With Schedule
        .Range("B1").Value = True                'Exit out of Check For Change Loop
        If Application.WorksheetFunction.CountA(.Range("M4:M6")) < 3 Then
            MsgBox "Please make sure to add in an Appt Name, Date & Time"
            Exit Sub
        End If
        If .Range("B10").Value = Empty Then      'New Appt
            ApptRow = ApptsDB.Range("A99999").End(xlUp).Row + 1 'First Avail. Row
            .Range("B9").Value = .Range("B11").Value 'Set Appt ID
            ApptsDB.Range("A" & ApptRow).Value = .Range("B9").Value 'Appt ID
        Else                                     'Existing Appt
            ApptRow = .Range("B10").Value        'Existing Appt Row
        End If
        For ApptCol = 2 To 7
            ApptsDB.Cells(ApptRow, ApptCol).Value = .Range("M" & ApptCol + 2).Value 'Save/ Update data
        Next ApptCol
        If .Range("B12").Value = False Then      'Refresh & Set Saved Message except for recurring
            Schedule_Refresh                     'Refresh Schedule
            MsgBox "Appointment Saved"
        End If
    End With
End Sub

Sub Appt_Delete()
    If MsgBox("Are you sure you want to delete this Schedule Appointment?", vbYesNo, "Delete Appt") = vbNo Then Exit Sub
    With Schedule
        .Range("B1").Value = True                'Exit out of Check For Change Loop
        If .Range("B10").Value = Empty Then GoTo NotSaved
        ApptRow = .Range("B10").Value            'Appt row
        ApptsDB.Range(ApptRow & ":" & ApptRow).EntireRow.Delete 'Delete ROw
NotSaved:
        Appt_New
        Schedule_Refresh                         'Refresh Schedule
    End With
End Sub

