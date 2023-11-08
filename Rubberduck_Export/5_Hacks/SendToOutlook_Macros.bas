Attribute VB_Name = "SendToOutlook_Macros"
Option Explicit

Sub SendAppointmentToOutlook()
    Dim ApptDate As Date, ApptTime As Date, ContName As String
    Dim ApptName As String, Notes As String
    Dim Dur As Double
    With Schedule
        ApptName = .Range("M4").Value            'Appt Name
        ContName = .Range("M5").Value            'Contact Name
        ApptDate = .Range("M6").Value            'Appt Date
        ApptTime = .Range("M7").Value            'Time
        Dur = .Range("M8").Value * 24 * 60       'Duration (in minutes)
        Notes = .Range("M9").Value               'Notes
    End With
    
    Dim OutApp As Object, OutAppt As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutAppt = OutApp.CreateItem(1)
    With OutAppt
        .Subject = ContName & ": " & ApptName
        .Start = ApptDate + ApptTime
        .Duration = Dur
        .ReminderSet = True
        .ReminderMinutesBeforeStart = 15
        .Body = Notes
        .Save
        .Display
    End With
    Set OutApp = Nothing
    Set OutAppt = Nothing
    MsgBox "Schedule Appt has been sent to Outlook"
End Sub

