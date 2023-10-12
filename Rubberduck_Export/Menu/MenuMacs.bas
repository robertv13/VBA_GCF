Attribute VB_Name = "MenuMacs"
Option Explicit

Dim Wdth As Long
Global Const gMinPos As Integer = 32
Global Const gMaxPos As Integer = 105

Sub SlideOut_TEC()
    With ActiveSheet.Shapes("btnTEC")
        For Wdth = gMinPos To gMaxPos
            .Height = Wdth
            ActiveSheet.Shapes("imgIconeSablier").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "T.E.C."
    End With
End Sub

Sub SlideIn_TEC()
    With ActiveSheet.Shapes("btnTEC")
        For Wdth = gMaxPos To gMinPos Step -1
            .Height = Wdth
            .Left = Wdth - gMinPos
            ActiveSheet.Shapes("imgIconeSablier").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub

Sub SlideOut_Schedule()
    With ActiveSheet.Shapes("btnSchedule")
        For Wdth = gMinPos To gMaxPos
            .Height = Wdth
            ActiveSheet.Shapes("imgIconeSchedule").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "Schedule"
    End With
End Sub

Sub SlideIn_Schedule()
    With ActiveSheet.Shapes("btnSchedule")
        For Wdth = gMaxPos To gMinPos Step -1
            .Height = Wdth
            .Left = Wdth - gMinPos
            ActiveSheet.Shapes("imgIconeSchedule").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub

Sub SlideOut_Email()
    With ActiveSheet.Shapes("btnEmails")
        For Wdth = gMinPos To gMaxPos
            .Height = Wdth
            ActiveSheet.Shapes("imgIconeEmails").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "Email"
    End With
End Sub

Sub SlideIn_Email()
    With ActiveSheet.Shapes("btnEmails")
        For Wdth = gMaxPos To gMinPos Step -1
            .Height = Wdth
            .Left = Wdth - gMinPos
            ActiveSheet.Shapes("imgIconeEmails").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub

Sub SlideOut_Reports()
    With ActiveSheet.Shapes("btnReports")
        For Wdth = gMinPos To gMaxPos '
            .Height = Wdth
            ActiveSheet.Shapes("imgIconeReports").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "Reports"
    End With
End Sub

Sub SlideIn_Reports()
    With ActiveSheet.Shapes("btnReports")
        For Wdth = gMaxPos To gMinPos Step -1
            .Height = Wdth
            .Left = Wdth - gMinPos
            ActiveSheet.Shapes("imgIconeReports").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub

Sub SlideOut_Graphs()
    With ActiveSheet.Shapes("btnGraphs")
        For Wdth = gMinPos To gMaxPos
            .Height = Wdth
            ActiveSheet.Shapes("imgIconeGraphs").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "Graphs"
    End With
End Sub

Sub SlideIn_Graphs()
    With ActiveSheet.Shapes("btnGraphs")
        For Wdth = gMaxPos To gMinPos Step -1
            .Height = Wdth
            .Left = Wdth - gMinPos
            ActiveSheet.Shapes("imgIconeGraphs").Left = Wdth - gMinPos
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub

Sub TEC_Click()

    MsgBox "This is the TEC button"

End Sub

Sub Schedule_Click()

    MsgBox "This is the Schedule button"

End Sub

Sub Email_Click()

    MsgBox "This is the Emails button"

End Sub

Sub Reports_Click()

    MsgBox "This is the Reports button"

End Sub

Sub Graphs_Click()

    MsgBox "This is the Graphs button"

End Sub



