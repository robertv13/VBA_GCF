Attribute VB_Name = "Admin_Menu_Macros"
Option Explicit

Sub Company_AddLogo()
    Dim LogoPic As FileDialog, PicPath As String
    Set LogoPic = Application.FileDialog(msoFileDialogFilePicker)
    With LogoPic
        .Title = "Please select a logo picture"
        .Filters.Add "Picture Files", "*.jpg,*.png,*.gif", 1
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NoSelection
        PicPath = .SelectedItems(1)
    End With
    Admin.Range("C10").Value = PicPath
NoSelection:
End Sub

Sub EditMode()
    Dim MovInc As Long, BtnMov As Long
    With Invoice
        If .Shapes("EditModeBack").TextFrame2.TextRange.Text = "Off" Then 'Turn On
            .Shapes("EditModeSwitch").Fill.ForeColor.RGB = RGB(0, 176, 80) '  'Turn To Green Color
            .Shapes("EditModeBack").TextFrame2.TextRange.Text = "On"
            MovInc = 1
            .Range("B10").Value = True           'Set Edit Mode To True
        Else                                     'Turn Off
            MovInc = -1
            .Shapes("EditModeSwitch").Fill.ForeColor.RGB = RGB(127, 125, 127) 'Turn to gray color
            .Shapes("EditModeBack").TextFrame2.TextRange.Text = "Off"
            .Range("B10").Value = False          'Set Edit Mode To False
        End If
        Application.ScreenUpdating = False
        For BtnMov = 1 To 29
            .Shapes("EditModeSwitch").Left = .Shapes("EditModeSwitch").Left + MovInc
        Next BtnMov
        Application.ScreenUpdating = True
    End With
End Sub

Sub AppEvents_Stop()
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Sub

Sub AppEvents_Start()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

