Attribute VB_Name = "modInteractionAndView"
Option Explicit

'ActiveWindow.WindowState = xlMinimized
'ActiveWindow.WindowState = xlMaximized
'ActiveWindow.WindowState = xlNormal
'
'Application.userName = "MonkeyCoder"
'
'Application.DisplayAlerts = False
'Application.DisplayAlerts = True
'
'ActiveWindow.FreezePanes = False
'Rows("1:1").Select
'ActiveWindow.FreezePanes = True
'
'Application.DisplayFullScreen = False
'Application.DisplayFullScreen = True
'
'Sheet1.Activate                                  'Turn OFF preview mode
'ActiveWindow.View = xlNormalView
'
'Sheet1.Activate                                  'Turn ON preview mode
'ActiveWindow.View = xlPageBreakPreview
'
'Application.ScreenUpdating = False
'Application.ScreenUpdating = True
'
'ActiveWindow.ScrollColumn = 5
'ActiveWindow.ScrollRow = 5
'
'Application.StatusBar = "I'm working Now!!!"
'
'Application.StatusBar = False
'Application.DisplayStatusBar = True
'
'Sheet1.Range("A1:F15").Select
''Set range zoom
'ActiveWindow.Zoom = True

Public Sub KillFilter()
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If

End Sub

Public Sub StartFilter()
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    End If

End Sub

Sub OffFormulaBar()
    
    Application.DisplayFormulaBar = False

End Sub

Sub OnFormulaBar()
    
    Application.DisplayFormulaBar = True

End Sub

Sub Macro1()

    Call Macro2

End Sub

Private Sub Macro2()

    MsgBox "You can only see Macro1"

End Sub

Sub Hide_Tabs()

    'Hide sheet tabs
    ActiveWindow.DisplayWorkbookTabs = False

End Sub

Sub TurnOffScroll()

    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With

End Sub

Sub TurnOnScroll()
    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With

End Sub

Sub ZoomAll()

    Dim WS As Worksheet

    For Each WS In Worksheets
        WS.Activate
        ActiveWindow.Zoom = 50
    Next

End Sub

Private Sub auto_open()
    
    'This Macro will run every time the workbook opens.
    Application.Caption = ("AutomateExcel.com")

End Sub

