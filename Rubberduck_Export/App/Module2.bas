Attribute VB_Name = "Module2"
Option Explicit

Sub AddCustomMenu()
    Dim cb As CommandBar
    Dim cbc As CommandBarControl
    
    ' Delete the menu if it already exists
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("Custom Exit").delete
    On Error GoTo 0
    
    ' Add a new menu item to the Worksheet Menu Bar
    Set cb = Application.CommandBars("Worksheet Menu Bar")
    Set cbc = cb.Controls.add(Type:=msoControlButton, Temporary:=True)
    
    With cbc
        .Caption = "Custom Exit"
        .OnAction = "CustomExit"
    End With
End Sub

Sub CustomExit()
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit Application")
    
    If answer = vbYes Then
        ThisWorkbook.Saved = True
        Application.Quit
    End If
    
End Sub

