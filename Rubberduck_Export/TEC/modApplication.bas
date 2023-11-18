Attribute VB_Name = "modApplication"
Option Explicit

Sub BackToMainMenu()

    ActiveSheet.Visible = xlSheetHidden
    wshMenu.Activate
    wshMenu.Range("B1").Select

End Sub

