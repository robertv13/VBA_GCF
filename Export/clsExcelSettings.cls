VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private calculation As XlCalculation
Private displayStatus As Boolean
Private enableEvents As Boolean
Private screenUpdating As Boolean

' Procedure : TurnOffFunctionality
' Purpose   : Backup Current settings
' https://excelmacromastery.com/
Public Sub Backup()
    calculation = Application.calculation
    displayStatus = Application.DisplayStatusBar
    enableEvents = Application.enableEvents
    screenUpdating = Application.screenUpdating
End Sub

' Procedure : TurnOffFunctionality
' Purpose   : Backup Current settings
' https://excelmacromastery.com/
Public Sub Restore()
    Application.calculation = calculation
    Application.DisplayStatusBar = displayStatus
    Application.enableEvents = enableEvents
    Application.screenUpdating = screenUpdating
End Sub

' Purpose   : Turn off automatic calculations, events and screen updating
' https://excelmacromastery.com/
Public Sub TurnOff()

    Call Backup
    
    Application.calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.enableEvents = False
    Application.screenUpdating = False
    
End Sub

' Purpose   : turn on automatic calculations, events and screen updating
' https://excelmacromastery.com/
Public Sub TurnOn()
    Application.calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.enableEvents = True
    Application.screenUpdating = True
End Sub
