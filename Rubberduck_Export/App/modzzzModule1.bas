Attribute VB_Name = "modzzzModule1"
Option Explicit
Option Private Module

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
Public Declare PtrSafe Function GetWindowLong Lib "user32" _
                                    Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                    ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" _
                                    Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                    ByVal nIndex As Long, _
                                    ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, _
                                    ByVal lpWindowName As String) As Long

Sub HideTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub


