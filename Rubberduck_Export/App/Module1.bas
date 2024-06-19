Attribute VB_Name = "Module1"
Option Explicit

'TBD
'#If VBA7 Then
'    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
'    Private Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As LongPtr, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
'    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
'#Else
'    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
'    Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
'    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'#End If
'
'Private Const MF_BYCOMMAND As Long = &H0&
'Private Const MF_GRAYED As Long = &H1&
'Private Const SC_CLOSE As Long = &HF060&
'
'Sub DisableCloseButton()
'    Dim hWnd As LongPtr
'    Dim hMenu As LongPtr
'
'    ' Get the handle of the Excel window
'    hWnd = FindWindowA("XLMAIN", Application.Caption)
'
'    ' Get the handle of the system menu
'    hMenu = GetSystemMenu(hWnd, False)
'
'    ' Disable the Close button
'    Call EnableMenuItem(hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED)
'End Sub
