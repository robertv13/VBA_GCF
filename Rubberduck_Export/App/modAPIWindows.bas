Attribute VB_Name = "modAPIWindows"
Option Explicit

'API pour le nom de l'utilisateur Windows
Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'API pour le statut de l'application
#If VBA7 Then
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Public Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

#If VBA7 Then
    Public Declare PtrSafe Function GetLastInputInfo Lib "user32" (lpLastInputInfo As LASTINPUTINFO) As Long
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetLastInputInfo Lib "user32" (lpLastInputInfo As LASTINPUTINFO) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Public Function ApplicationIsActive() As Boolean

    Dim hwnd As LongPtr
    Dim pid As Long
    hwnd = GetForegroundWindow()
    Call GetWindowThreadProcessId(hwnd, pid)
    ApplicationIsActive = (pid = GetCurrentProcessId())
    
End Function


