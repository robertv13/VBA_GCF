Attribute VB_Name = "modAPIWindows"
Option Explicit

'API pour le nom de l'utilisateur Windows
Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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

