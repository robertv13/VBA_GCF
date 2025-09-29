Attribute VB_Name = "modAPIWindows"
Option Explicit

'API pour le nom de l'utilisateur Windows
Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'API pour détecter les touches spéciales - 2025-09-10 @ 07:46
#If VBA7 Then
    Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Declare Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

