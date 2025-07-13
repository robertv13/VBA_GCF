Attribute VB_Name = "modAPIWindows"
Option Explicit

'API pour le nom de l'utilisateur Windows
Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

