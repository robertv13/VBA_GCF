Attribute VB_Name = "Module4"
Option Explicit

Sub Unprotect_Sheet_without_Password()
    
    Dim P As String
    P = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    'P = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    'P = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    Dim x1 As Integer, x2 As Integer, x3 As Integer, x4 As Integer, x5 As Integer, x6 As Integer
    
    Dim root As String, password As String
    root = InputBox("Quelle est la racine du mot de passe ? ")
    
    For x1 = 1 To Len(P)
        For x2 = 1 To Len(P)
            Debug.Print password
            For x3 = 1 To Len(P)
                For x4 = 1 To Len(P)
                    password = Trim(root & Mid(P, x1, 1) & Mid(P, x2, 1) & Mid(P, x3, 1) & Mid(P, x4, 1))
                    
                    On Error Resume Next
                    ActiveSheet.Unprotect password
                    On Error GoTo 0
                    If ActiveSheet.ProtectContents = False Then
                        MsgBox "Password is " & password
                        Exit Sub
                    End If
                Next x4
                MsgBox password
            Next x3
        Next x2
    Next x1

End Sub



