VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    Dim WkSht As Worksheet
    If Sheet2.Range("B6").Value = True Then
        LoginForm.Hide
        Sheet2.Range("B7").Value = Sheet2.Range("B4").Value 'Set Current User
        For Each WkSht In ThisWorkbook.Worksheets
            If WkSht.Name = "Admin" Then
                If Sheet2.Range("B8").Value = "Yes" Then 'Admin
                    WkSht.Visible = xlSheetVisible
                Else
                    WkSht.Visible = xlSheetHidden
                End If
            Else 'Not Admin
                WkSht.Visible = xlSheetVisible
            End If
        Next WkSht
        Sheet2.Range("B4:B5").ClearContents
    Else
        MsgBox "Please enter valid Username and Password"
    End If
End Sub

Private Sub btnCancel_Click()
    LoginForm.Hide
    Username.Text = ""
    Password.Text = ""
    Sheet2.Range("B4:B5").ClearContents
End Sub

Private Sub Username_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Sheet2.Range("B4").Value = Me.Username.Value
End Sub

Private Sub Password_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Sheet2.Range("B5").Value = Me.Password.Value
End Sub

