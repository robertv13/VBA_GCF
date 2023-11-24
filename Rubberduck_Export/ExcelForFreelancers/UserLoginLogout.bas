Attribute VB_Name = "UserLoginLogout"
Option Explicit

Sub UserLogin()
    Sheet2.Range("B4, B5, B7").ClearContents
    LoginForm.Username.Value = ""
    LoginForm.Password.Value = ""
    LoginForm.Show
End Sub

Sub UserLogout()
    Dim WkSht As Worksheet
    Sheets("Login").Activate
    For Each WkSht In ThisWorkbook.Worksheets
        If WkSht.Name <> "Login" Then WkSht.Visible = xlSheetVeryHidden
    Next WkSht
End Sub
