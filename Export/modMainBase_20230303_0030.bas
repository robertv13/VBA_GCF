Attribute VB_Name = "modMain"
Option Explicit

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
' Displays the UserForm and retrieves a value
Public Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "w\n14"

    ' Create the UserForm
    Dim frm As UserFormBook
    Set frm = UserForms.Add(UserFormBook.Name)
    
    ' Set the range
    frm.ListData = shBooks.Range("A1").CurrentRegion
    frm.show
    
    ' Display the book that was selected
    If frm.Cancelled = True Then
        MsgBox "You have cancelled the form. No book was selected"
    Else
        MsgBox "The book you selected was: " & frm.Book
    End If
    
    ' Clean up
    Unload frm
    Set frm = Nothing

End Sub

