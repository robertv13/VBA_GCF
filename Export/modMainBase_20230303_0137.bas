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
    frm.ListData = shClients.Range("A1").CurrentRegion
    frm.show
    
    ' Display the book that was selected
    If frm.Cancelled = True Then
        MsgBox "Vous quittez la recherche sans avoir trouvé de client"
    Else
        MsgBox "Le client trouvé est " & Chr(10) & Chr(13) & _
                                Chr(10) & Chr(13) & "'" & frm.Book & "'"
    End If
    
    ' Clean up
    Unload frm
    Set frm = Nothing

End Sub

