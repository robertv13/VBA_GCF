Attribute VB_Name = "modMain"
Option Explicit

'Author: Paul Kelly - https://ExcelMacroMastery.com/

'Displays the UserForm and retrieves a value
Public Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "w\n14"

    'Create the UserForm
    Dim frm As UserFormCompany
    Set frm = UserForms.Add(UserFormCompany.Name)
    
    'Set the range
    frm.ListData = shCompanies.Range("A1").CurrentRegion
    frm.show
    
    'Display the company that was selected
    If frm.Cancelled = True Then
        MsgBox "Vous n'avez choisi aucune compagnie dans votre recherche"
    Else
        MsgBox "Vous avez trouvé la compagnie suivante '" & frm.Company & "'"
    End If
    
    'Clean up
    Unload frm
    Set frm = Nothing

End Sub
