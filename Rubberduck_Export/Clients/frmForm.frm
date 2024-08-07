VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmForm 
   Caption         =   "Gestion du fichier Clients"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13095
   OleObjectBlob   =   "frmForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EnableEvents As Boolean

Private Sub cmbSearchColumn_Change()

    If Me.EnableEvents = False Then Exit Sub
    
    If Me.cmbSearchColumn.Value = "Tous" Then
        Call Reset
    Else
        Me.txtSearch.Value = ""
        Me.txtSearch.Enabled = True
        Me.cmdSearch.Enabled = True
    End If

End Sub

Private Sub cmdDelete_Click()
    
    If Selected_List = 0 Then
        MsgBox "Aucun client n'a été choisi.", vbOKOnly + vbInformation, "Destruction"
        Exit Sub
    End If
    
    Dim i As VbMsgBoxResult
    i = MsgBox("Désirez-vous DÉTRUIRE ce client ?", vbYesNo + vbQuestion, "Confirmation")
    If i = vbNo Then Exit Sub
    
    Dim iRow As Long
    iRow = Application.WorksheetFunction.Match(Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0), _
    ThisWorkbook.Sheets("Database").Range("A:A"), 0)
    
    ThisWorkbook.Sheets("Database").Rows(iRow).Delete
    
    Call Reset
    
    MsgBox "Le client a été DÉTRUIT.", vbOKOnly + vbInformation, "Deleted"
    
End Sub

Private Sub cmdEdit_Click()
    
    If Selected_List = 0 Then
        MsgBox "Aucun client n'a été choisi.", vbOKOnly + vbInformation, "Modification"
        Exit Sub
    End If
    
    'Code to update the value to respective controls
    Dim sGender As String
    Me.txtRowNumber.Value = Application.WorksheetFunction.Match(Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0), _
    ThisWorkbook.Sheets("Database").Range("A:A"), 0)
    Me.txtID.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 1)
    Me.txtName.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 2)
    sGender = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 3)
    If sGender = "Female" Then
        Me.optFemale.Value = True
    Else
        Me.optMale.Value = True
    End If
    Me.cmbDepartment.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 4)
    Me.txtCity.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 5)
    Me.txtCountry.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 6)
    
    MsgBox "Veuillez apporter les changements et cliquer sur 'Sauvegarde' pour enregistrer.", vbOKOnly + vbInformation, "Changement"
    
End Sub

Private Sub cmdReset_Click()

    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Désirez-vous VIDER le formulaire ?", vbYesNo + vbInformation, "Confirmation")
    If msgValue = vbNo Then Exit Sub
    
    Call Reset

End Sub

Private Sub cmdSave_Click()
    
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Désirez-vous SAUVEGARDER ces informations ?", vbYesNo + vbInformation, "Confirmation")
    If msgValue = vbNo Then Exit Sub
    
    If ValidateEntries() = True Then
    
        Call Submit
        Call Reset
    
    End If
    
End Sub

Private Sub cmdSearch_Click()

    If Me.txtSearch.Value = "" Then
        MsgBox "SVP, saisir la valeur à rechercher.", vbOKOnly + vbInformation, "Recherche"
        Exit Sub
    End If
    
    Call DonnéesRecherche
    
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub UserForm_Initialize()

    Call Reset

End Sub
