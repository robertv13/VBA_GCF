VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmForm 
   Caption         =   "Gestion du fichier Clients"
   ClientHeight    =   11160
   ClientLeft      =   7125
   ClientTop       =   2970
   ClientWidth     =   17955
   OleObjectBlob   =   "frmForm.frx":0000
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

'Private Sub cmdDelete_Click()
'
'    If Selected_List = 0 Then
'        MsgBox "Aucun client n'a été choisi.", vbOKOnly + vbInformation, "Destruction"
'        Exit Sub
'    End If
'
'    Dim i As VbMsgBoxResult
'    i = MsgBox("Désirez-vous DÉTRUIRE ce client ?", vbYesNo + vbQuestion, "Confirmation")
'    If i = vbNo Then Exit Sub
'
'    Dim iRow As Long
'    iRow = Application.WorksheetFunction.Match(Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0), _
'    ThisWorkbook.Sheets("Données").Range("A:A"), 0)
'
'    ThisWorkbook.Sheets("Données").Rows(iRow).Delete
'
'    Call Reset
'
'    MsgBox "Le client a été DÉTRUIT.", vbOKOnly + vbInformation, "Deleted"
'
'End Sub
'
Private Sub cmdEdit_Click()
    
    If Selected_List = 0 Then
        MsgBox "Aucun enregistrement n'a été choisi.", vbOKOnly + vbInformation, "Modification"
        Exit Sub
    End If
    
    'Code to update the value to respective controls
    Me.txtRowNumber.Value = Application.WorksheetFunction.Match(Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0), _
        ThisWorkbook.Sheets("Données").Range("A:A"), 0)
    Me.txtNomClient.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0)
    Me.txtCodeClient.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 1)
    Me.txtContactFact.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 2)
    Me.txtTitreContact.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 3)
    Me.txtCourrielFact.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 4)
    Me.txtAdresse1.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 5)
    Me.txtAdresse2.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 6)
    Me.txtVille.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 7)
    Me.txtProvince.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 8)
    Me.txtCodePostal.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 9)
    Me.txtPays.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 10)
    Me.txtReferePar.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 11)
    Me.txtFinAnnee.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 12)
    Me.txtComptable.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 13)
    Me.txtNotaireAvocat.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 14)
    
'    MsgBox "Veuillez apporter les changements et cliquer sur 'Sauvegarde' pour enregistrer.", vbOKOnly + vbInformation, "Changement"
    
    frmForm.cmdSave.Enabled = True
    frmForm.cmdCancel.Enabled = True
    
End Sub

Private Sub cmdCancel_Click()

    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Désirez-vous vraiment ANNULER la présente modification ?", vbYesNo + vbInformation, "Annuler les modifications courantes")
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
        
        frmForm.cmdSave.Enabled = False
        frmForm.cmdCancel.Enabled = False
    
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

Private Sub lstDatabase_Click()

    frmForm.cmdEdit.Enabled = True

End Sub

Private Sub lstDatabase_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call cmdEdit_Click

End Sub

Private Sub UserForm_Initialize()

    Call Reset
    
    frmForm.cmdSave.Enabled = False
    frmForm.cmdCancel.Enabled = False

End Sub
