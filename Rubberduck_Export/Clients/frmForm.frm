VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmForm 
   Caption         =   "Gestion du fichier Clients"
   ClientHeight    =   11160
   ClientLeft      =   6615
   ClientTop       =   2460
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
Public nouveauClient As Boolean

Private Sub cmbFinAnnee_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    frmForm.txtFinAnnee.Value = Fix_Txt_Fin_Annee(frmForm.cmbFinAnnee.Value)

End Sub

Private Sub cmbSearchColumn_Change()

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:cmbSearchColumn_Change", 0)
    
    If Me.EnableEvents = False Then Exit Sub
    
    If Me.cmbSearchColumn.Value = "Tous" Then
        Call Reset
    Else
        Me.txtSearch.Value = ""
        Me.txtSearch.Enabled = True
        Me.cmdSearch.Enabled = True
    End If

    Call Log_Record("frmForm:cmbSearchColumn_Change", startTime)

End Sub

Private Sub cmdAddClient_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:cmdAddClient_Click", 0)
    
    Call Reset
    
    frmForm.txtCodeClient.Enabled = True
    frmForm.txtCodeClient.SetFocus
    
    nouveauClient = True

    Call Log_Record("frmForm:cmdAddClient_Click", startTime)

End Sub

Private Sub cmdCancel_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:cmdCancel_Click", 0)
    
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Désirez-vous vraiment ANNULER la présente modification ?", vbYesNo + vbInformation, "Annuler les modifications courantes")
    If msgValue = vbNo Then Exit Sub

    Call Reset
    
    Me.cmdAddClient.Enabled = True
    Me.cmdCancel.Enabled = False
    Me.cmdSave.Enabled = False

    Call Log_Record("frmForm:cmdCancel_Click", startTime)

End Sub

Private Sub cmdEdit_Click()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:cmdEdit_Click", 0)
    
    If Selected_List = 0 Then
        MsgBox "Aucun enregistrement n'a été choisi.", vbOKOnly + vbInformation, "Modification"
        Exit Sub
    End If
    
    'Save selected line number
    wshMENU.Range("B3").Value = Selected_List
    
    'Code to update the value to respective controls
    Me.txtRowNumber.Value = Application.WorksheetFunction.Match(Me.lstDonnées.List(Me.lstDonnées.ListIndex, 0), _
        ThisWorkbook.Sheets("Données").Range("A:A"), 0)
    Me.txtNomClient.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 0)
    Me.txtCodeClient.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 1)
    Me.txtContactFact.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 2)
    Me.txtTitreContact.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 3)
    Me.txtCourrielFact.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 4)
    Me.txtAdresse1.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 5)
    Me.txtAdresse2.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 6)
    Me.txtVille.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 7)
    Me.txtProvince.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 8)
    Me.txtCodePostal.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 9)
    Me.txtPays.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 10)
    Me.txtReferePar.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 11)
    Me.txtFinAnnee.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 12)
    Me.txtComptable.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 13)
    Me.txtNotaireAvocat.Value = Me.lstDonnées.List(Me.lstDonnées.ListIndex, 14)
    
    nouveauClient = False
    
    frmForm.cmdEdit.Enabled = False
    frmForm.cmdSave.Enabled = True
    frmForm.cmdCancel.Enabled = True
    
    Call Log_Record("frmForm:cmdEdit_Click", startTime)

End Sub

Private Sub cmdSave_Click()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:cmdSave_Click", 0)
    
    If ValidateEntries() = False Then
        Exit Sub
    End If
   
    Call Fix_Some_Fields '2024-08-10 @ 08:36
    
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Désirez-vous SAUVEGARDER ces informations ?", vbYesNo + vbInformation, "Confirmation")
    If msgValue = vbNo Then Exit Sub
    
    Dim clientExists As Boolean
    clientExists = Fn_Does_Client_Code_Exist
    
    If clientExists = True Then
        Call Submit_GCF_BD_Entrée_Clients("UPDATE")
        Call Submit_Locally("UPDATE")
    Else
        Call Submit_GCF_BD_Entrée_Clients("NEW_RECORD")
        Call Submit_Locally("NEW_RECORD")
    End If
    
    Call Reset
    
    frmForm.cmdSave.Enabled = False
    frmForm.cmdCancel.Enabled = False
    
    If wshMENU.Range("B3").Value >= 0 And _
        wshMENU.Range("B3").Value < frmForm.lstDonnées.ListCount Then
            frmForm.lstDonnées.ListIndex = wshMENU.Range("B3").Value
            If wshMENU.Range("B3").Value > 15 Then
                frmForm.lstDonnées.TopIndex = wshMENU.Range("B3").Value - 8 'Guess ?
            End If
    End If
    
'    Me.txtSearch.SetFocus

    Call Log_Record("frmForm:cmdSave_Click", startTime)

End Sub

Private Sub cmdSearch_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:cmdSearch_Click", 0)
    
    If Me.txtSearch.Value = "" Then
        MsgBox "SVP, saisir la valeur à rechercher.", vbOKOnly + vbInformation, "Recherche"
        Exit Sub
    End If
    
    Call DonnéesRecherche
    
    Call Log_Record("frmForm:cmdSearch_Click", startTime)

End Sub

Private Sub Fix_Some_Fields()

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:Fix_Some_Fields", 0)
    
    'Add the contact name to the client's name within square brackets
    If InStr(frmForm.txtNomClient.Value, "[") = 0 And _
        InStr(frmForm.txtNomClient.Value, "]") = 0 And _
            InStr(frmForm.txtNomClient.Value, frmForm.txtContactFact.Value) = 0 Then
                frmForm.txtNomClient.Value = Trim(frmForm.txtNomClient.Value) & " [" & Trim(frmForm.txtContactFact.Value) & "]"
    End If
    
    Call Log_Record("frmForm:Fix_Some_Fields", startTime)

End Sub

Private Sub lstDonnées_Click()

    frmForm.cmdEdit.Enabled = True

End Sub

Private Sub lstDonnées_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:lstDonnées_DblClick", 0)
    
    Me.cmdEdit.Enabled = False
    Me.cmdAddClient.Enabled = False
    
    nouveauClient = False
    
    Call cmdEdit_Click

    Call Log_Record("frmForm:lstDonnées_DblClick", startTime)

End Sub

Private Sub txtCodeClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:txtCodeClient_Exit", 0)
    
    Dim clientExists As Boolean
    clientExists = Fn_Does_Client_Code_Exist
    
    If clientExists = True Then
        frmForm.txtCodeClient.BackColor = vbRed
        MsgBox "Ce code de client '" & frmForm.txtCodeClient.Value & "' existe déjà en base de données." & vbNewLine & vbNewLine & _
               "Veuillez choisir un AUTRE code qui n'existe pas, SVP", vbCritical + vbOKOnly, "Doublon de code de client"
        frmForm.txtCodeClient.BackColor = vbWhite
        frmForm.txtCodeClient.Value = ""
        frmForm.txtCodeClient.SetFocus
        
    End If
    
    Call Log_Record("frmForm:txtCodeClient_Exit", startTime)

End Sub

Private Sub txtNomClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:txtNomClient_Exit", 0)
    
    If Trim(frmForm.txtNomClient) <> "" Then
        frmForm.cmdSave.Enabled = True
    End If
    
    frmForm.cmdCancel.Enabled = True

    Call Log_Record("frmForm:txtNomClient_Exit", startTime)

End Sub

Private Sub txtSearch_Change()

    frmForm.cmdSearch.Enabled = True

End Sub

Private Sub UserForm_Initialize()

    Dim startTime As Double: startTime = Timer: Call Log_Record("frmFrom:UserForm_Initialize", 0)
    
    Call Reset
    
    With frmForm.cmbFinAnnee
        .AddItem "Janvier"
        .AddItem "Février"
        .AddItem "Mars"
        .AddItem "Avril"
        .AddItem "Mai"
        .AddItem "Juin"
        .AddItem "Juillet"
        .AddItem "Août"
        .AddItem "Septembre"
        .AddItem "Octobre"
        .AddItem "Novembre"
        .AddItem "Décembre"
    End With
    
    frmForm.cmdSave.Enabled = False
    frmForm.cmdCancel.Enabled = False

    Call Log_Record("frmForm:UserForm_Initialize", startTime)

End Sub
