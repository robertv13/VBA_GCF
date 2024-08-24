VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufClientMF 
   Caption         =   "Gestion du fichier Clients (version 2.2)"
   ClientHeight    =   11415
   ClientLeft      =   6615
   ClientTop       =   2460
   ClientWidth     =   18270
   OleObjectBlob   =   "ufClientMF.frx":0000
End
Attribute VB_Name = "ufClientMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EnableEvents As Boolean
Public nouveauClient As Boolean
Public isActiveSearch As Boolean

Private Sub cmbSearchColumn_Change()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:cmbSearchColumn_Change", "", 0)
    
    If Me.EnableEvents = False Then GoTo Clean_Exit
    
    If Me.cmbSearchColumn.Value = "Tous" Then
        Call CM_Reset_UserForm
    Else
        Me.txtSearch.Value = ""
        Me.txtSearch.Enabled = True
        Me.cmdSearch.Enabled = True
    End If

Clean_Exit:

    Me.lblResultCount = ""
    
    Call CM_Log_Record("ufClientMF:cmbSearchColumn_Change", Me.cmbSearchColumn.Value, startTime)

End Sub

Private Sub cmdAddClient_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:cmdAddClient_Click", "", 0)
    
    Call CM_Reset_UserForm
    
    ufClientMF.txtCodeClient.Enabled = True
    ufClientMF.txtCodeClient.SetFocus
    
    nouveauClient = True

    Call CM_Log_Record("ufClientMF:cmdAddClient_Click", "", startTime)

End Sub

Private Sub cmdCancel_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:cmdCancel_Click", "", 0)
    
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("D�sirez-vous vraiment ANNULER la pr�sente modification ?", vbYesNo + vbInformation, "Annuler les modifications courantes")
    If msgValue = vbNo Then GoTo CleanExit

    Call CM_Reset_UserForm
    
    ufClientMF.cmdAddClient.Enabled = True
    ufClientMF.cmdCancel.Enabled = False
    ufClientMF.cmdSave.Enabled = False

CleanExit:

    Call CM_Log_Record("ufClientMF:cmdCancel_Click", CStr(msgValue), startTime)

End Sub

Private Sub cmdEdit_Click()
    
    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:cmdEdit_Click", "", 0)
    
    If Fn_Selected_List = 0 Then
        MsgBox "Aucun enregistrement n'a �t� choisi.", vbOKOnly + vbInformation, "Modification"
        GoTo Clean_Exit
    End If
    
    'Save selected line number
    wshMENU.Range("B100").Value = Fn_Selected_List
    
    'Code to update the value to respective controls
    Me.txtRowNumber.Value = Application.WorksheetFunction.Match(Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 0), _
        ThisWorkbook.Sheets("Donn�es").Range("A:A"), 0)
    Me.txtNomClient.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 0)
    Me.txtCodeClient.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 1)
    Me.txtContactFact.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 2)
    Me.txtTitreContact.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 3)
    Me.txtCourrielFact.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 4)
    Me.txtAdresse1.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 5)
    Me.txtAdresse2.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 6)
    Me.txtVille.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 7)
    Me.txtProvince.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 8)
    Me.txtCodePostal.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 9)
    Me.txtPays.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 10)
    Me.txtReferePar.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 11)
    Me.txtFinAnnee.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 12)
    Me.txtComptable.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 13)
    Me.txtNotaireAvocat.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 14)
    
    nouveauClient = False
    
    ufClientMF.cmdEdit.Enabled = False
    ufClientMF.cmdSave.Enabled = True
    ufClientMF.cmdCancel.Enabled = True
    
Clean_Exit:

    Call CM_Log_Record("ufClientMF:cmdEdit_Click", Me.txtNomClient.Value, startTime)

End Sub

Private Sub cmdSave_Click()
    
    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:cmdSave_Click", "", 0)
    
    If Fn_ValidateEntries() = False Then
        GoTo Clean_Exit
    End If
   
    Call Fix_Some_Fields '2024-08-10 @ 08:36
    
    'Confirm the Update
    Dim msgValue As VbMsgBoxResult, msgValueLog As String
    msgValue = MsgBox("D�sirez-vous SAUVEGARDER ces informations ?", vbYesNo + vbInformation, "Confirmation avant la sauvegarde")
    msgValueLog = msgValue
    If msgValue = vbNo Then
        GoTo Clean_Exit
    End If
    
    Dim clientExists As Boolean
    clientExists = Fn_Does_Client_Code_Exist
    
    If clientExists = True Then
        Call CM_Update_External_GCF_BD_Entree("UPDATE")
        Call CM_Update_Locally_GCF_BD_Entr�e("UPDATE")
    Else
        Call CM_Update_External_GCF_BD_Entree("NEW_RECORD")
        Call CM_Update_Locally_GCF_BD_Entr�e("NEW_RECORD")
    End If
    
    Call CM_Reset_UserForm
    
    ufClientMF.cmdSave.Enabled = False
    ufClientMF.cmdCancel.Enabled = False
    
    If wshMENU.Range("B4").Value >= 0 And _
        wshMENU.Range("B4").Value < ufClientMF.lstDonn�es.ListCount Then
            ufClientMF.lstDonn�es.ListIndex = wshMENU.Range("B4").Value
            If wshMENU.Range("B4").Value > 15 Then
                ufClientMF.lstDonn�es.TopIndex = wshMENU.Range("B4").Value - 8
            End If
    End If
    
'    Me.txtSearch.SetFocus

Clean_Exit:

    Call CM_Log_Record("ufClientMF:cmdSave_Click", msgValueLog, startTime)

End Sub

Private Sub cmdSearch_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:cmdSearch_Click", "", 0)
    
    If Me.txtSearch.Value <> "" Then
        Call CM_Build_Donn�es_Recherche
    Else
        MsgBox "SVP, saisir la valeur � rechercher.", vbOKOnly + vbInformation, "Recherche"
    End If

    Call CM_Log_Record("ufClientMF:cmdSearch_Click", "", startTime)

End Sub

Private Sub Fix_Some_Fields()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:Fix_Some_Fields", "", 0)
    
    'Add the contact name to the client's name within square brackets
    If InStr(ufClientMF.txtNomClient.Value, "[") = 0 And _
        InStr(ufClientMF.txtNomClient.Value, "]") = 0 And _
            InStr(ufClientMF.txtNomClient.Value, ufClientMF.txtContactFact.Value) = 0 Then
                ufClientMF.txtNomClient.Value = Trim(ufClientMF.txtNomClient.Value) & " [" & Trim(ufClientMF.txtContactFact.Value) & "]"
    End If
    
    Call CM_Log_Record("ufClientMF:Fix_Some_Fields", "", startTime)

End Sub

Private Sub lstDonn�es_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:lstDonn�es_Click", "", 0)
    
    ufClientMF.cmdAddClient.Enabled = True
    ufClientMF.cmdEdit.Enabled = True

    wshMENU.Range("B4").Value = Me.lstDonn�es.ListIndex
    
    Call CM_Log_Record("ufClientMF:lstDonn�es_Click", "", startTime)

End Sub

Private Sub lstDonn�es_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:lstDonn�es_DblClick", "", 0)
    
    Me.cmdEdit.Enabled = False
    Me.cmdAddClient.Enabled = False
    
    nouveauClient = False
    
    wshMENU.Range("B4").Value = Me.lstDonn�es.ListIndex
    
    Call cmdEdit_Click

    Call CM_Log_Record("ufClientMF:lstDonn�es_DblClick", "", startTime)

End Sub

Private Sub txtCodeClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:txtCodeClient_Exit", "", 0)
    
    Dim clientExists As Boolean
    clientExists = Fn_Does_Client_Code_Exist
    
'    MsgBox "X - " & clientExists & nouveauClient '2024-08-22 @ 07:41
    
    If clientExists = True And nouveauClient = True Then
        ufClientMF.txtCodeClient.BackColor = vbRed
        MsgBox "Ce code de client '" & ufClientMF.txtCodeClient.Value & "' existe d�j� en base de donn�es." & vbNewLine & vbNewLine & _
               "Veuillez choisir un AUTRE code qui n'existe pas, SVP", vbCritical + vbOKOnly, "Doublon de code de client"
        ufClientMF.txtCodeClient.BackColor = vbWhite
        ufClientMF.txtCodeClient.Value = ""
        ufClientMF.txtCodeClient.SetFocus
    End If
    
    Call CM_Log_Record("ufClientMF:txtCodeClient_Exit", ufClientMF.txtCodeClient.Value, startTime)

End Sub

Private Sub txtNomClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:txtNomClient_Exit", "", 0)
    
    If Trim(ufClientMF.txtNomClient.Value) <> "" Then
        ufClientMF.cmdSave.Enabled = True
    End If
    
    ufClientMF.cmdCancel.Enabled = True

    Call CM_Log_Record("ufClientMF:txtNomClient_Exit", ufClientMF.txtNomClient.Value, startTime)

End Sub

Private Sub cmbFinAnnee_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    ufClientMF.txtFinAnnee.Value = Fn_Fix_Txt_Fin_Annee(ufClientMF.cmbFinAnnee.Value)

End Sub

Private Sub txtSearch_Change()

    ufClientMF.cmdSearch.Enabled = True

End Sub

Private Sub UserForm_Initialize()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Record("ufClientMF:UserForm_Initialize", "", 0)
    
    Call CM_Reset_UserForm
    
    With ufClientMF.cmbFinAnnee
        .AddItem "Janvier"
        .AddItem "F�vrier"
        .AddItem "Mars"
        .AddItem "Avril"
        .AddItem "Mai"
        .AddItem "Juin"
        .AddItem "Juillet"
        .AddItem "Ao�t"
        .AddItem "Septembre"
        .AddItem "Octobre"
        .AddItem "Novembre"
        .AddItem "D�cembre"
    End With
    
    ufClientMF.cmdSave.Enabled = False
    ufClientMF.cmdCancel.Enabled = False

    Call CM_Log_Record("ufClientMF:UserForm_Initialize", "", startTime)

End Sub
