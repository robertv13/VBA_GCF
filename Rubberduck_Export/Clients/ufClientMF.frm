VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufClientMF 
   Caption         =   "Gestion du fichier Clients (version 5.4)"
   ClientHeight    =   12030
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

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:cmbSearchColumn_Change", "", 0)
    
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
    
    Call CM_Log_Activities("ufClientMF:cmbSearchColumn_Change", Me.cmbSearchColumn.Value, startTime)

End Sub

Private Sub cmdAddClient_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:cmdAddClient_Click", "", 0)
    
    Call CM_Reset_UserForm
    
    ufClientMF.txtCodeClient.Enabled = True
    ufClientMF.txtCodeClient.SetFocus
    
    'Obtenir le dernier num�ro de client utilis� (particulier / corporatif)
    Dim maxSmallCode As String, maxLargeCode As String
    Call Max_Code_Values_From_GCF_Entree(maxSmallCode, maxLargeCode)
    
    MsgBox "Code � utiliser pour les particuliers = '" & maxSmallCode & "'" & vbNewLine & vbNewLine & _
           "Code � utiliser pour les entreprises  = '" & maxLargeCode & "'", vbInformation, _
                                        "Codes � utiliser pour la cr�ation d'un nouveau client"

    Call CM_Log_Activities("ufClientMF:cmdAddClient_Click", "", startTime)

End Sub

Private Sub cmdDelete_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:cmdDelete_Click", "", 0)
    
    Dim clientUtilise As Boolean
    Call Valider_Client_Avant_Effacement(Me.txtCodeClient.Value, clientUtilise)
    If clientUtilise = True Then
        MsgBox "Ce client est utilis� dans au moins une table de donn�es" & _
                vbNewLine & vbNewLine & "Il est donc impossible de d�truire ce client", _
                vbInformation, "Code de client est utilis� dans les tables"
        GoTo Clean_Exit
    End If
    
    Call Delete_Client(Me.txtCodeClient.Value)
            
    Call CM_Reset_UserForm
        
    ufClientMF.txtCodeClient.Enabled = False
    
Clean_Exit:
    
    ufClientMF.cmdDelete = False
    
    Call CM_Log_Activities("ufClientMF:cmdDelete_Click", "", startTime)

End Sub

Private Sub cmdCancel_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:cmdCancel_Click", "", 0)
    
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("D�sirez-vous vraiment ANNULER la pr�sente modification ?", vbYesNo + vbInformation, "Annuler les modifications courantes")
    If msgValue = vbNo Then GoTo CleanExit

    Call CM_Reset_UserForm
    
    ufClientMF.cmdAddClient.Enabled = True
    ufClientMF.cmdCancel.Enabled = False
    ufClientMF.cmdSave.Enabled = False

CleanExit:

    Call CM_Log_Activities("ufClientMF:cmdCancel_Click", CStr(msgValue), startTime)

End Sub

Private Sub cmdEdit_Click()
    
    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:cmdEdit_Click", "", 0)
    
    If Fn_Selected_List = 0 Then
        MsgBox "Aucun enregistrement n'a �t� choisi.", vbOKOnly + vbInformation, "Modification"
        GoTo Clean_Exit
    End If
    
    'Save selected line number
    wshMENU.Range("B100").Value = Fn_Selected_List
    
    'Code to update the value to respective controls - .ListIndex is based '0'
    Me.txtRowNumber.Value = Application.WorksheetFunction.Match(Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 0), _
        ThisWorkbook.Sheets("Donn�es").Range("A:A"), 0)
    Me.txtNomClient.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 0)
    Me.txtCodeClient.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 1)
    Me.txtNomClientSysteme.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 2)
    Me.txtContactFact.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 3)
    Me.txtTitreContact.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 4)
    Me.txtCourrielFact.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 5)
    Me.txtAdresse1.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 6)
    Me.txtAdresse2.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 7)
    Me.txtVille.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 8)
    Me.txtProvince.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 9)
    Me.txtCodePostal.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 10)
    Me.txtPays.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 11)
    Me.txtReferePar.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 12)
    Me.txtFinAnnee.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 13)
    Me.txtComptable.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 14)
    Me.txtNotaireAvocat.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 15)
    Me.txtNomClientPlusNomClientSyst�me.Value = Me.lstDonn�es.List(Me.lstDonn�es.ListIndex, 16)
    
    nouveauClient = False
    
    ufClientMF.cmdEdit.Enabled = False
    ufClientMF.cmdSave.Enabled = True
    ufClientMF.cmdCancel.Enabled = True
    
Clean_Exit:

    Call CM_Log_Activities("ufClientMF:cmdEdit_Click", Me.txtNomClient.Value, startTime)

End Sub

Private Sub cmdSave_Click()

    Call MAJ_Fichier_Client

End Sub

Private Sub MAJ_Fichier_Client()
    
    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:MAJ_Fichier_Client", "", 0)
    
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
    clientExists = Fn_Is_Client_Code_Already_Used
    
    If clientExists = True Then
        Call CM_Update_External_GCF_Entr�e_BD("UPDATE")
        Call CM_Update_Locally_GCF_BD_Entr�e("UPDATE")
    Else
        Call CM_Update_External_GCF_Entr�e_BD("NEW_RECORD")
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

Clean_Exit:

    Call CM_Log_Activities("ufClientMF:MAJ_Fichier_Client", msgValueLog, startTime)

End Sub

Private Sub Delete_Client(clientID)
    
    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:Delete_Client", "", 0)
    
    'D�finir le nom du fichier en fonction de l'utilisateur
    Dim targetFileName As String
    If Not Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        targetFileName = "P:\Administration\APP\GCF\DataFiles\GCF_BD_Entr�e.xlsx"
    Else
        targetFileName = "C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Entr�e.xlsx"
    End If
    Dim targetWorksheet As String: targetWorksheet = "Clients"

    Dim wb As Workbook: Set wb = Workbooks.Open(targetFileName)
    Dim ws As Worksheet: Set ws = wb.Sheets(targetWorksheet)

    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("D�sirez-vous vraiment D�TRUIRE le pr�sent client ?", vbYesNo + vbInformation, "D�truire le client de fa�on PERMANENTE")
    If msgValue = vbYes Then
        
        Dim foundCell As Range
        Set foundCell = ws.Cells.Find(What:=clientID, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            ws.Rows(foundCell.Row).Delete
        Else
            MsgBox "Le client '" & clientID & "' ne peut �tre trouv� dans Clients", vbCritical
            msgValue = vbNo
        End If
        
        'Onglet Donn�es
        Set ws = wshClients
        Set foundCell = ws.Cells.Find(What:=clientID, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            ws.Rows(foundCell.Row).Delete
        Else
            MsgBox "Le client '" & clientID & "' ne peut �tre trouv� dans Donn�es", vbCritical
            msgValue = vbNo
        End If
        
        'Onglet Donn�esRecherche
        Set ws = wshSearchData
        Set foundCell = ws.Cells.Find(What:=clientID, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            ws.Rows(foundCell.Row).Delete
        Else
            'Pas n�cessairement dans l'onglet 'Donn�esRecherche'
        End If
        
        MsgBox "Le client '" & Me.txtCodeClient.Value & "' a �t� d�truit" & vbNewLine & _
                vbNewLine & "de fa�on PERMANENTE", vbInformation
    End If
    
    Me.txtSearch.Value = ""
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
    
    'Ferme ET sauvegarde le fichier Excel
    If msgValue = vbYes Then
        wb.Close SaveChanges:=True
    Else
        wb.Close SaveChanges:=False
    End If

Clean_Exit:

    Call CM_Log_Activities("ufClientMF:Delete_Client", "", startTime)

End Sub

Private Sub cmdSearch_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:cmdSearch_Click", "", 0)
    
    If Me.txtSearch.Value <> "" Then
        Call CM_Build_Donn�es_Recherche
    Else
        MsgBox "SVP, saisir la valeur � rechercher.", vbOKOnly + vbInformation, "Recherche"
    End If

    Call CM_Log_Activities("ufClientMF:cmdSearch_Click", "", startTime)

End Sub

Private Sub Fix_Some_Fields()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:Fix_Some_Fields", "", 0)
    
    'Add the contact name to the client's name within square brackets
    If InStr(ufClientMF.txtNomClient.Value, "[") = 0 And _
        InStr(ufClientMF.txtNomClient.Value, "]") = 0 And _
            InStr(ufClientMF.txtNomClient.Value, ufClientMF.txtContactFact.Value) = 0 Then
                ufClientMF.txtNomClient.Value = Trim(ufClientMF.txtNomClient.Value) & " [" & Trim(ufClientMF.txtContactFact.Value) & "]"
    End If
    
    'Construire la cellule NomClientPlusNomClientSyst�me
    Dim nomClient As String, nomClientSyst�me As String, NomClientPlusNomClientSyst�me As String
    
    nomClient = Trim(ufClientMF.txtNomClient)
    
    nomClientSyst�me = Trim(ufClientMF.txtNomClientSysteme)
    nomClientSyst�me = Replace(nomClientSyst�me, "<", "")
    nomClientSyst�me = Replace(nomClientSyst�me, ">", "")
    nomClientSyst�me = Replace(nomClientSyst�me, "(", "")
    nomClientSyst�me = Replace(nomClientSyst�me, ")", "")
    nomClientSyst�me = Replace(nomClientSyst�me, ",", "")
    nomClientSyst�me = Replace(nomClientSyst�me, "CGA", "")
    nomClientSyst�me = Replace(nomClientSyst�me, "CA", "")
    nomClientSyst�me = Replace(nomClientSyst�me, "CGA", "")
    nomClientSyst�me = Replace(nomClientSyst�me, "CPA", "")
        
    NomClientPlusNomClientSyst�me = nomClient + " "
    
    Dim mots() As String
    mots = Split(nomClientSyst�me, " ")
    
    'S'il manque des mots dans NomClientPlusNomClientSyst�me, on ajoute ces mots un � un
    Dim m As Integer
    If UBound(mots, 1) > 0 Then
        For m = 0 To UBound(mots, 1)
            If InStr(UCase(NomClientPlusNomClientSyst�me), UCase(mots(m))) = 0 Then
                If InStr("AU~DES~DU~ET~FILS~INC~LA~LE~", UCase(mots(m))) = 0 Then
                    NomClientPlusNomClientSyst�me = NomClientPlusNomClientSyst�me & mots(m) & " "
                End If
            End If
        Next m
    End If
    
    NomClientPlusNomClientSyst�me = Trim(NomClientPlusNomClientSyst�me)

    ufClientMF.txtNomClientPlusNomClientSyst�me.Value = NomClientPlusNomClientSyst�me
    
    Call CM_Log_Activities("ufClientMF:Fix_Some_Fields", "", startTime)

End Sub

Private Sub lstDonn�es_Click()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:lstDonn�es_Click", "", 0)
    
    ufClientMF.cmdAddClient.Enabled = True
    ufClientMF.cmdEdit.Enabled = True

    wshMENU.Range("B4").Value = Me.lstDonn�es.ListIndex
    
    Call CM_Log_Activities("ufClientMF:lstDonn�es_Click", "", startTime)

End Sub

Private Sub lstDonn�es_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:lstDonn�es_DblClick", "", 0)
    
    Me.cmdDelete.Enabled = True
    Me.cmdEdit.Enabled = False
    Me.cmdAddClient.Enabled = False
    
    nouveauClient = False
    
    wshMENU.Range("B4").Value = Me.lstDonn�es.ListIndex
    
    Call cmdEdit_Click

    Call CM_Log_Activities("ufClientMF:lstDonn�es_DblClick", "", startTime)

End Sub

Private Sub txtCodeClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:txtCodeClient_Exit", "", 0)
    
    Dim clientExists As Boolean
    clientExists = Fn_Is_Client_Code_Already_Used
    
    If clientExists = True And nouveauClient = True Then
        ufClientMF.txtCodeClient.BackColor = vbRed
        MsgBox "Ce code de client '" & ufClientMF.txtCodeClient.Value & "' existe d�j� en base de donn�es." & vbNewLine & vbNewLine & _
               "Veuillez choisir un AUTRE code qui n'existe pas, SVP", vbCritical + vbOKOnly, "Doublon de code de client"
        ufClientMF.txtCodeClient.BackColor = vbWhite
        ufClientMF.txtCodeClient.Value = ""
        ufClientMF.txtCodeClient.SetFocus
    End If
    
    Call CM_Log_Activities("ufClientMF:txtCodeClient_Exit", ufClientMF.txtCodeClient.Value, startTime)

End Sub

Private Sub txtNomClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:txtNomClient_Exit", "", 0)
    
    If Trim(ufClientMF.txtNomClient.Value) <> "" Then
        ufClientMF.cmdSave.Enabled = True
    End If
    
    ufClientMF.cmdCancel.Enabled = True

    Call CM_Log_Activities("ufClientMF:txtNomClient_Exit", ufClientMF.txtNomClient.Value, startTime)

End Sub

Private Sub cmbFinAnnee_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    ufClientMF.txtFinAnnee.Value = Fn_Fix_Txt_Fin_Annee(ufClientMF.cmbFinAnnee.Value)

End Sub

Private Sub txtSearch_Change()

    ufClientMF.cmdSearch.Enabled = True

End Sub

Private Sub UserForm_Initialize()

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("ufClientMF:UserForm_Initialize", "", 0)
    
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
    
    ufClientMF.cmdDelete.Enabled = False
    ufClientMF.cmdSave.Enabled = False
    ufClientMF.cmdCancel.Enabled = False

    Call CM_Log_Activities("ufClientMF:UserForm_Initialize", "", startTime)

End Sub

