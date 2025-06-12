VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   10308
   ClientLeft      =   192
   ClientTop       =   780
   ClientWidth     =   15984
   OleObjectBlob   =   "ufSaisieHeures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSaisieHeures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oEventHandler As New clsSearchableDropdown '2023-03-21 @ 09:16
'Private tracker As clsFormActivityTracker '2025-05-30 @ 12:49

'Allows the calling code to set the data
Public Property Let ListData(ByVal rg As Range)

    oEventHandler.List = rg.value

End Property

Private Sub UserForm_Initialize() '2025-05-30 @ 13:26

    Call ConnectFormControls(Me)
    Call VerifierEtatBoutonAjouter
    Call RafraichirActivite("Activité dans userForm '" & Me.Name & "'")
    
End Sub

Sub UserForm_Activate() '2024-07-31 @ 07:57

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:UserForm_Activate", "", 0)
    
    gLogSaisieHeuresVeryDetailed = False
    
    Call modImport.ImporterClients
    Call modImport.ImporterTEC
    
    'Mise en place de la colonne à rechercher dans BD_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wsdBD_Clients.Cells(wsdBD_Clients.Rows.count, 1).End(xlUp).row
    ufSaisieHeures.ListData = wsdBD_Clients.Range("Q1:Q" & lastUsedRow) '2025-01-11 @ 18:00
    
    With oEventHandler
        Set .SearchListBox = lstboxNomClient
        Set .SearchTextBox = txtClient
        .MaxRows = 10
        .ShowAllMatches = False
        .CompareMethod = vbTextCompare
    End With

    Call VerifierEtatBoutonAjouter '2025-06-09 @ 08:28
'    Call ActiverButtonsVraiOuFaux("UserFormActivate", False, False, False, False)

    ufSaisieHeures.txtDate.value = "" 'On vide la date pour forcer la saisie
    
    On Error Resume Next
    ufSaisieHeures.cmbProfessionnel.SetFocus
    On Error GoTo 0
   
    savedHeures = 0 '2025-05-07 @17:03
    
    rmv_state = rmv_modeInitial
    
    Call Log_Record("ufSaisieHeures:UserForm_Activate", "", startTime)
    
End Sub

Private Sub lstboxNomClient_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:lstboxNomClient_DblClick", "", 0)
    
    Dim i As Long
    With Me.lstboxNomClient
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Me.txtClient.value = .List(i, 0)
                Me.txtClientID.value = Fn_Cell_From_BD_Client(Me.txtClient.value, 17, 2)
                Me.txtClientReel.value = Fn_Cell_From_BD_Client(Me.txtClientID.value, 2, 1)
                Exit For
            End If
        Next i
    End With
    
    'Force à cacher le listbox pour les résultats de recherche
    On Error Resume Next
    Me.lstboxNomClient.Visible = False
    On Error GoTo 0
    
    Me.txtClient.TextAlign = fmTextAlignLeft

    Call Log_Record("ufSaisieHeures:lstboxNomClient_DblClick", "", startTime)

End Sub

Private Sub UserForm_Terminate()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:UserForm_Terminate", "", 0)

    'Libérer la mémoire
    Set oEventHandler = Nothing
    
    ufSaisieHeures.Hide
    Unload ufSaisieHeures
    
    If ufSaisieHeures.Name = "ufSaisieHeures" Then
        On Error GoTo MenuSelect
        wshMenuTEC.Select
        On Error GoTo 0
    Else
        wshMenu.Select
    End If
    
    GoTo Exit_Sub
    
MenuSelect:
    wshMenu.Activate
    wshMenu.Select
    
Exit_Sub:

    Call Log_Record("ufSaisieHeures:UserForm_Terminate", "", startTime)

End Sub

Private Sub cmbProfessionnel_Enter() '2025-05-31 @ 16:31

    Dim ws As Worksheet
    Dim plageInitiales As Range
    Dim cell As Range, cellInit As Range
    Dim listeInitiales As Collection
    Dim utilisateur As String
    Dim toutesInitiales As Boolean

    Set ws = wsdADMIN
    ' Plage de la table WindowsUser_Initials : colonnes D à F, lignes 63 à 78
    Set plageInitiales = ws.Range("D63:D78")
    
    utilisateur = GetNomUtilisateur() ' Variable globale utilisateur Windows

    Set listeInitiales = New Collection
    toutesInitiales = False
    
    ' Chercher utilisateur dans la liste et récupérer initiales permises
    For Each cell In plageInitiales
        If Trim(cell.value) <> "" Then
            If StrComp(cell.value, utilisateur, vbTextCompare) = 0 Then
                If Trim(cell.offset(0, 2).value) <> "" Then
                    'Initiales spécifiques autorisées pour cet utilisateur
                    listeInitiales.Add Trim(cell.offset(0, 2).value)
                Else
                    'Pas de restriction, on doit autoriser toutes les initiales
                    toutesInitiales = True
                End If
                Exit For ' Utilisateur trouvé, on peut sortir
            End If
        End If
    Next cell

    ' Si toutes les initiales sont permises, on ajoute toutes celles listées en colonne F
    If toutesInitiales Then
        'Ajoute GC qui est la valeur par défut
        listeInitiales.Add "GC", "GC"
        For Each cellInit In plageInitiales.offset(, 2).Resize(, 1)
            If Trim(cellInit.value) <> "" Then
                On Error Resume Next 'Eviter doublons dans la collection
                If Trim(cellInit.value) <> "Init. Permises" Then
'                    Debug.Print "XYZ - " & Trim(cellInit.value)
                    listeInitiales.Add Trim(cellInit.value), CStr(Trim(cellInit.value))
                End If
                On Error GoTo 0
            End If
        Next cellInit
    End If

    ' Remplir le ComboBox
    With Me.cmbProfessionnel
        .Clear
        Dim item As Variant
        For Each item In listeInitiales
            .AddItem item
        Next item
        ' Optionnel : sélection automatique de la première initiale
        If .ListCount > 0 Then .ListIndex = 0
    End With

End Sub

Private Sub cmbProfessionnel_AfterUpdate() '2025-05-31 @ 16:11

'    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:cmbProfessionnel_AfterUpdate", "", 0)

    Dim initProfAutorises As String
    
    initProfAutorises = GetInitialesObligatoiresFromADMIN(GetNomUtilisateur())

    Select Case initProfAutorises
        Case "INVALID"
            MsgBox "Votre code Windows n'est pas reconnu dans la liste d'administration.", vbExclamation
            cmbProfessionnel.value = ""
            Exit Sub
        Case ""
            'Aucune restriction sur les initiales à utiliser
        Case Else
            If cmbProfessionnel.value <> initProfAutorises Then
                MsgBox "Selon votre code d'utilisateur Windows" & vbNewLine & vbNewLine & _
                       "Vous devez obligatoirement utiliser le code '" & initProfAutorises & "'", vbInformation
            End If
            cmbProfessionnel.value = initProfAutorises
    End Select

    With ufSaisieHeures
        If .cmbProfessionnel.value <> "" Then
            .txtProfID.value = Fn_GetID_From_Initials(.cmbProfessionnel.value)
            If .txtDate.value <> "" Then
                Call TEC_Get_All_TEC_AF
                Call TEC_Refresh_ListBox_And_Add_Hours
            End If
        End If
    End With
    
    Call VerifierEtatBoutonAjouter '2025-06-09 @ 08:25

End Sub

Private Sub txtDate_Enter()

    If ufSaisieHeures.txtDate.value = "" Then
        ufSaisieHeures.txtDate.value = Format$(Date, wsdADMIN.Range("B1").value)
    Else
        ufSaisieHeures.txtDate.value = Format$(ufSaisieHeures.txtDate.value, wsdADMIN.Range("B1").value)
    End If
    
End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtDate_BeforeUpdate", "", 0)
    
    Dim fullDate As Variant
    
    fullDate = Fn_Complete_Date(ufSaisieHeures.txtDate.value, 600, 15)
    If fullDate <> "Invalid Date" Then
        Call Log_Saisie_Heures("info     ", "@00199 - fullDate = " & fullDate & _
                                "   y = " & year(fullDate) & _
                                "   m = " & month(fullDate) & _
                                "   d = " & day(fullDate) & _
                                "   type = " & TypeName(fullDate))
    End If
    
    'Update the cell with the full date, if valid
    If fullDate <> "Invalid Date" Then
        ufSaisieHeures.txtDate.value = Format$(fullDate, wsdADMIN.Range("B1").value)
    Else
        Cancel = True
        With ufSaisieHeures.txtDate
            .SetFocus 'Remettre le focus sur la TextBox
            .SelStart = 0 'Début de la sélection
            .SelLength = Len(.Text) 'Sélectionner tout le texte
        End With
        Exit Sub
    End If
    
    If fullDate > DateSerial(year(Date), month(Date), day(Date)) Then
        If MsgBox("En êtes-vous CERTAIN de vouloir cette date ?" & vbNewLine & vbNewLine & _
                    "La date saisie est '" & Format$(fullDate, wsdADMIN.Range("B1").value) & "'", vbYesNo + vbQuestion, _
                    "Utilisation d'une date FUTURE") = vbNo Then
            txtDate.SelStart = 0
            txtDate.SelLength = Len(Me.txtDate.value)
            txtDate.SetFocus
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = False
    
    Call Log_Record("ufSaisieHeures:txtDate_BeforeUpdate", "", startTime)
    
End Sub

Private Sub txtDate_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtDate_AfterUpdate", "", 0)
    
    If IsDate(ufSaisieHeures.txtDate.value) Then
        Dim dateStr As String, dateFormated As Date
        dateStr = ufSaisieHeures.txtDate.value
        dateFormated = DateSerial(year(dateStr), month(dateStr), day(dateStr))
        ufSaisieHeures.txtDate.value = Format$(dateFormated, wsdADMIN.Range("B1").value)
    Else
        ufSaisieHeures.txtDate.SetFocus
        ufSaisieHeures.txtDate.SelLength = Len(ufSaisieHeures.txtDate.value)
        ufSaisieHeures.txtDate.SelStart = 0
        Exit Sub
    End If

    If ufSaisieHeures.txtProfID.value <> "" Then
        Call TEC_Get_All_TEC_AF
        Call TEC_Refresh_ListBox_And_Add_Hours
    End If
    
    Call VerifierEtatBoutonAjouter '2025-06-09 @ 08:25
    
    Call Log_Record("ufSaisieHeures:txtDate_AfterUpdate", "", startTime)
    
End Sub

Private Sub txtClient_Enter()

    If rmv_state = rmv_modeInitial Then
        rmv_state = rmv_modeCreation
    End If

End Sub

Private Sub txtClient_AfterUpdate()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtClient_AfterUpdate", ufSaisieHeures.txtClient.value, 0)
    
'    If Me.txtClient.value <> savedClient Then '2025-03-25 @ 13:05
'        If Me.txtTECID = "" Then
'            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtClient_AfterUpdate", False, False, False, True)
'        Else
'            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtClient_AfterUpdate", False, True, False, True)
'        End If
'    End If

    'Force à cacher le listbox pour les résultats de recherche
    On Error Resume Next
    Me.lstboxNomClient.Visible = False
    On Error GoTo 0
    
    Call VerifierEtatBoutonAjouter '2025-06-09 @ 08:25
    
    Call Log_Record("ufSaisieHeures:txtClient_AfterUpdate", Me.txtTECID, startTime)
    
End Sub

Private Sub txtActivite_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtActivite_AfterUpdate", Me.txtActivite.value, 0)
    
'    If Me.txtActivite.value <> savedActivite Then '2025-03-25 @ 13:05
''        Debug.Print "txtActivite_AfterUpdate : ", Me.txtActivite.value, " vs ", savedActivite, " - TECID=" & Me.txtTECID
'        If Me.txtTECID = "" Then
'            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtActivite_AfterUpdate", False, False, False, True)
'        Else
'            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtActivite_AfterUpdate", False, True, False, True)
'        End If
'    End If
'
'    If Me.txtActivite.value <> savedActivite Then '2025-01-16 @ 16:46
'        If Me.txtHeures.value <> "" Then
'            If CCur(Me.txtHeures.value) <> 0 Then
'                Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtActivite_AfterUpdate", True, False, False, True)
'            Else
'                Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtActivite_AfterUpdate", False, True, False, True)
'            End If
'        End If
'    End If

    Me.txtActivite.value = Fn_Nettoyer_Fin_Chaine(Me.txtActivite.value)
    
    Call VerifierEtatBoutonAjouter '2025-06-09 @ 08:25
    
    Call Log_Record("ufSaisieHeures:txtActivite_AfterUpdate", Me.txtTECID, startTime)
    
End Sub

Private Sub txtHeures_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtHeures_Exit", Me.txtHeures.value, 0)
    
    Dim heure As Currency
    
    On Error Resume Next
    heure = CCur(Me.txtHeures.value)
    On Error GoTo 0
    
    If Not IsNumeric(Me.txtHeures.value) Then
        MsgBox Prompt:="La valeur saisie ne peut être utilisée comme valeur numérique!", _
                Title:="Validation d'une valeur numérique", _
                Buttons:=vbCritical
'        Cancel = True
        Me.txtHeures.SelStart = 0
        Me.txtHeures.SelLength = Len(Me.txtHeures.value)
        Me.txtHeures.SetFocus
        DoEvents
        Exit Sub
    End If

    If heure < 0 Or heure > 24 Then
        MsgBox _
            Prompt:="Le nombre d'heures ne peut être une valeur négative" & vbNewLine & vbNewLine & _
                    "ou dépasser 24 heures pour une charge", _
            Title:="Validation d'une valeur numérique", _
            Buttons:=vbCritical
        Cancel = True
        Me.txtHeures.SetFocus
        DoEvents
        Me.txtHeures.SelStart = 0
        Me.txtHeures.SelLength = Len(Me.txtHeures.value)
        Exit Sub
    End If
    
    If Fn_Valider_Portion_Heures(heure) = False Then
        MsgBox _
            Prompt:="La portion fractionnaire (" & heure & ") des heures est invalide" & vbNewLine & vbNewLine & _
                    "Seules les valeurs de dixième et de quart d'heure sont" & vbNewLine & vbNewLine & _
                    "acceptables", _
            Title:="Les valeurs permises sont les dixièmes et les quarts d'heure seulement", _
            Buttons:=vbCritical
            
        Cancel = True  'Empêche de quitter le TextBox
        DoEvents
        Me.txtHeures.SetFocus 'Remet le focus explicitement
        Me.txtHeures.SelStart = 0
        Me.txtHeures.SelLength = Len(Me.txtHeures.Text)
    End If
    
    Call Log_Record("ufSaisieHeures:txtHeures_Exit", "", startTime)
    
End Sub

Sub txtHeures_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate", Me.txtHeures.value, 0)
    
    'Validation des heures saisies
    Dim strHeures As String
    strHeures = Me.txtHeures.value
    
    strHeures = Replace(strHeures, ".", ",")
    
    Me.txtHeures.value = Format$(strHeures, "#0.00")
    
'    If CCur(Me.txtHeures.value) <> savedHeures Then '2025-03-25 @ 13:05
''        Debug.Print "txtHeures_AfterUpdate : ", Me.txtHeures.value, " vs ", savedHeures, " - TECID=" & Me.txtTECID
'        If Me.txtTECID = "" Then 'Création d'une nouvelle charge
'            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtHeures_AfterUpdate", True, False, False, True)
'        Else 'Modification d'une charge existante
'            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtHeures_AfterUpdate", False, True, False, True)
'        End If
'    End If
'
    Call VerifierEtatBoutonAjouter '2025-06-09 @ 08:25
    
    Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate", Me.txtTECID, startTime)
    
End Sub

Private Sub chbFacturable_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:chbFacturable_AfterUpdate", "", 0)
    
    If Me.chbFacturable.value <> savedFacturable Then '2025-03-25 @ 13:05
        Debug.Print "chbFacturable_AfterUpdate : ", Me.chbFacturable.value, " vs ", savedFacturable, " - TECID=" & Me.txtTECID
        If Me.txtTECID = "" Then
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("chbFacturable_AfterUpdate", True, False, False, True) '2024-10-06 @ 14:33
        Else
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("chbFacturable_AfterUpdate", False, True, False, True)
        End If
    End If

    Call VerifierEtatBoutonAjouter '2025-06-09 @ 08:25
    
    Call Log_Record("ufSaisieHeures:chbFacturable_AfterUpdate", Me.txtTECID, startTime)
    
End Sub

Private Sub txtCommNote_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtCommNote_AfterUpdate", Me.txtCommNote.value, 0)
    
    If Me.txtCommNote.value <> savedCommNote Then '2025-03-25 @ 13:05
        Debug.Print "txtCommNote_AfterUpdate : ", Me.txtCommNote.value, " vs ", savedCommNote, " - TECID=" & Me.txtTECID
        If Me.txtTECID = "" Then
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtCommNote_AfterUpdate", True, False, False, True) '2024-10-06 @ 14:33
        Else
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux("txtCommNote_AfterUpdate", False, True, True, True)
        End If
    End If

    Call VerifierEtatBoutonAjouter '2025-06-09 @ 08:25
    
    Call Log_Record("ufSaisieHeures:txtCommNote_AfterUpdate", Me.txtTECID, startTime)
    
End Sub

'----------------------------------------------------------------- ButtonsEvents
Private Sub cmdClear_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:cmdClear_Click", "", 0)
    
    Call TEC_Efface_Formulaire

    Call Log_Record("ufSaisieHeures:cmdClear_Click", "", startTime)

End Sub

Private Sub cmdAdd_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:cmdAdd_Click", "", 0)
    
    Call TEC_Ajoute_Ligne

    Call Log_Record("ufSaisieHeures:cmdAdd_Click", "", startTime)

End Sub

Private Sub cmdUpdate_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:cmdUpdate_Click", ufSaisieHeures.txtTECID.value, 0)
    
    If ufSaisieHeures.txtTECID.value <> "" Then
        Call TEC_Modifie_Ligne
    Else
        MsgBox Prompt:="Vous devez choisir un enregistrement à modifier !", _
               Title:="", _
               Buttons:=vbCritical
    End If

    Call Log_Record("ufSaisieHeures:cmdUpdate_Click", "", startTime)

End Sub

Private Sub cmdDelete_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:cmdDelete_Click", ufSaisieHeures.txtTECID.value, 0)
    
    If ufSaisieHeures.txtTECID.value <> "" Then
        Call TEC_Efface_Ligne
    Else
        MsgBox Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
               Title:="", _
               Buttons:=vbCritical
    End If

    Call Log_Record("ufSaisieHeures:cmdDelete_Click", "", startTime)

End Sub

'Get a specific row from listBox and display it in the userform
Sub lsbHresJour_dblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:lsbHresJour_dblClick", ufSaisieHeures.lsbHresJour.ListIndex, 0)
    
    rmv_state = rmv_modeAffichage
    
    With ufSaisieHeures
        Dim tecID As Long
        tecID = .lsbHresJour.List(.lsbHresJour.ListIndex, 0)
        ufSaisieHeures.txtTECID.value = tecID
        txtTECID = tecID
        
        'Retrieve the record in wsdTEC_Local
        Dim lookupRange As Range, lastTECRow As Long, rowTECID As Long
        lastTECRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "A").End(xlUp).row
        Set lookupRange = wsdTEC_Local.Range("A3:A" & lastTECRow)
        rowTECID = Fn_Find_Row_Number_TECID(tecID, lookupRange)
        
        Dim isBilled As Boolean
        isBilled = wsdTEC_Local.Range("L" & rowTECID).value

        'Remplir le userForm, si cette charge n'a pas été facturée
        If Not isBilled Then
            Application.EnableEvents = False
            .cmbProfessionnel.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 1)
            .cmbProfessionnel.Enabled = False
    
            .txtDate.value = Format$(.lsbHresJour.List(.lsbHresJour.ListIndex, 2), wsdADMIN.Range("B1").value) '2025-01-31 @ 13:31
            .txtDate.Enabled = False
    
            .txtClient.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 3)
            savedClient = .txtClient.value
'            .txtSavedClient.value = .txtClient.value
            .txtClientID.value = wsdTEC_Local.Range("E" & rowTECID).value
    
            .txtActivite.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 4)
            savedActivite = .txtActivite.value
'            .txtSavedActivite.value = .txtActivite.value
    
            .txtHeures.value = Format$(.lsbHresJour.List(.lsbHresJour.ListIndex, 5), "#0.00")
            savedHeures = CCur(.txtHeures.value)
'            .txtSavedHeures.value = .txtHeures.value
    
            .txtCommNote.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 6)
            savedCommNote = .txtCommNote.value
'            .txtSavedCommNote.value = .txtCommNote.value
    
            .chbFacturable.value = CBool(.lsbHresJour.List(.lsbHresJour.ListIndex, 7))
            savedFacturable = .chbFacturable.value
'            .txtSavedFacturable.value = .chbFacturable.value
            Application.EnableEvents = True

        Else
            MsgBox "Il est impossible de modifier ou de détruire" & vbNewLine & _
                        vbNewLine & "une charge déjà FACTURÉE", vbExclamation
        End If
        
    End With

    Call modTEC_Saisie.ActiverButtonsVraiOuFaux("lsbHresJour_dblClick", False, False, True, True)
    
    rmv_state = rmv_modeModification
    
    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call Log_Record("ufSaisieHeures:lsbHresJour_dblClick[" & tecID & "]", "", startTime)

End Sub

Sub imgLogoGCF_Click()

    If ufSaisieHeures.cmbProfessionnel.value <> "" Then
            Application.EnableEvents = False
            
            wshTEC_TDB_Data.Range("S7").value = ufSaisieHeures.cmbProfessionnel.value
        
            Call ActualiserTEC_TDB
            
            Call Stats_Heures_AF
            
            'Mettre à jour les 4 tableaux croisés dynamiques (Semaine, Mois, Trimestre & Année Financière)
            Call UpdatePivotTables
            
            Application.EnableEvents = True
            
            ufStatsHeures.show vbModeless
    Else
        MsgBox "Vous devez minimalement saisir un code de Professionnel" & vbNewLine & vbNewLine & _
                "avant de pouvoir afficher vos statistiques", vbInformation, _
                "Statistiques personnelles des heures"
    End If

End Sub

Sub imgStats_Click()

    Application.EnableEvents = False
    
    ufSaisieHeures.Hide
    
    Application.EnableEvents = True
    
    gFromMenu = True
    
    With wshStatsHeuresPivotTables
        .Visible = xlSheetVisible
        .Activate
    End With

End Sub

Private Sub VerifierEtatBoutonAjouter() '2025-06-09 @ 08:13
    
    If Me.txtTECID = "" Then 'Mode création: tous les champs obligatoires doivent être remplis
        If _
            Trim(Me.cmbProfessionnel.value) <> "" And _
            Trim(Me.txtClient.value) <> "" And _
            Trim(Me.txtActivite.value) <> "" And _
            Trim(Me.txtHeures.value) <> "" Then
            Me.cmdAdd.Enabled = True
        Else
            Me.cmdAdd.Enabled = False
        End If
    Else
        'Mode modification: bouton Ajouter inactif, tu peux gérer le bouton Modifier ici
        Me.cmdAdd.Enabled = False
        'Par exemple : Me.btnModifier.Enabled = ...
    End If

End Sub

