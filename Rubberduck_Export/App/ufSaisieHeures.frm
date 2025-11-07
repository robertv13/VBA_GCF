VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   10365
   ClientLeft      =   30
   ClientTop       =   -15
   ClientWidth     =   15870
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

'Private colSurveillance As Collection
'
Private wrappers As Collection

'Sauvegarde des valeurs lues
Public valeurSauveeClient As String
Public valeurSauveeActivite As String
Public valeurSauveeHeures As Currency
Public valeurSauveeEstFacturable As String
Public valeurSauveeCommNote As String

'Allows the calling code to set the data
Public Property Let ListData(ByVal rg As Range)

    oEventHandler.List = rg.Value

End Property

Private Sub UserForm_Initialize() '2025-05-30 @ 13:26

'    Set colSurveillance = New Collection
'    Dim ctrl As Control
'    Dim obj As clsSurveillanceActivite
'
'    Dim ctrlType As String
'    For Each ctrl In Me.Controls
'        ctrlType = TypeName(ctrl)
'        Debug.Print ctrl.Name & " - " & ctrlType
'        Select Case ctrlType
'            Case "TextBox", "ComboBox", "CommandButton"
'                Set obj = New clsSurveillanceActivite
'                Select Case ctrlType
'                    Case "TextBox": Set obj.tb = ctrl
'                    Case "ComboBox": Set obj.cb = ctrl
'                    Case "CommandButton": Set obj.btn = ctrl
'                End Select
'                If obj Is Nothing Then
'                    Debug.Print "Objet non initialisé pour : " & ctrl.Name
'                Else
'                    colSurveillance.Add obj
'                End If
'            Case Else
'                Debug.Print "Contrôle ignoré : " & ctrl.Name & " (" & ctrlType & ")"
'        End Select
'    Next ctrl
    
    Call InitialiserSurveillanceForm(Me, wrappers)
    
End Sub

Sub UserForm_Activate() '2024-07-31 @ 07:57

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:UserForm_Activate", vbNullString, 0)
    
    gLogSaisieHeuresVeryDetailed = False
    
    Call ImporterClients
    
    'Mise en place de la colonne à rechercher dans BD_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wsdBD_Clients.Cells(wsdBD_Clients.Rows.count, 1).End(xlUp).Row
    ufSaisieHeures.ListData = wsdBD_Clients.Range("Q1:Q" & lastUsedRow) '2025-01-11 @ 18:00
    
    With oEventHandler
        Set .SearchListBox = lstNomClient
        Set .SearchTextBox = txtClient
        .MaxRows = 10
        .ShowAllMatches = False
        .CompareMethod = vbTextCompare
    End With

    ufSaisieHeures.txtDate.Value = vbNullString 'On vide la date pour forcer la saisie
    
    On Error Resume Next
    ufSaisieHeures.cmbProfessionnel.SetFocus
    On Error GoTo 0
   
    rmv_state = MODE_INITIAL_FAC
    
    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:UserForm_Activate", vbNullString, startTime)
    
End Sub

Private Sub lstNomClient_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:lstNomClient_DblClick", vbNullString, 0)
    
    Dim i As Long
    With Me.lstNomClient
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Me.txtClient.Value = .List(i, 0)
                Me.txtClientID.Value = Fn_CellSpecifiqueDeBDClient(Me.txtClient.Value, 17, 2)
                Me.txtClientReel.Value = Fn_CellSpecifiqueDeBDClient(Me.txtClientID.Value, 2, 1)
                Exit For
            End If
        Next i
    End With
    
    'Force à cacher le listbox pour les résultats de recherche
    On Error Resume Next
    Me.lstNomClient.Visible = False
    On Error GoTo 0
    
    Me.txtClient.TextAlign = fmTextAlignLeft

    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:lstNomClient_DblClick", vbNullString, startTime)

End Sub

Private Sub UserForm_Terminate()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:UserForm_Terminate", vbNullString, 0)

    'Libérer la mémoire
    Set oEventHandler = Nothing
    
    ufSaisieHeures.Hide
    Unload ufSaisieHeures
    
    Call RetournerAuMenu

    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:UserForm_Terminate", vbNullString, startTime)

End Sub

Sub RetournerAuMenu()

    Call modAppli.QuitterFeuillePourMenu(wshMenuTEC, True)

End Sub

Private Sub cmbProfessionnel_Enter() '2025-05-31 @ 16:31

    Dim ws As Worksheet
    Dim plageInitiales As Range
    Dim cell As Range, cellInit As Range
    Dim listeInitiales As Collection
    Dim utilisateur As String
    Dim toutesInitiales As Boolean

    Set ws = wsdADMIN
    'Plage de la table WindowsUser_Initials : colonnes D à F, lignes 63 à 78
    Set plageInitiales = ws.Range("D68:D77")
    
    utilisateur = modFunctions.Fn_UtilisateurWindows() ' Variable globale utilisateur Windows

    Set listeInitiales = New Collection
    toutesInitiales = False
    
    'Chercher utilisateur dans la liste et récupérer initiales permises
    For Each cell In plageInitiales
        If Trim(cell.Value) <> vbNullString Then
            If StrComp(cell.Value, utilisateur, vbTextCompare) = 0 Then
                If Trim(cell.offset(0, 2).Value) <> vbNullString Then
                    'Initiales spécifiques autorisées pour cet utilisateur
                    listeInitiales.Add Trim(cell.offset(0, 2).Value)
                Else
                    'Pas de restriction, on doit autoriser toutes les initiales
                    toutesInitiales = True
                End If
                Exit For ' Utilisateur trouvé, on peut sortir
            End If
        End If
    Next cell

    'Si toutes les initiales sont permises, on ajoute toutes celles listées en colonne F
    If toutesInitiales Then
        'Ajoute GC qui est la valeur par défut
        listeInitiales.Add "GC", "GC"
        For Each cellInit In plageInitiales.offset(, 2).Resize(, 1)
            If Trim(cellInit.Value) <> vbNullString Then
                On Error Resume Next 'Eviter doublons dans la collection
                If Trim(cellInit.Value) <> "Init. Permises" Then
'                    Debug.Print "XYZ - " & Trim(cellInit.value)
                    listeInitiales.Add Trim(cellInit.Value), CStr(Trim(cellInit.Value))
                End If
                On Error GoTo 0
            End If
        Next cellInit
    End If

    'Remplir le ComboBox
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

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:cmbProfessionnel_AfterUpdate", vbNullString, 0)

    Dim initProfAutorises As String
    
    initProfAutorises = GetInitialesObligatoiresFromADMIN(modFunctions.Fn_UtilisateurWindows())

    Select Case initProfAutorises
        Case "INVALID"
            MsgBox "Les initiales saisies ne sont pas permise pour votre code d'utilisateur Windows", vbExclamation
            cmbProfessionnel.Value = vbNullString
            Exit Sub
        Case vbNullString
'            cmbProfessionnel.Value = vbNullString 'Aucune restriction sur les initiales à utiliser
        Case Else
            If cmbProfessionnel.Value <> initProfAutorises Then
                MsgBox "Selon votre code d'utilisateur Windows" & vbNewLine & vbNewLine & _
                       "Vous devez obligatoirement utiliser le code '" & initProfAutorises & "'", vbInformation
            End If
            cmbProfessionnel.Value = initProfAutorises
    End Select

    With ufSaisieHeures
        If .cmbProfessionnel.Value <> vbNullString Then
            .txtProfID.Value = Fn_ProfIDAPartirDesInitiales(.cmbProfessionnel.Value)
            If .txtDate.Value <> vbNullString Then
                Call ObtenirTousLesTECDateAvecAF
                Call RafraichirListBoxEtAddtionnerHeures
            End If
        End If
    End With
    
    'Lorsqu'on change de professionnel, on force l'importation des TEC - 2025-06-13 @ 08:46
    Call ImporterTEC
    Me.txtLastImport.Value = "Les TEC ont été importés à " & Format$(Now, "hh:mm:ss")
    
    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:cmbProfessionnel_AfterUpdate", vbNullString, startTime)

End Sub

Private Sub txtDate_Enter()

    If ufSaisieHeures.txtDate.Value = vbNullString Then
        ufSaisieHeures.txtDate.Value = Format$(Date, wsdADMIN.Range("B1").Value)
    Else
        ufSaisieHeures.txtDate.Value = Format$(ufSaisieHeures.txtDate.Value, wsdADMIN.Range("B1").Value)
    End If
    
End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    'Routine de validation de date
    Dim valeur As Variant
    valeur = Fn_DateNormalisee(Me.txtDate.Value)
    
    If IsError(valeur) Then
        MsgBox "La date saisie est invalide. Veuillez corriger la saisie.", vbExclamation
        Cancel = True
    Else
        Me.txtDate.Value = Format(valeur, wsdADMIN.Range("B1").Value)
    End If
    
    Dim fullDate As Variant
    
    fullDate = Fn_CompleteLaDate(ufSaisieHeures.txtDate.Value, 600, 15)
    If fullDate <> "Invalid Date" Then
        Call EnregistrerLogSaisieHeures("info     ", "@00199 - fullDate = " & fullDate & _
                                "   y = " & year(fullDate) & _
                                "   m = " & month(fullDate) & _
                                "   d = " & day(fullDate) & _
                                "   type = " & TypeName(fullDate))
    End If
    
    'Update the cell with the full date, if valid
    If fullDate <> "Invalid Date" Then
        ufSaisieHeures.txtDate.Value = Format$(fullDate, wsdADMIN.Range("B1").Value)
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
                    "La date saisie est '" & Format$(fullDate, wsdADMIN.Range("B1").Value) & "'", vbYesNo + vbQuestion, _
                    "Utilisation d'une date FUTURE") = vbNo Then
            txtDate.SelStart = 0
            txtDate.SelLength = Len(Me.txtDate.Value)
            txtDate.SetFocus
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = False
    
End Sub

Private Sub txtDate_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:txtDate_AfterUpdate", vbNullString, 0)
    
    If IsDate(ufSaisieHeures.txtDate.Value) Then
        Dim dateStr As String, dateFormated As Date
        dateStr = ufSaisieHeures.txtDate.Value
        dateFormated = DateSerial(year(dateStr), month(dateStr), day(dateStr))
        ufSaisieHeures.txtDate.Value = Format$(dateFormated, wsdADMIN.Range("B1").Value)
    Else
        ufSaisieHeures.txtDate.SetFocus
        ufSaisieHeures.txtDate.SelLength = Len(ufSaisieHeures.txtDate.Value)
        ufSaisieHeures.txtDate.SelStart = 0
        Exit Sub
    End If

    If ufSaisieHeures.txtProfID.Value <> vbNullString Then
        Call ObtenirTousLesTECDateAvecAF
        Call RafraichirListBoxEtAddtionnerHeures
    End If
    
    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:txtDate_AfterUpdate", vbNullString, startTime)
    
End Sub

Private Sub txtClient_Enter()

    If rmv_state = MODE_INITIAL_FAC Then
        rmv_state = MODE_CREATION_FAC
    End If

End Sub

Private Sub txtClient_AfterUpdate()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:txtClient_AfterUpdate", ufSaisieHeures.txtClient.Value, 0)
    
    'Force à cacher le listbox pour les résultats de recherche
    On Error Resume Next
    Me.lstNomClient.Visible = False
    On Error GoTo 0
    
    Call MettreAJourEtatBoutons
    
    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:txtClient_AfterUpdate", Me.txtTECID, startTime)
    
End Sub

Private Sub txtActivite_AfterUpdate()

    Me.txtActivite.Value = Fn_ChaineNettoyeeCaracteresSpeciaux(Me.txtActivite.Value)
    
    Call MettreAJourEtatBoutons
    
End Sub

Private Sub txtHeures_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim heure As Currency
    
    On Error Resume Next
    heure = CCur(Me.txtHeures.Value)
    On Error GoTo 0
    
    If Not IsNumeric(Me.txtHeures.Value) Then
        MsgBox Prompt:="La valeur saisie ne peut être utilisée comme valeur numérique!", _
                Title:="Validation d'une valeur numérique", _
                Buttons:=vbCritical
'        Cancel = True
        Me.txtHeures.SelStart = 0
        Me.txtHeures.SelLength = Len(Me.txtHeures.Value)
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
        Me.txtHeures.SelLength = Len(Me.txtHeures.Value)
        Exit Sub
    End If
    
    If Fn_FractionHeureEstValide(heure) = False Then
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
        Me.txtHeures.SelLength = Len(Me.txtHeures.text)
    End If
    
End Sub

Sub txtHeures_AfterUpdate()

    'Validation des heures saisies
    Dim strHeures As String
    strHeures = Me.txtHeures.Value
    
    strHeures = Replace(strHeures, ".", ",")
    
    Me.txtHeures.Value = Format$(strHeures, "#0.00")
    
    Call MettreAJourEtatBoutons
    
End Sub

Private Sub chkFacturable_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:chkFacturable_AfterUpdate", vbNullString, 0)
    
    If Me.chkFacturable.Value <> valeurSauveeEstFacturable Then '2025-03-25 @ 13:05
        If Me.txtTECID = vbNullString Then
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(True, False, False, True) '2024-10-06 @ 14:33
        Else
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, True, False, True)
        End If
    End If

    Call MettreAJourEtatBoutons
    
    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:chkFacturable_AfterUpdate", Me.txtTECID, startTime)
    
End Sub

Private Sub txtCommNote_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:txtCommNote_AfterUpdate", Me.txtCommNote.Value, 0)
    
    If Me.txtCommNote.Value <> valeurSauveeCommNote Then '2025-03-25 @ 13:05
        Debug.Print "txtCommNote_AfterUpdate : ", Me.txtCommNote.Value, " vs ", valeurSauveeCommNote, " - TECID=" & Me.txtTECID
        If Me.txtTECID = vbNullString Then
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(True, False, False, True) '2024-10-06 @ 14:33
        Else
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, True, True, True)
        End If
    End If

    Call MettreAJourEtatBoutons
    
    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:txtCommNote_AfterUpdate", Me.txtTECID, startTime)
    
End Sub

Private Sub shpAdd_Click()

    Call AjouterLigneTEC

End Sub

Private Sub shpUpdate_Click()

    If ufSaisieHeures.txtTECID.Value <> vbNullString Then
        Call ModifierLigneTEC
    Else
        MsgBox Prompt:="Vous devez choisir un enregistrement à modifier !", _
               Title:=vbNullString, _
               Buttons:=vbCritical
    End If

End Sub

Private Sub shpDelete_Click()

    If ufSaisieHeures.txtTECID.Value <> vbNullString Then
        Call DetruireLigneTEC
    Else
        MsgBox Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
               Title:=vbNullString, _
               Buttons:=vbCritical
    End If

End Sub

Private Sub shpClear_Click()

    Call EffacerFormulaireTEC
    
    Call MettreAJourEtatBoutons

End Sub

'Get a specific row from listBox and display it in the userform
Sub lstHresJour_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:lstHresJour_dblClick", ufSaisieHeures.lstHresJour.ListIndex, 0)
    
    rmv_state = MODE_AFFICHAGE_FAC
    
    With ufSaisieHeures
        Dim tecID As Long
        tecID = .lstHresJour.List(.lstHresJour.ListIndex, 0)
        ufSaisieHeures.txtTECID.Value = tecID
        txtTECID = tecID
        
        'Retrieve the record in wsdTEC_Local
        Dim lookupRange As Range, lastTECRow As Long, rowTECID As Long
        lastTECRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "A").End(xlUp).Row
        Set lookupRange = wsdTEC_Local.Range("A3:A" & lastTECRow)
        rowTECID = Fn_Find_Row_Number_TECID(tecID, lookupRange)
        
        Dim isBilled As Boolean
        isBilled = wsdTEC_Local.Range("L" & rowTECID).Value

        'Remplir le userForm, si cette charge n'a pas été facturée
        If Not isBilled Then
            Application.EnableEvents = False
            .cmbProfessionnel.Value = .lstHresJour.List(.lstHresJour.ListIndex, 1)
            .cmbProfessionnel.Enabled = False
    
            .txtDate.Value = Format$(.lstHresJour.List(.lstHresJour.ListIndex, 2), wsdADMIN.Range("B1").Value) '2025-01-31 @ 13:31
            .txtDate.Enabled = False
    
            .txtClient.Value = .lstHresJour.List(.lstHresJour.ListIndex, 3)
            valeurSauveeClient = .txtClient.Value
            .txtClientID.Value = wsdTEC_Local.Range("E" & rowTECID).Value
    
            .txtActivite.Value = .lstHresJour.List(.lstHresJour.ListIndex, 4)
            valeurSauveeActivite = .txtActivite.Value
    
            .txtHeures.Value = Format$(.lstHresJour.List(.lstHresJour.ListIndex, 5), "#0.00")
            valeurSauveeHeures = CCur(.txtHeures.Value)
    
            .txtCommNote.Value = .lstHresJour.List(.lstHresJour.ListIndex, 6)
            valeurSauveeCommNote = .txtCommNote.Value
    
            .chkFacturable.Value = CBool(.lstHresJour.List(.lstHresJour.ListIndex, 7))
            valeurSauveeEstFacturable = .chkFacturable.Value
            Application.EnableEvents = True

        Else
            MsgBox "Il est impossible de modifier ou de détruire" & vbNewLine & _
                    vbNewLine & "une charge déjà FACTURÉE", vbExclamation
        End If
        
    End With
    
    Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, False, True, True)    'Ajustement des boutons
    
    rmv_state = MODE_MODIFICATION_FAC
    
    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("ufSaisieHeures:lstHresJour_dblClick[" & tecID & "]", vbNullString, startTime)

End Sub

Sub imgLogoGCF_Click()

    If ufSaisieHeures.cmbProfessionnel.Value <> vbNullString Then
            Application.EnableEvents = False
            
            wshTEC_TDB_Data.Range("S7").Value = ufSaisieHeures.cmbProfessionnel.Value
        
            Call modTEC_TDB.ActualiserTECTableauDeBord
            
            Call ExecuterAdvancedFilterSurTECTDBData
            
            'Mettre à jour les 4 tableaux croisés dynamiques (Semaine, Mois, Trimestre & Année Financière)
            Call MettreAJourPivotTables
            
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

Private Sub MettreAJourEtatBoutons() '2025-07-03 @ 07:09

    Dim enAjout As Boolean
    Dim enModification As Boolean
    Dim estModifie As Boolean
    Dim tousChampsRemplis As Boolean

    enAjout = (Trim(Me.txtTECID.Value & vbNullString) = vbNullString)
    enModification = Not enAjout

    'Vérification champs requis pour Ajout
    tousChampsRemplis = _
        Trim(Me.cmbProfessionnel.Value & vbNullString) <> vbNullString And _
        Trim(Me.txtClient.Value & vbNullString) <> vbNullString And _
        Trim(Me.txtActivite.Value & vbNullString) <> vbNullString And _
        Trim(Me.txtHeures.Value & vbNullString) <> vbNullString

    'Bouton Ajouter
    Me.shpAdd.Enabled = enAjout And tousChampsRemplis

    'Comparaison avec valeurs originales (stockées à la lecture en BD)
    estModifie = False
    estModifie = _
        EstChampModifie(Me.txtClient.Value, valeurSauveeClient) Or _
        EstChampModifie(Me.txtActivite.Value, valeurSauveeActivite) Or _
        (Me.txtHeures.Value <> valeurSauveeHeures) Or _
        (Me.chkFacturable.Value <> valeurSauveeEstFacturable) Or _
        EstChampModifie(Me.txtCommNote.Value, valeurSauveeCommNote)

    'Bouton estModifier
    Me.shpUpdate.Enabled = enModification And estModifie

    'Bouton Détruire
    Me.shpDelete.Enabled = enModification And Not estModifie

    'Bouton Annuler
    Me.shpClear.Enabled = Me.shpAdd.Enabled Or Me.shpUpdate.Enabled

End Sub


