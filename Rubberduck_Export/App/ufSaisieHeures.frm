VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   10305
   ClientLeft      =   195
   ClientTop       =   780
   ClientWidth     =   15975
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

'Allows the calling code to set the data
Public Property Let ListData(ByVal rg As Range)

    oEventHandler.List = rg.value

End Property

Sub UserForm_Activate() '2024-07-31 @ 07:57

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:UserForm_Activate", "", 0)
    
    logSaisieHeuresVeryDetailed = False
    
    Call Client_List_Import_All
    Call TEC_Import_All
    
    'Mise en place de la colonne à rechercher dans BD_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wshBD_Clients.Cells(wshBD_Clients.Rows.count, 1).End(xlUp).row
    ufSaisieHeures.ListData = wshBD_Clients.Range("Q1:Q" & lastUsedRow) '2025-01-11 @ 18:00
'    ufSaisieHeures.ListData = wshBD_Clients.Range("A1:B" & lastUsedRow) '2024-11-05 @ 07:05
    
    With oEventHandler
        Set .SearchListBox = lstboxNomClient
        Set .SearchTextBox = txtClient
        .MaxRows = 10
        .ShowAllMatches = False
        .CompareMethod = vbTextCompare
    End With

    Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, False, False, False)

    'Default Professionnal - 2024-08-19 @ 07:59
    Select Case Fn_Get_Windows_Username
        Case "Guillaume", "GuillaumeCharron", "gchar", "Robert M. Vigneault", "Robertmv"
            cmbProfessionnel.value = "GC"
        Case "vgervais"
            cmbProfessionnel.value = "VG"
        Case "User"
            cmbProfessionnel.value = "ML"
        Case "vgervais"
            cmbProfessionnel.value = "AR"
        Case Else
            cmbProfessionnel.value = ""
    End Select
    
    ufSaisieHeures.txtDate.value = "" 'On vide la date pour forcer la saisie
    
    On Error Resume Next
    ufSaisieHeures.cmbProfessionnel.SetFocus
    On Error GoTo 0
   
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
                Me.txtClientRéel.value = Fn_Cell_From_BD_Client(Me.txtClientID.value, 2, 1)
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

Public Sub cmbProfessionnel_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:cmbProfessionnel_AfterUpdate", "", 0)

    'Restreindre l'accès au professionnel par défaut du code d'utilisateur
    Select Case Fn_Get_Windows_Username
        Case "Guillaume", "GuillaumeCharron", "gchar", "Robert M. Vigneault", "robertmv"
            'Accès à tous les utilisateurs
        Case "vgervais"
            If cmbProfessionnel.value <> "VG" Then
                MsgBox "Selon votre code d'utilisateur Windows" & vbNewLine & vbNewLine & _
                    "Vous devez obligatoirement utiliser le code 'VG'", _
                    vbInformation
            End If
            cmbProfessionnel.value = "VG"
        Case "User"
            If cmbProfessionnel.value <> "ML" Then
                MsgBox "Selon votre code d'utilisateur Windows" & vbNewLine & vbNewLine & _
                        "Vous devez obligatoirement utiliser le code 'ML'", _
                        vbInformation
            End If
            cmbProfessionnel.value = "ML"
        Case "Annie"
            If cmbProfessionnel.value <> "AR" Then
                MsgBox "Selon votre code d'utilisateur Windows" & vbNewLine & vbNewLine & _
                    "Vous devez obligatoirement utiliser le code 'AR'", _
                    vbInformation
            End If
            cmbProfessionnel.value = "AR"
        Case Else
            cmbProfessionnel.value = ""
    End Select

    If ufSaisieHeures.cmbProfessionnel.value <> "" Then
        ufSaisieHeures.txtProfID.value = Fn_GetID_From_Initials(ufSaisieHeures.cmbProfessionnel.value)
        If ufSaisieHeures.txtDate.value <> "" Then '2024-09-05 @ 20:50
            Call TEC_Get_All_TEC_AF
            Call TEC_Refresh_ListBox_And_Add_Hours
        End If
    End If

    Call Log_Record("ufSaisieHeures:cmbProfessionnel_AfterUpdate", "", startTime)

End Sub

Private Sub txtDate_Enter()

    If ufSaisieHeures.txtDate.value = "" Then
        ufSaisieHeures.txtDate.value = Format$(Date, wshAdmin.Range("B1").value)
    Else
        ufSaisieHeures.txtDate.value = Format$(ufSaisieHeures.txtDate.value, wshAdmin.Range("B1").value)
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
        ufSaisieHeures.txtDate.value = Format$(fullDate, wshAdmin.Range("B1").value)
    Else
        Cancel = True
        With ufSaisieHeures.txtDate
            .SetFocus 'Remettre le focus sur la TextBox
            .SelStart = 0 'Début de la sélection
            .SelLength = Len(.Text) 'Sélectionner tout le texte
        End With
        Exit Sub
    End If
    
    Call Log_Saisie_Heures("info     ", "@00223 - fullDate = " & fullDate & _
                                "   y = " & year(fullDate) & _
                                "   m = " & month(fullDate) & _
                                "   d = " & day(fullDate) & _
                                "   type = " & TypeName(fullDate))
    Call Log_Saisie_Heures("info     ", "@00228 - DateSerial = " & DateSerial(year(Date), month(Date), day(Date)) & _
                                "   y = " & year(Date) & _
                                "   m = " & month(Date) & _
                                "   d = " & day(Date) & _
                                "   type = " & TypeName(Date))
                                
    If fullDate > DateSerial(year(Date), month(Date), day(Date)) Then
        Call Log_Saisie_Heures("future  ", "@00230 - fullDate = " & fullDate & _
                                            "   y = " & year(fullDate) & _
                                            "   m = " & month(fullDate) & _
                                            "   d = " & day(fullDate) & _
                                            "   type = " & TypeName(fullDate))
        If MsgBox("En êtes-vous CERTAIN de vouloir cette date ?" & vbNewLine & vbNewLine & _
                    "La date saisie est '" & Format$(fullDate, wshAdmin.Range("B1").value) & "'", vbYesNo + vbQuestion, _
                    "Utilisation d'une date FUTURE") = vbNo Then
            txtDate.SelStart = 0
            txtDate.SelLength = Len(Me.txtDate.value)
            txtDate.SetFocus
            Cancel = True
            Exit Sub
        Else
            Call Log_Saisie_Heures("FUTURE_OK", "@00249 - fullDate = " & fullDate & _
                                                "   y = " & year(fullDate) & _
                                                "   m = " & month(fullDate) & _
                                                "   d = " & day(fullDate) & _
                                                "   type = " & TypeName(fullDate))
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
        ufSaisieHeures.txtDate.value = Format$(dateFormated, wshAdmin.Range("B1").value)
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
    
    Call Log_Record("ufSaisieHeures:txtDate_AfterUpdate", "", startTime)
    
End Sub

Private Sub txtClient_Enter()

    If rmv_state = rmv_modeInitial Then
        rmv_state = rmv_modeCreation
    End If

End Sub

Private Sub txtClient_AfterUpdate()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtClient_AfterUpdate", ufSaisieHeures.txtClient.value, 0)
    
    If Me.txtClient.value <> Me.txtSavedClient.value Then
        If Me.txtTECID = "" Then
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, False, False, True)
        Else
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, True, False, True)
        End If
    End If

    'Force à cacher le listbox pour les résultats de recherche
    On Error Resume Next
    Me.lstboxNomClient.Visible = False
    On Error GoTo 0
    
    Call Log_Record("ufSaisieHeures:txtClient_AfterUpdate", "", startTime)
    
End Sub

Private Sub txtActivite_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtActivite_AfterUpdate", Me.txtActivite.value, 0)
    
    If Me.txtActivite.value <> Me.txtSavedActivite.value Then
        If Me.txtTECID = "" Then
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, False, False, True)
        Else
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, True, False, True)
        End If
    End If
    
    If Me.txtActivite.value <> Me.txtSavedActivite.value Then '2025-01-16 @ 16:46
        If Me.txtHeures.value <> "" Then
            If CCur(Me.txtHeures.value) <> 0 Then
                Call modTEC_Saisie.ActiverButtonsVraiOuFaux(True, False, False, True)
            Else
                Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, True, False, True)
            End If
        End If
    End If
    
    Me.txtActivite.value = Fn_Nettoyer_Fin_Chaine(Me.txtActivite.value)
    
    Call Log_Record("ufSaisieHeures:txtActivite_AfterUpdate", "", startTime)
    
End Sub

Private Sub txtHeures_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtHeures_Exit", Me.txtHeures.value, 0)
    
    Dim heure As Currency
    
    On Error Resume Next
    heure = CCur(Me.txtHeures.value)
    On Error GoTo 0
    
    If Not IsNumeric(Me.txtHeures.value) Then
        MsgBox prompt:="La valeur saisie ne peut être utilisée comme valeur numérique!", _
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
            prompt:="Le nombre d'heures ne peut être une valeur négative" & vbNewLine & vbNewLine & _
                    "ou dépasser 24 pour une charge", _
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
        MsgBox "La portion fractionnaire (" & heure & ") des heures est invalide" & vbNewLine & vbNewLine & _
                "Seul les valeurs de dixième et de quart d'heure sont acceptables", vbCritical, _
                "Les valeurs permises sont les dixièmes et les quarts d'heure seulement"
        Cancel = True
        Me.txtHeures.SetFocus
        DoEvents
        Me.txtHeures.SelStart = 0
        Me.txtHeures.SelLength = Len(Me.txtHeures.value)
        Exit Sub
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
    
    If Me.txtHeures.value <> Me.txtSavedHeures.value Then
        If Me.txtTECID = "" Then 'Création d'une nouvelle charge
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(True, False, False, True)
        Else 'Modification d'une charge existante
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, True, False, True)
        End If
    End If
    
    Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate", "", startTime)
    
End Sub

Private Sub chbFacturable_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:chbFacturable_AfterUpdate", "", 0)
    
    If Me.chbFacturable.value <> Me.txtSavedFacturable.value Then
        If Me.txtTECID = "" Then
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(True, False, False, True) '2024-10-06 @ 14:33
        Else
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, True, False, True)
        End If
    End If

    Call Log_Record("ufSaisieHeures:chbFacturable_AfterUpdate", "", startTime)
    
End Sub

Private Sub txtCommNote_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtCommNote_AfterUpdate", Me.txtCommNote.value, 0)
    
    If Me.txtCommNote.value <> Me.txtSavedCommNote.value Then
        If Me.txtTECID = "" Then
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(True, False, False, True) '2024-10-06 @ 14:33
        Else
            Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, True, True, True)
        End If
    End If

    Call Log_Record("ufSaisieHeures:txtCommNote_AfterUpdate", "", startTime)
    
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
        MsgBox prompt:="Vous devez choisir un enregistrement à modifier !", _
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
        MsgBox prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
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
        
        'Retrieve the record in wshTEC_Local
        Dim lookupRange As Range, lastTECRow As Long, rowTECID As Long
        lastTECRow = wshTEC_Local.Cells(wshTEC_Local.Rows.count, "A").End(xlUp).row
        Set lookupRange = wshTEC_Local.Range("A3:A" & lastTECRow)
        rowTECID = Fn_Find_Row_Number_TECID(tecID, lookupRange)
        
        Dim isBilled As Boolean
        isBilled = wshTEC_Local.Range("L" & rowTECID).value

        'Remplir le userForm, si cette charge n'a pas été facturée
        If Not isBilled Then
            Application.EnableEvents = False
            .cmbProfessionnel.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 1)
            .cmbProfessionnel.Enabled = False
    
            .txtDate.value = Format$(.lsbHresJour.List(.lsbHresJour.ListIndex, 2), wshAdmin.Range("B1").value) '2025-01-31 @ 13:31
            .txtDate.Enabled = False
    
            .txtClient.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 3)
            savedClient = .txtClient.value
            .txtSavedClient.value = .txtClient.value
            .txtClientID.value = wshTEC_Local.Range("E" & rowTECID).value
    
            .txtActivite.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 4)
            savedActivite = .txtActivite.value
            .txtSavedActivite.value = .txtActivite.value
    
            .txtHeures.value = Format$(.lsbHresJour.List(.lsbHresJour.ListIndex, 5), "#0.00")
            savedHeures = .txtHeures.value
            .txtSavedHeures.value = .txtHeures.value
    
            .txtCommNote.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 6)
            savedCommNote = .txtCommNote.value
            .txtSavedCommNote.value = .txtCommNote.value
    
            .chbFacturable.value = CBool(.lsbHresJour.List(.lsbHresJour.ListIndex, 7))
            savedFacturable = .chbFacturable.value
            .txtSavedFacturable.value = .chbFacturable.value
            Application.EnableEvents = True

        Else
            MsgBox "Il est impossible de modifier ou de détruire" & vbNewLine & _
                        vbNewLine & "une charge déjà FACTURÉE", vbExclamation
        End If
        
    End With

    Call modTEC_Saisie.ActiverButtonsVraiOuFaux(False, False, True, True)
    
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
    
    fromMenu = True
    
    With wshStatsHeuresPivotTables
        .Visible = xlSheetVisible
        .Activate
    End With

End Sub


