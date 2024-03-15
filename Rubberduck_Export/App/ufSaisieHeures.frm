VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   12750
   ClientLeft      =   135
   ClientTop       =   570
   ClientWidth     =   17430
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
Private MyListBoxClass As clsListboxAlign

'Allows the calling code to set the data
Public Property Let ListData(ByVal rg As Range)

    oEventHandler.List = rg.value

End Property

Private Sub lstboxNomClient_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim timerStart As Double: timerStart = Timer
    
    Dim i As Long
    With Me.lstboxNomClient
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Me.txtClient.value = .List(i, 0)
                wshAdmin.Range("TEC_Client_ID").value = GetID_From_Client_Name(Me.txtClient.value)
                Exit For
            End If
        Next i
    End With
    
    Call Output_Timer_Results("lstboxNomClient_DblClick()", timerStart)

End Sub

Private Sub txtTEC_ID_Change()

End Sub

Sub UserForm_Activate()

    Dim timer3Start As Double: timer3Start = Timer
    
    Call Client_List_Import_All
    Call TEC_Import_All
    
    Dim lastUsedRow As Long
    lastUsedRow = wshBD_Clients.Range("A9999").End(xlUp).row
    ufSaisieHeures.ListData = wshBD_Clients.Range("A1:J" & lastUsedRow)
    
    With oEventHandler
        Set .SearchListBox = lstboxNomClient
        Set .SearchTextBox = txtClient
        .MaxRows = 10
        .ShowAllMatches = False
        .CompareMethod = vbTextCompare
    End With

    Call Buttons_Enabled_True_Or_False(False, False, False, False)

    cmbProfessionnel.value = ""
    cmbProfessionnel.SetFocus
   
    rmv_state = rmv_modeInitial
    
    Call Output_Timer_Results("ufSaisieHeures - UserForm_Activate()", timer3Start)
    
End Sub

Private Sub UserForm_Initialize()

    Set MyListBoxClass = New clsListboxAlign
    
End Sub

Private Sub UserForm_Terminate()
    
    Dim timerStart As Double: timerStart = Timer

    'Clear the admin control cells
    wshAdmin.Range("B3:B7").ClearContents
    
    ThisWorkbook.Save
    
    'Clean up
    Set oEventHandler = Nothing
    
    Me.Hide
    Unload Me
    
    If Me.name = "ufSaisieHeures" Then
        On Error GoTo MenuSelect
        wshMenuTEC.Select
        On Error GoTo 0
    Else
        wshMenu.Select
    End If
    Exit Sub
    
MenuSelect:
    wshMenu.Activate
    wshMenu.Select
    
    Call Output_Timer_Results("ufSaisieHeures - UserForm_Terminate()", timerStart)

End Sub

Public Sub cmbProfessionnel_AfterUpdate()

    Dim timerStart As Double: timerStart = Timer

    If Me.cmbProfessionnel.value = "" Then GoTo exit_sub
    
    wshAdmin.Range("TEC_Initials").value = Me.cmbProfessionnel.value
    wshAdmin.Range("TEC_Prof_ID").value = GetID_FromInitials(Me.cmbProfessionnel.value)
    
    If wshAdmin.Range("TEC_Date").value <> "" Then
        Call TEC_AdvancedFilter_And_Sort
        Call Refresh_ListBox_And_Add_Hours
    End If

exit_sub:

    Call Output_Timer_Results("ufSaisieHeures - cmbProfessionnel_AfterUpdate()", timerStart)

End Sub

Private Sub txtDate_Enter()

    If Me.txtDate.value = "" Then
        Me.txtDate.value = Format(Now(), "dd-mm-yyyy")
    End If

End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    Dim strDate As String
    strDate = Validate_A_Date(Me.txtDate.value) 'Returns a Valid date -OR- Empty string
    
    If strDate = "" Then '2024-03-02 @ 09:36 - RMV_MSGBOX
    MsgBox Prompt:="La valeur saisie ne peut être utilisée comme une date valide", _
        Title:="Validation de la date", _
        Buttons:=vbCritical
        txtDate.SelStart = 0
        txtDate.SelLength = Len(Me.txtDate.value)
        txtDate.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    Me.txtDate.value = strDate
    
    Debug.Print strDate & " - " & Format(Now(), "dd-mm-yyyy")
    If CDate(strDate) > Format(Now(), "dd-mm-yyyy") Then
        If MsgBox("En êtes-vous CERTAIN ?", vbYesNo + vbQuestion, "Utilisation d'une date FUTURE") = vbNo Then
            txtDate.SelStart = 0
            txtDate.SelLength = Len(Me.txtDate.value)
            txtDate.SetFocus
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = False
    
End Sub

Private Sub txtDate_AfterUpdate()

    If IsDate(Me.txtDate.value) Then
        wshAdmin.Range("TEC_Date").value = CDate(Me.txtDate.value)
    Else
        Me.txtDate.SetFocus
        Me.txtDate.SelLength = Len(Me.txtDate.value)
        Me.txtDate.SelStart = 0
        Exit Sub
    End If

    If wshAdmin.Range("TEC_Prof_ID").value <> "" Then
        Call TEC_AdvancedFilter_And_Sort
        Call Refresh_ListBox_And_Add_Hours
    End If
    
    'Enabled the NEW & ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> vbNullString And _
        Trim(Me.txtDate.value) <> vbNullString And _
        Trim(Me.txtClient.value) <> vbNullString And _
        Trim(Me.txtHeures.value) <> 0 Then
        Call Buttons_Enabled_True_Or_False(True, True, False, False)
    End If
    
End Sub

Private Sub txtClient_Enter()

    If rmv_state = rmv_modeInitial Then
        rmv_state = rmv_modeCreation
    End If

End Sub

Private Sub txtClient_AfterUpdate()
    
    If Me.txtClient.value <> Me.txtSavedClient.value Then
        If Me.txtTEC_ID.value = "" Then
            Call Buttons_Enabled_True_Or_False(True, False, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, False, False)
        End If
    End If
    
End Sub

Private Sub txtActivite_AfterUpdate()

    If Me.txtActivite.value <> Me.txtSavedActivite.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, False, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If

End Sub

Sub txtHeures_AfterUpdate()

    'Validation des heures saisies
    Dim strHeures As String
    strHeures = Me.txtHeures.value
    
    If InStr(".", strHeures) Then
        strHeures = Replace(strHeures, ".", ",")
    End If
    
    If IsNumeric(strHeures) = False Or strHeures > 24 Then
        MsgBox _
        Prompt:="La valeur saisie ne peut être utilisée comme valeur numérique!", _
        Title:="Validation d'une valeur numérique", _
        Buttons:=vbCritical
        Me.txtHeures.value = ""
        Me.txtHeures.SetFocus
        Exit Sub
    End If
    
    Me.txtHeures.value = Format(strHeures, "#0.00")
    
    If Me.txtHeures.value <> Me.txtSavedHeures.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If
    
End Sub

Private Sub chbFacturable_AfterUpdate()

    If Me.chbFacturable.value <> Me.txtSavedFacturable.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If

End Sub

Private Sub txtCommNote_AfterUpdate()

    If Me.txtCommNote.value <> Me.txtSavedCommNote.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If

'    If rmv_state = rmv_modeAffichage Then
'        If Me.txtCommNote.value <> savedCommNote Then
'            Call Buttons_Enabled_True_Or_False(True, False, True, True)
'        End If
'    End If

End Sub

'----------------------------------------------------------------- ButtonsEvents
Private Sub cmdClear_Click()

    Call TEC_Efface_Formulaire

End Sub

Private Sub cmdAdd_Click()

    Call TEC_Ajoute_Ligne

End Sub

Private Sub cmdUpdate_Click()

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à modifier !", _
               Title:="", _
               Buttons:=vbCritical
        Exit Sub
    End If

    Call TEC_Modifie_Ligne

End Sub

Private Sub cmdDelete_Click()

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
               Title:="", _
               Buttons:=vbCritical
        Exit Sub
    End If
    
    Call TEC_Efface_Ligne

End Sub

'****************************************** Get a row and display it in the form
Sub lsbHresJour_dblClick(ByVal Cancel As MSForms.ReturnBoolean)

    rmv_state = rmv_modeAffichage
    
    With ufSaisieHeures
        Dim tecID As Long
        tecID = .lsbHresJour.List(.lsbHresJour.ListIndex, 0)
        wshAdmin.Range("TEC_Current_ID").value = tecID
        txtTEC_ID = tecID
        
        'Retrieve the record in wshTEC_Local
        Dim lookupRange As Range, lastTECRow As Long, rowTecID As Long
        lastTECRow = wshTEC_Local.Range("A99999").End(xlUp).row
        Set lookupRange = wshTEC_Local.Range("A3:A" & lastTECRow)
        rowTecID = Get_TEC_Row_Number_By_TEC_ID(tecID, lookupRange)
        
        Dim isBilled As Boolean
        isBilled = wshTEC_Local.Range("L" & rowTecID).value

        'Has this charge beeing INVOICED ?
        If isBilled Then
            MsgBox "Il est impossible de modifier ou de détruire" & vbNewLine & _
                        vbNewLine & "une charge déjà FACTURÉE", vbExclamation
            GoTo exit_sub
        End If
        
        .cmbProfessionnel.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 1)
        .cmbProfessionnel.Enabled = False

        .txtDate.value = Format(.lsbHresJour.List(.lsbHresJour.ListIndex, 2), "dd/mm/yyyy")
        .txtDate.Enabled = False

        .txtClient.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 3)
        savedClient = .txtClient.value
        .txtSavedClient.value = .txtClient.value
        wshAdmin.Range("TEC_Client_ID").value = GetID_From_Client_Name(savedClient)

        .txtActivite.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 4)
        savedActivite = .txtActivite.value
        .txtSavedActivite.value = .txtActivite.value

        .txtHeures.value = Format(.lsbHresJour.List(.lsbHresJour.ListIndex, 5), "#0.00")
        savedHeures = .txtHeures.value
        .txtSavedHeures.value = .txtHeures.value

        .txtCommNote.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 6)
        savedCommNote = .txtCommNote.value
        .txtSavedCommNote.value = .txtCommNote.value

        .chbFacturable.value = CBool(.lsbHresJour.List(.lsbHresJour.ListIndex, 7))
        savedFacturable = .chbFacturable.value
        .txtSavedFacturable.value = .chbFacturable.value
    End With

exit_sub:

    Call Buttons_Enabled_True_Or_False(True, False, False, True)
    
    rmv_state = rmv_modeModification
    
    Set lookupRange = Nothing
    
End Sub

'Public Sub InitializeListBoxClass()
'
'    Set MyListBoxClass = New clsListboxAlign
'
'End Sub

'Sub CopyRangeToListBoxWithoutRowSource()
'    Dim ws As Worksheet: Set ws = wshTEC_Local
'    Dim rng As Range: Set rng = wshTEC_Local("Y2:AL6")
'    Dim lb As Object: Set lb = ufSaisieHeures.lsbHresJour
'    Dim cell As Range
'
'    'Clear any existing items in the ListBox
'    lb.Clear
'
'    'Copy the range values to the ListBox, excluding rows based on the condition
'    For Each cell In rng
'        If cell.Offset(0, 11).value <> "VRAI" Then
'            lb.AddItem cell.value
'        End If
'    Next cell
'End Sub

