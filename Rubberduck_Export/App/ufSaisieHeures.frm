VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   9255.001
   ClientLeft      =   135
   ClientTop       =   570
   ClientWidth     =   15555
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

Sub UserForm_Activate() '2024-07-31 @ 07:57

    'Special timer for log purpose
    Dim timer3Start As Double: timer3Start = Timer: Call Start_Timer("ufSaisieHeures:UserForm_Activate()")
    
    Call Client_List_Import_All
    
    Call TEC_Import_All
    
    Dim lastUsedRow As Long
    lastUsedRow = wshBD_Clients.Range("A9999").End(xlUp).Row
    ufSaisieHeures.ListData = wshBD_Clients.Range("A1:J" & lastUsedRow)
    
    With oEventHandler
        Set .SearchListBox = lstboxNomClient
        Set .SearchTextBox = txtClient
        .MaxRows = 10
        .ShowAllMatches = False
        .CompareMethod = vbTextCompare
    End With

    Call Buttons_Enabled_True_Or_False(False, False, False, False)

    'Default Professionnal - 2024-08-19 @ 07:59
    Dim userName As String
    userName = Fn_Get_Windows_Username
    Select Case userName
        Case "Guillaume", "GuillaumeCharron"
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
    
    cmbProfessionnel.SetFocus
   
    rmv_state = rmv_modeInitial
    
    'Close the special timer (log purpose)
    Call End_Timer("ufSaisieHeures:UserForm_Activate()", timer3Start)
    
End Sub

Private Sub lstboxNomClient_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:lstboxNomClient_DblClick()")
    
    Dim i As Long
    With Me.lstboxNomClient
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Me.txtClient.value = .List(i, 0)
                wshAdmin.Range("TEC_Client_ID").value = Fn_GetID_From_Client_Name(Me.txtClient.value)
                Exit For
            End If
        Next i
    End With
    
    Call End_Timer("ufSaisieHeures:lstboxNomClient_DblClick()", timerStart)

End Sub

Private Sub UserForm_Initialize()

    Set MyListBoxClass = New clsListboxAlign
    
End Sub

Private Sub UserForm_Terminate()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:UserForm_Terminate()")

    'Clear the admin control cells
    wshAdmin.Range("B3:B7").ClearContents
    
    ThisWorkbook.Save
    
    'Clean up
    Set oEventHandler = Nothing
    
    ufSaisieHeures.Hide
    Unload ufSaisieHeures
    
    If ufSaisieHeures.name = "ufSaisieHeures" Then
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

    Call End_Timer("ufSaisieHeures:UserForm_Terminate()", timerStart)

End Sub

Public Sub cmbProfessionnel_AfterUpdate()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:cmbProfessionnel_AfterUpdate()")

    If ufSaisieHeures.cmbProfessionnel.value <> "" Then
        wshAdmin.Range("TEC_Initials").value = ufSaisieHeures.cmbProfessionnel.value
        wshAdmin.Range("TEC_Prof_ID").value = Fn_GetID_From_Initials(ufSaisieHeures.cmbProfessionnel.value)
        
        If wshAdmin.Range("TEC_Date").value <> "" Then
            ufSaisieHeures.txtDate.value = wshAdmin.Range("TEC_Date").value
            Call TEC_AdvancedFilter_And_Sort
            Call TEC_Refresh_ListBox_And_Add_Hours
        End If
    End If

    Call End_Timer("ufSaisieHeures:cmbProfessionnel_AfterUpdate()", timerStart)

End Sub

Private Sub txtDate_Enter()

    If ufSaisieHeures.txtDate.value = "" Then
        ufSaisieHeures.txtDate.value = Format$(Now(), "dd/mm/yyyy")
    End If

End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:txtDate_BeforeUpdate()")
    
    Dim fullDate As Variant
    
    fullDate = CompleteDate(CStr(ufSaisieHeures.txtDate.value))
        
    'Update the cell with the full date, if valid
    If fullDate <> "Invalid Date" Then
        ufSaisieHeures.txtDate.value = fullDate
    Else
        Call MsgBoxInvalidDate
        Application.EnableEvents = False
        ufSaisieHeures.txtDate.value = ""
        Application.EnableEvents = True
        txtDate.SelStart = 0
        txtDate.SelLength = Len(Me.txtDate.value)
        txtDate.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If CDate(fullDate) > Format$(Now(), "dd/mm/yyyy") Then
        If MsgBox("En êtes-vous CERTAIN ?", vbYesNo + vbQuestion, "Utilisation d'une date FUTURE") = vbNo Then
            txtDate.SelStart = 0
            txtDate.SelLength = Len(Me.txtDate.value)
            txtDate.SetFocus
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = False
    
    Call End_Timer("ufSaisieHeures:txtDate_BeforeUpdate()", timerStart)
    
End Sub

Private Sub txtDate_AfterUpdate()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:txtDate_AfterUpdate()")
    
    If IsDate(ufSaisieHeures.txtDate.value) Then
        wshAdmin.Range("TEC_Date").value = CDate(ufSaisieHeures.txtDate.value)
    Else
        ufSaisieHeures.txtDate.SetFocus
        ufSaisieHeures.txtDate.SelLength = Len(ufSaisieHeures.txtDate.value)
        ufSaisieHeures.txtDate.SelStart = 0
        Exit Sub
    End If

    If wshAdmin.Range("TEC_Prof_ID").value <> "" Then
        Call TEC_AdvancedFilter_And_Sort
        Call TEC_Refresh_ListBox_And_Add_Hours
    End If
    
    'Enabled the NEW & ADD button if the minimum fields are non empty
    If Trim(ufSaisieHeures.cmbProfessionnel.value) <> vbNullString And _
        Trim(ufSaisieHeures.txtDate.value) <> vbNullString And _
        Trim(ufSaisieHeures.txtClient.value) <> vbNullString And _
        Trim(ufSaisieHeures.txtHeures.value) <> 0 Then
        Call Buttons_Enabled_True_Or_False(True, True, False, False)
    End If
    
    Call End_Timer("ufSaisieHeures:txtDate_AfterUpdate()", timerStart)
    
End Sub

Private Sub txtClient_Enter()

    Call SetNumLockOn '2024-08-26 @ 09:55
    
    If rmv_state = rmv_modeInitial Then
        rmv_state = rmv_modeCreation
    End If

End Sub

Private Sub txtClient_AfterUpdate()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:txtClient_AfterUpdate()")
    
    If Me.txtClient.value <> Me.txtSavedClient.value Then
'        If Me.txtTEC_ID.value = "" Then
            Call Buttons_Enabled_True_Or_False(True, False, False, False)
'        Else
'            Call Buttons_Enabled_True_Or_False(True, False, False, False)
'        End If
    End If
    
    Call End_Timer("ufSaisieHeures:txtClient_AfterUpdate()", timerStart)
    
End Sub

Private Sub txtActivite_AfterUpdate()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:txtActivite_AfterUpdate()")
    
    If Me.txtActivite.value <> Me.txtSavedActivite.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, False, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If

    Call End_Timer("ufSaisieHeures:txtActivite_AfterUpdate()", timerStart)
    
End Sub

Sub txtHeures_AfterUpdate()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:txtHeures_AfterUpdate()")
    
    'Validation des heures saisies
    Dim strHeures As String
    strHeures = Me.txtHeures.value
    
    If InStr(strHeures, ".") > 0 Then
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
    
    Me.txtHeures.value = Format$(strHeures, "#0.00")
    
    If Me.txtHeures.value <> Me.txtSavedHeures.value Then
        Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate - is '" & Me.txtHeures.value & "' <> '" & Me.txtSavedHeures.value & "' ?", -1)
        If Me.txtTEC_ID = "" Then
        Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate - Me.txtTEC_ID is Empty '" & Me.txtTEC_ID & "', alors True, True, False, False", -1)
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        Else
        Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate - Me.txtTEC_ID is NOT Empty '" & Me.txtTEC_ID & "' , alors True, True, False, False", -1)
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If
    
'ufSaisieHeures:txtHeures_AfterUpdate - ? 0,30 <>  (sortie)|Temps écoulé: 34195,6016 seconds
'ufSaisieHeures:txtHeures_AfterUpdate - ? Me.txtTEC_ID = " (sortie)|Temps écoulé: 34195,6055 seconds

    Call End_Timer("ufSaisieHeures:txtHeures_AfterUpdate()", timerStart)
    
End Sub

Private Sub chbFacturable_AfterUpdate()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:chbFacturable_AfterUpdate()")
    
    If Me.chbFacturable.value <> Me.txtSavedFacturable.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If

    Call End_Timer("ufSaisieHeures:chbFacturable_AfterUpdate()", timerStart)
    
End Sub

Private Sub txtCommNote_AfterUpdate()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("ufSaisieHeures:txtCommNote_AfterUpdate()")
    
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

    Call End_Timer("ufSaisieHeures:txtCommNote_AfterUpdate()", timerStart)
    
End Sub

'----------------------------------------------------------------- ButtonsEvents
Private Sub cmdClear_Click()

    Call TEC_Efface_Formulaire

End Sub

Private Sub cmdAdd_Click()

    Call TEC_Ajoute_Ligne

End Sub

Private Sub cmdUpdate_Click()

    If wshAdmin.Range("TEC_Current_ID").value <> "" Then
        Call TEC_Modifie_Ligne
    Else
        MsgBox Prompt:="Vous devez choisir un enregistrement à modifier !", _
               Title:="", _
               Buttons:=vbCritical
    End If

End Sub

Private Sub cmdDelete_Click()

    If wshAdmin.Range("TEC_Current_ID").value <> "" Then
        Call TEC_Efface_Ligne
    Else
        MsgBox Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
               Title:="", _
               Buttons:=vbCritical
    End If

End Sub

'Get a specific row from listBox and display it in the userform
Sub lsbHresJour_dblClick(ByVal Cancel As MSForms.ReturnBoolean)

    rmv_state = rmv_modeAffichage
    
    With ufSaisieHeures
        Dim TECID As Long
        TECID = .lsbHresJour.List(.lsbHresJour.ListIndex, 0)
        wshAdmin.Range("TEC_Current_ID").value = TECID
        txtTEC_ID = TECID
        
        'Retrieve the record in wshTEC_Local
        Dim lookupRange As Range, lastTECRow As Long, rowTecID As Long
        lastTECRow = wshTEC_Local.Range("A99999").End(xlUp).Row
        Set lookupRange = wshTEC_Local.Range("A3:A" & lastTECRow)
        rowTecID = Fn_Find_Row_Number_TEC_ID(TECID, lookupRange)
        
        Dim isBilled As Boolean
        isBilled = wshTEC_Local.Range("L" & rowTecID).value

        'Has this charge beeing INVOICED ?
        If Not isBilled Then
            .cmbProfessionnel.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 1)
            .cmbProfessionnel.Enabled = False
    
            .txtDate.value = Format$(.lsbHresJour.List(.lsbHresJour.ListIndex, 2), "dd/mm/yyyy") '2024-08-10 @ 07:23
            .txtDate.Enabled = False
    
            .txtClient.value = .lsbHresJour.List(.lsbHresJour.ListIndex, 3)
            savedClient = .txtClient.value
            .txtSavedClient.value = .txtClient.value
            wshAdmin.Range("TEC_Client_ID").value = wshTEC_Local.Range("E" & rowTecID).value
    
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
        Else
            MsgBox "Il est impossible de modifier ou de détruire" & vbNewLine & _
                        vbNewLine & "une charge déjà FACTURÉE", vbExclamation
        End If
        
    End With

    Call Buttons_Enabled_True_Or_False(True, False, False, True)
    
    rmv_state = rmv_modeModification
    
    'Cleaning memory - 2024-07-31 2 08:34
    Set lookupRange = Nothing
    
End Sub


