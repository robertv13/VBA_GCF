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
'Private MyListBoxClass As clsListboxAlign

'Allows the calling code to set the data
Public Property Let ListData(ByVal rg As Range)

    oEventHandler.List = rg.value

End Property

Sub UserForm_Activate() '2024-07-31 @ 07:57

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:UserForm_Activate", 0)
    
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
    
    wshAdmin.Range("TEC_Date").value = "" 'On vide la date pour forcer la saisie
    
    cmbProfessionnel.SetFocus
   
    rmv_state = rmv_modeInitial
    
'    Application.SendKeys "{NUMLOCK}", True
    
    Call VerifierNumLockAvecGetKeyboardState '2024-09-03 @ 22:43
    
    Call Log_Record("ufSaisieHeures:UserForm_Activate()", startTime)
    
End Sub

Private Sub lstboxNomClient_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:lstboxNomClient_DblClick", 0)
    
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
    
'    Application.SendKeys "{NUMLOCK}", True
    
    Call VerifierNumLockAvecGetKeyboardState '2024-09-03 @ 22:43

    Call Log_Record("ufSaisieHeures:lstboxNomClient_DblClick()", startTime)

End Sub

Private Sub UserForm_Initialize()

'    Set MyListBoxClass = New clsListboxAlign
    
End Sub

Private Sub UserForm_Terminate()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:UserForm_Terminate", 0)

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
    
    GoTo Exit_sub
    
MenuSelect:
    wshMenu.Activate
    wshMenu.Select
    
Exit_sub:

    Call Log_Record("ufSaisieHeures:UserForm_Terminate()", startTime)

End Sub

Public Sub cmbProfessionnel_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:cmbProfessionnel_AfterUpdate", 0)

    If ufSaisieHeures.cmbProfessionnel.value <> "" Then
        wshAdmin.Range("TEC_Initials").value = ufSaisieHeures.cmbProfessionnel.value
        wshAdmin.Range("TEC_Prof_ID").value = Fn_GetID_From_Initials(ufSaisieHeures.cmbProfessionnel.value)
        
        Call Settrace("@0137", "ufSaisieHeures", "cmbProfessionnel_AfterUpdate", "wshAdmin.Range(""TEC_Date"").value = '" & wshAdmin.Range("TEC_Date").value & "'", "type = " & TypeName(wshAdmin.Range("TEC_Date").value))
        If wshAdmin.Range("TEC_Date").value <> "" Then '2024-09-05 @ 20:50
            ufSaisieHeures.txtDate.value = wshAdmin.Range("TEC_Date").value
            Debug.Print "ufSaisieHeures_cmbProfessionnel_AfterUpdate_136   ufSaisieHeures.txtDate.value = "; ufSaisieHeures.txtDate.value; "   "; TypeName(ufSaisieHeures.txtDate.value)
            Call TEC_Get_All_TEC_AF
            Call TEC_Refresh_ListBox_And_Add_Hours
        End If
    End If

'    Application.SendKeys "{NUMLOCK}", True
    
    Call VerifierNumLockAvecGetKeyboardState '2024-09-03 @ 22:43

    Call Log_Record("ufSaisieHeures:cmbProfessionnel_AfterUpdate()", startTime)

End Sub

Private Sub txtDate_Enter()

    If ufSaisieHeures.txtDate.value = "" Then
        ufSaisieHeures.txtDate.value = Now()
'        ufSaisieHeures.txtDate.value = Format$(Now(), "dd/mm/yyyy") '2024-09-04 @ 11:03
    End If
    
'    Debug.Print "DP-102", year(ufSaisieHeures.txtDate.value), month(ufSaisieHeures.txtDate.value), day(ufSaisieHeures.txtDate.value)

End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtDate_BeforeUpdate", 0)
    
    Dim fullDate As Variant
    
    Call Settrace("@0171", "ufSaisieHeures", "txtDate_BeforeUpdate", "ufSaisieHeures.txtDate.value = '" & ufSaisieHeures.txtDate.value & "'", "type = " & TypeName(ufSaisieHeures.txtDate.value))
    fullDate = Fn_Complete_Date(ufSaisieHeures.txtDate.value)
    Call Settrace("@0173", "ufSaisieHeures", "txtDate_BeforeUpdate", "fullDate = '" & fullDate & "'", "type = " & TypeName(fullDate))
        
    'Update the cell with the full date, if valid
    If fullDate <> "Invalid Date" Then
        ufSaisieHeures.txtDate.value = fullDate
    Else
        Call MsgBoxInvalidDate("ufSaisieHeures_179")
        Cancel = True
        With ufSaisieHeures.txtDate
            .SetFocus 'Remettre le focus sur la TextBox
            .SelStart = 0 'Début de la sélection
            .SelLength = Len(.Text) 'Sélectionner tout le texte
        End With
        Exit Sub
    End If
    
    Call Settrace("@00189", "ufSaisieHeures", "txtDate_BeforeUpdate", "fullDate = '" & fullDate & "'", "type = " & TypeName(fullDate))
    Call Settrace("@00190", "ufSaisieHeures", "txtDate_BeforeUpdate", "DateSerial = '" & DateSerial(year(Now), month(Now), day(Now)) & "'", "Type = " & TypeName(DateSerial(year(Now), month(Now), day(Now))) & "   Future ? " & fullDate > DateSerial(year(Now), month(Now), day(Now)))
    If fullDate > DateSerial(year(Now), month(Now), day(Now)) Then
        Debug.Print "ufSaisieHeures:BeforeUpdate @00192 - fullDate = "; fullDate; "   "; TypeName(fullDate)
        If MsgBox("En êtes-vous CERTAIN de vouloir cette date ?" & vbNewLine & vbNewLine & _
                    "La date saisie est '" & fullDate & "'", vbYesNo + vbQuestion, _
                    "Utilisation d'une date FUTURE") = vbNo Then
            txtDate.SelStart = 0
            txtDate.SelLength = Len(Me.txtDate.value)
            txtDate.SetFocus
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = False
    
    Call Log_Record("ufSaisieHeures:txtDate_BeforeUpdate()", startTime)
    
End Sub

Private Sub txtDate_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtDate_AfterUpdate", 0)
    
    Call Settrace("@0214", "ufSaisieHeures", "txtDate_AfterUpdate", "ufSaisieHeures.txtDate.value = '" & ufSaisieHeures.txtDate.value & "'", "type = " & TypeName(ufSaisieHeures.txtDate.value))
'    MsgBox "ufSaisieHeures_txtDate_AfterUpdate_208   " & vbNewLine & vbNewLine & _
                "y = " & year(ufSaisieHeures.txtDate.value) & vbNewLine & vbNewLine & _
                "m = " & month(ufSaisieHeures.txtDate.value) & vbNewLine & vbNewLine & _
                "d = " & day(ufSaisieHeures.txtDate.value), vbInformation, "Format détaillé de la date"
    If IsDate(ufSaisieHeures.txtDate.value) Then
        Dim dateStr As String, dateFormated As Date
        dateStr = ufSaisieHeures.txtDate.value
        Call Settrace("@0222", "ufSaisieHeures", "txtDate_AfterUpdate", "dateStr = '" & dateStr & "'", "type = " & TypeName(dateStr))
        dateFormated = DateSerial(year(dateStr), month(dateStr), day(dateStr))
        Call Settrace("@0224", "ufSaisieHeures", "txtDate_AfterUpdate", "dateFormated = '" & dateFormated & "'", "type = " & TypeName(dateFormated))
        wshAdmin.Range("TEC_Date").value = dateFormated
        Call Settrace("@0226", "ufSaisieHeures", "txtDate_AfterUpdate", "wshAdmin.Range(""TEC_Date"").value = '" & wshAdmin.Range("TEC_Date").value & "'", "type = " & TypeName(wshAdmin.Range("TEC_Date").value))
    Else
        ufSaisieHeures.txtDate.SetFocus
        ufSaisieHeures.txtDate.SelLength = Len(ufSaisieHeures.txtDate.value)
        ufSaisieHeures.txtDate.SelStart = 0
        Exit Sub
    End If

    If wshAdmin.Range("TEC_Prof_ID").value <> "" Then
        Call TEC_Get_All_TEC_AF
        Call TEC_Refresh_ListBox_And_Add_Hours
    End If
    
    'Enabled the NEW & ADD button if the minimum fields are non empty
    If Trim(ufSaisieHeures.cmbProfessionnel.value) <> vbNullString And _
        Trim(ufSaisieHeures.txtDate.value) <> vbNullString And _
        Trim(ufSaisieHeures.txtClient.value) <> vbNullString And _
        Trim(ufSaisieHeures.txtHeures.value) <> 0 Then
        Call Buttons_Enabled_True_Or_False(True, True, False, False)
    End If
    
    Call Log_Record("ufSaisieHeures:txtDate_AfterUpdate()", startTime)
    
End Sub

Private Sub txtClient_Enter()

    Call SetNumLockOn '2024-08-26 @ 09:55
    
    If rmv_state = rmv_modeInitial Then
        rmv_state = rmv_modeCreation
    End If

End Sub

Private Sub txtClient_AfterUpdate()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtClient_AfterUpdate", 0)
    
    Call SetNumLockOn '2024-08-26 @ 09:55
    
    If Me.txtClient.value <> Me.txtSavedClient.value Then
'        If Me.txtTEC_ID.value = "" Then
            Call Buttons_Enabled_True_Or_False(True, False, False, False)
'        Else
'            Call Buttons_Enabled_True_Or_False(True, False, False, False)
'        End If
    End If
    
'    Application.SendKeys "{NUMLOCK}", True
    
    Call VerifierNumLockAvecGetKeyboardState
    
    Call Log_Record("ufSaisieHeures:txtClient_AfterUpdate()", startTime)
    
End Sub

Private Sub txtActivite_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtActivite_AfterUpdate", 0)
    
    If Me.txtActivite.value <> Me.txtSavedActivite.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, False, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If

'    Application.SendKeys "{NUMLOCK}", True
    
    Call VerifierNumLockAvecGetKeyboardState
    
    Call Log_Record("ufSaisieHeures:txtActivite_AfterUpdate()", startTime)
    
End Sub

Sub txtHeures_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate", 0)
    
    'Validation des heures saisies
    Dim strHeures As String
    strHeures = Me.txtHeures.value
    
    strHeures = Replace(strHeures, ".", ",")
    
    If IsNumeric(strHeures) = False Then
        MsgBox Prompt:="La valeur saisie ne peut être utilisée comme valeur numérique!", _
                Title:="Validation d'une valeur numérique", _
                Buttons:=vbCritical
            Me.txtHeures.SetFocus 'Mettre le focus sur le contrôle en premier
            Exit Sub
    End If
    
    If strHeures > 24 Then
        MsgBox _
        Prompt:="Le nombre d'heures pour une entrée ne peut dépasser 24!", _
        Title:="Validation d'une valeur numérique", _
        Buttons:=vbCritical
        Me.txtHeures.SetFocus
        Me.txtHeures.SelLength = Len(ufSaisieHeures.txtHeures.value)
        Me.txtHeures.SelStart = 0
        Exit Sub
    End If
    Me.txtHeures.value = Format$(strHeures, "#0.00")
    
    If Me.txtHeures.value <> Me.txtSavedHeures.value Then
        Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate - is '" & Me.txtHeures.value & "' <> '" & Me.txtSavedHeures.value & "' ?", -1)
        If Me.txtTEC_ID = "" Then
            Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate - Me.txtTEC_ID is Empty '" & Me.txtTEC_ID & "', alors True, True, False, False", -1)
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        Else
            Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate - Me.txtTEC_ID is NOT Empty '" & Me.txtTEC_ID & "' , alors True, False, True, True", -1)
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If
    
    Call Log_Record("ufSaisieHeures:txtHeures_AfterUpdate()", startTime)
    
End Sub

Private Sub chbFacturable_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:chbFacturable_AfterUpdate", 0)
    
    If Me.chbFacturable.value <> Me.txtSavedFacturable.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If

    Call Log_Record("ufSaisieHeures:chbFacturable_AfterUpdate()", startTime)
    
End Sub

Private Sub txtCommNote_AfterUpdate()

    Dim startTime As Double: startTime = Timer: Call Log_Record("ufSaisieHeures:txtCommNote_AfterUpdate", 0)
    
    If Me.txtCommNote.value <> Me.txtSavedCommNote.value Then
        If Me.txtTEC_ID = "" Then
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        Else
            Call Buttons_Enabled_True_Or_False(True, False, True, True)
        End If
    End If

    Call Log_Record("ufSaisieHeures:txtCommNote_AfterUpdate()", startTime)
    
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

'Sub VerifierNumLock()
'
'    Const VK_NUMLOCK As Long = &H90 ' Code de la touche NUM LOCK
'
'    ' Appel de l'API Windows pour obtenir l'état de NUM LOCK
'    If GetKeyState(VK_NUMLOCK) And 1 Then
'        MsgBox "Le NUM LOCK est activé.", vbInformation
'    Else
'        MsgBox "Le NUM LOCK est désactivé.", vbInformation
'    End If
'
'End Sub
