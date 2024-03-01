VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   8550.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13950
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

Private Sub lstNomClient_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim timerStart As Double: timerStart = Timer
    
    Dim i As Long
    With Me.lstNomClient
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Me.txtClient.value = .List(i, 0)
                wshAdmin.Range("TEC_Client_ID").value = GetID_From_Client_Name(Me.txtClient.value)
                Exit For
            End If
        Next i
    End With
    
    Call Output_Timer_Results("lstNomClient_DblClick()", timerStart)

End Sub

'******************************************* Execute when UserForm is displayed
Sub UserForm_Activate()

    Dim timer3Start As Double: timer3Start = Timer
    
    Call Client_List_Import_All
    
    Dim lastUsedRow As Long
    lastUsedRow = wshClientDB.Range("A9999").End(xlUp).row
    ufSaisieHeures.ListData = wshClientDB.Range("A1:J" & lastUsedRow)
    
    With oEventHandler
        Set .SearchListBox = lstNomClient
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

    Dim MyListBoxClass As clsCListboxAlign
    Set MyListBoxClass = New clsCListboxAlign 'declare the class

    'Corrige le format des colonnes (Left, Center & Right)
'    MyListBoxClass.Right Me.ListBox2, 1
'    MyListBoxClass.Right Me.ListBox2, 2
'    MyListBoxClass.Right Me.ListBox2, 3
'    MyListBoxClass.Right Me.ListBox2, 4
'    MyListBoxClass.Right Me.ListBox2, 5
'    MyListBoxClass.Right Me.ListBox2, 6
'    MyListBoxClass.Right Me.ListBox2, 7
'    MyListBoxClass.Right Me.ListBox2, 8
'    MyListBoxClass.Right Me.ListBox2, 9

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
    
    'Enabled the NEW & ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> "" And _
        Trim(Me.txtDate.value) <> "" And _
        Trim(Me.txtClient.value) <> "" And _
        Trim(Me.txtHeures.value) <> "" Then
        Call Buttons_Enabled_True_Or_False(True, True, False, False)
    End If

exit_sub:

    Call Output_Timer_Results("ufSaisieHeures - cmbProfessionnel_AfterUpdate()", timerStart)

End Sub

Private Sub txtDate_Enter()

    If txtDate.value = vbNullString Then
        txtDate.value = Format(CDate(Now()), "dd/mm/yyyy")
    End If
    
    'Enabled the NEW & ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> "" And _
        Trim(Me.txtDate.value) <> "" And _
        Trim(Me.txtClient.value) <> "" And _
        Trim(Me.txtHeures.value) <> "" Then
        Call Buttons_Enabled_True_Or_False(True, True, False, False)
    End If

End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If txtDate.value = vbNullString Then
        txtDate.value = Format(CDate(Now()), "dd/mm/yyyy")
    End If

    Dim strDate As String
    strDate = txtDate.value

    Dim separateur As String
    separateur = "/"

    Dim currentYear As Integer, currentMonth As Integer, currentDay As Integer
    currentYear = Format(Year(Now()), "0000")
    currentMonth = Format(Month(Now()), "00")
    currentDay = Format(Day(Now()), "00")

    If Len(strDate) <= 2 Then
        strDate = Format(strDate, "00") & separateur & currentMonth & _
                  separateur & currentYear
    ElseIf Len(strDate) = 5 Then
        strDate = strDate & separateur & currentYear
    End If

    'Validation de la date
    If IsDate(strDate) = False Then
        MsgBox Prompt:="La valeur saisie ne peut être utilisée comme une date valide!", _
            Title:="Validation de la date", _
            Buttons:=vbCritical
            txtDate.SelStart = 0
            txtDate.SelLength = Len(txtDate.value)
        Exit Sub
    End If

    If Not IsDate(strDate) Then
        txtDateShowError ("Dates seulement!")
        Cancel = True
    Else
        If CDate(strDate) > Date Then
            Call txtDateShowError("Pas de date future!")
            Cancel = True
        End If
    End If
    
    txtDate.value = strDate

End Sub

Private Sub txtDateShowError(ErrorCaption As String)

    txtDate.BackColor = rgbPink
    lblDate.ForeColor = rgbRed
    lblDate.Caption = ErrorCaption
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.value)

End Sub

Private Sub txtDate_AfterUpdate()

    txtDate.BackColor = rgbWhite
    lblDate.Caption = "Date *"
    lblDate.ForeColor = Me.ForeColor
    
    wshAdmin.Range("TEC_Date").value = CDate(Me.txtDate.value)

    If wshAdmin.Range("TEC_Prof_ID").value <> "" Then
        Call TEC_AdvancedFilter_And_Sort
        Call Refresh_ListBox_And_Add_Hours
    End If
    
    'Enabled the NEW & ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> vbNullString And _
        Trim(Me.txtDate.value) <> vbNullString And _
        Trim(Me.txtClient.value) <> vbNullString And _
        Trim(Me.txtHeures.value) <> vbNullString Then
        Call Buttons_Enabled_True_Or_False(True, True, False, False)
    End If
    
End Sub

Private Sub txtClient_Enter()

    If rmv_state = rmv_modeInitial Then
        rmv_state = rmv_modeCreation
    End If

End Sub

Sub txtClient_AfterUpdate()
    
    'Enabled the ADD button if the minimum fields are non empty
    If rmv_state = rmv_modeCreation Then
        If Trim(Me.cmbProfessionnel.value) <> "" And _
            Trim(Me.txtDate.value) <> "" And _
            Trim(Me.txtClient.value) <> "" And _
            Trim(Me.txtHeures.value) <> "" Then
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        End If
    ElseIf rmv_state = rmv_modeAffichage Then
        If savedClient <> Me.txtClient.value Or _
            savedActivite <> Me.txtActivite.value Or _
            savedHeures <> Me.txtHeures.value Or _
            savedCommNote <> Me.txtCommNote Or _
            savedFacturable <> Me.chbFacturable Then
            Call Buttons_Enabled_True_Or_False(False, False, True, True)
        End If
    End If
    
End Sub

Private Sub txtActivite_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If txtActivite.value <> savedActivite Then
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
    
    'Enabled the NEW & ADD button if the minimum fields are non empty
    If rmv_state = rmv_modeCreation Then
        If Trim(Me.cmbProfessionnel.value) <> "" And _
            Trim(Me.txtDate.value) <> "" And _
            Trim(Me.txtClient.value) <> "" And _
            Trim(Me.txtHeures.value) <> "" Then
            Call Buttons_Enabled_True_Or_False(True, True, False, False)
        End If
    End If

    If rmv_state = rmv_modeAffichage Then
        If Me.txtHeures.value <> savedHeures Then
            Call Buttons_Enabled_True_Or_False(True, False, True, False)
        End If
    End If
    
End Sub

Private Sub chbFacturable_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If Me.chbFacturable.value <> savedFacturable Then
            Call Buttons_Enabled_True_Or_False(True, False, True, False)
        End If
    End If

End Sub

Private Sub txtCommNote_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If Me.txtCommNote.value <> savedCommNote Then
            Call Buttons_Enabled_True_Or_False(True, False, True, False)
        End If
    End If

End Sub

'----------------------------------------------------------------- ButtonsEvents
Private Sub cmdClear_Click()

    TEC_Efface_Formulaire

End Sub

Private Sub cmdAdd_Click()

    Call TEC_Ajoute_Ligne_Detail

End Sub

Private Sub cmdUpdate_Click()

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à modifier !", _
               Title:="", _
               Buttons:=vbCritical
        Exit Sub
    End If

    Call TEC_Modifie_Ligne_Detail

End Sub

Private Sub cmdDelete_Click()

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
               Title:="", _
               Buttons:=vbCritical
        Exit Sub
    End If
    
    TEC_Efface_Ligne_Detail

End Sub

'****************************************** Get a row and display it in the form
Sub ListBox2_dblClick(ByVal Cancel As MSForms.ReturnBoolean)

    rmv_state = rmv_modeAffichage
'    Stop
'    ufSaisieHeures.ListBox2.ColumnWidths = "35; 30; 55; 130; 180; 35; 80; 40; 85"
    
    With ufSaisieHeures
        Dim tecID As Long
        tecID = .ListBox2.List(.ListBox2.ListIndex, 0)
        wshAdmin.Range("TEC_Current_ID").value = tecID
        
        'Retrieve the record in wshBaseHours
        Dim lookupRange As Range, lastTECRow As Long, rowTecID As Long
        lastTECRow = wshBaseHours.Range("A99999").End(xlUp).row
        Set lookupRange = wshBaseHours.Range("A3:A" & lastTECRow)
        rowTecID = Get_TEC_Row_Number_By_TEC_ID(tecID, lookupRange)
        
        Dim isBilled As Boolean
        isBilled = wshBaseHours.Range("L" & rowTecID).value

        'Has this charge beeing INVOICED ?
        If isBilled Then
            MsgBox "Il est impossible de modifier ou de détruire" & vbNewLine & _
                        vbNewLine & "une charge déjà FACTURÉE", vbExclamation
            GoTo exit_sub
        End If
        
        .cmbProfessionnel.value = .ListBox2.List(.ListBox2.ListIndex, 1)
        .cmbProfessionnel.Enabled = False

        .txtDate.value = Format(.ListBox2.List(.ListBox2.ListIndex, 2), "dd/mm/yyyy")
        .txtDate.Enabled = False

        .txtClient.value = .ListBox2.List(.ListBox2.ListIndex, 3)
        savedClient = .txtClient.value
        wshAdmin.Range("TEC_Client_ID").value = GetID_From_Client_Name(savedClient)

        .txtActivite.value = .ListBox2.List(.ListBox2.ListIndex, 4)
        savedActivite = .txtActivite.value

        .txtHeures.value = Format(.ListBox2.List(.ListBox2.ListIndex, 5), "#0.00")
        savedHeures = .txtHeures.value

        .txtCommNote.value = .ListBox2.List(.ListBox2.ListIndex, 6)
        savedCommNote = .txtCommNote.value

        .chbFacturable.value = CBool(.ListBox2.List(.ListBox2.ListIndex, 7))
        savedFacturable = .chbFacturable.value
    End With

exit_sub:

    Call Buttons_Enabled_True_Or_False(True, False, False, True)
    
    rmv_state = rmv_modeModification
'    rmv_state = rmv_modeAffichage
    
    Set lookupRange = Nothing
    
End Sub

Sub Buttons_Enabled_True_Or_False(clear As Boolean, add As Boolean, _
                                  update As Boolean, delete As Boolean)
    With ufSaisieHeures
        .cmdClear.Enabled = clear
        .cmdAdd.Enabled = add
        .cmdUpdate.Enabled = update
        .cmdDelete.Enabled = delete
    End With

End Sub

'Sub CopyRangeToListBoxWithoutRowSource()
'    Dim ws As Worksheet: Set ws = wshBaseHours
'    Dim rng As Range: Set rng = wshBaseHours("Y2:AL6")
'    Dim lb As Object: Set lb = ufSaisieHeures.ListBox2
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

