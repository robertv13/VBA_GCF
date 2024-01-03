VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   9525.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13275
   OleObjectBlob   =   "frmSaisieHeures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSaisieHeures"
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

Private Sub icoGraph_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

    MsgBox "Affichage des statistiques à compléter"

End Sub

Private Sub lstNomClient_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim i As Long
    With Me.lstNomClient
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Me.txtClient.value = .List(i, 0)
                wshAdmin.Range("TEC_Client_ID").value = GetID_FromClientName(Me.txtClient.value)
                Exit For
            End If
        Next i
    End With

End Sub

'******************************************* Execute when UserForm is displayed
Sub UserForm_Activate()

    Call ImportClientList
    
    Dim lastUsedRow As Long
    lastUsedRow = wshClientDB.Range("A999999").End(xlUp).row
    frmSaisieHeures.ListData = wshClientDB.Range("B1:B" & lastUsedRow)
    
    With oEventHandler
        Set .SearchListBox = lstNomClient
        Set .SearchTextBox = txtClient
        .MaxRows = 10
        .ShowAllMatches = False
        .CompareMethod = vbTextCompare
    End With

    cmbProfessionnel.value = ""
    cmbProfessionnel.SetFocus
   
    rmv_state = rmv_modeInitial
   
End Sub

Private Sub UserForm_Terminate()
    
    'Clear the admin control cells
    wshAdmin.Range("B3:B7").ClearContents
    
    ThisWorkbook.Save
    
    'Clean up
    Set oEventHandler = Nothing
    
    Me.Hide
    Unload Me
    
    If Me.Name = "frmSaisieHeures" Then
        On Error GoTo MenuSelect
        wshMenuTEC.Select
        On Error GoTo 0
    Else
        wshMenu.Select
    End If
    Exit Sub
    
MenuSelect:
    wshMenu.Select
    
End Sub

Public Sub cmbProfessionnel_AfterUpdate()

    wshAdmin.Range("TEC_Initials").value = Me.cmbProfessionnel.value
    wshAdmin.Range("TEC_Prof_ID").value = GetID_FromInitials(Me.cmbProfessionnel.value)
    
    If wshAdmin.Range("TEC_Date").value <> "" Then
        Call TEC_FilterAndSort
        Call RefreshListBoxAndAddHours
    End If
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> "" And _
        Trim(Me.txtDate.value) <> "" And _
        Trim(Me.txtClient.value) <> "" And _
        Trim(Me.txtHeures.value) <> "" Then
        cmdAdd.Enabled = True
    End If
    
End Sub

Private Sub txtDate_Enter()

    If txtDate.value = vbNullString Then
        txtDate.value = Format(CDate(Now()), "dd/mm/yyyy")
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
        Call TEC_FilterAndSort
        Call RefreshListBoxAndAddHours
    End If
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> vbNullString And _
        Trim(Me.txtDate.value) <> vbNullString And _
        Trim(Me.txtClient.value) <> vbNullString And _
        Trim(Me.txtHeures.value) <> vbNullString Then
        cmdAdd.Enabled = True
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
            cmdAdd.Enabled = True
        End If
    End If
    
    If rmv_state = rmv_modeAffichage Then
        If savedClient <> Me.txtClient.value Or _
            savedActivite <> Me.txtActivite.value Or _
            savedHeures <> Me.txtHeures.value Or _
            savedCommNote <> Me.txtCommNote Or _
            savedFacturable <> Me.chbFacturable Then
            cmdUpdate.Enabled = True
        End If
    End If
    
    'Debug.Print rmv_state
    
    If rmv_state = rmv_modeAffichage Then
        If Me.txtClient.value <> savedClient Then
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = True
        End If
    End If
    
End Sub

Private Sub txtActivite_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If txtActivite.value <> savedActivite Then
            cmdUpdate.Enabled = True
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
    
    If IsNumeric(strHeures) = False Then
        MsgBox _
        Prompt:="La valeur saisie ne peut être utilisée comme valeur numérique!", _
        Title:="Validation d'une valeur numérique", _
        Buttons:=vbCritical
        Me.txtHeures.value = ""
        Me.txtHeures.SetFocus
        Exit Sub
    End If
    
    Me.txtHeures.value = Format(strHeures, "#0.00")
    
    'Enabled the ADD button if the minimum fields are non empty
    If rmv_state = rmv_modeCreation Then
        If Trim(Me.cmbProfessionnel.value) <> "" And _
           Trim(Me.txtDate.value) <> "" And _
           Trim(Me.txtClient.value) <> "" And _
           Trim(Me.txtHeures.value) <> "" Then
            cmdAdd.Enabled = True
        End If
    End If

    If rmv_state = rmv_modeAffichage Then
        If Me.txtHeures.value <> savedHeures Then
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = True
        End If
    End If
    
End Sub

Private Sub chbFacturable_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If Me.chbFacturable.value <> savedFacturable Then
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = True
        End If
    End If

End Sub

Private Sub txtCommNote_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If Me.txtCommNote.value <> savedCommNote Then
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = True
        End If
    End If

End Sub

'----------------------------------------------------------------- ButtonsEvents
Private Sub cmdClear_Click()

    EffaceFormulaire

End Sub

Private Sub cmdAdd_Click()

    AjouteLigneDetail

End Sub

Private Sub cmdUpdate_Click()

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à modifier !", _
               Title:="", _
               Buttons:=vbCritical
        Exit Sub
    End If

    ModifieLigneDetail

End Sub

Private Sub cmdDelete_Click()

    If wshAdmin.Range("TEC_Current_ID").value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
               Title:="", _
               Buttons:=vbCritical
        Exit Sub
    End If
    
    EffaceLigneDetail

End Sub

'****************************************** Get a row and display it in the form
Sub lstData_dblClick(ByVal Cancel As MSForms.ReturnBoolean)

    rmv_state = rmv_modeAffichage
    
    With frmSaisieHeures
        wshAdmin.Range("TEC_Current_ID").value = .lstData.List(.lstData.ListIndex, 0)
        'Debug.Print "Sauvegarde de l'ID de l'enregistrement à modifier - " & _
            wshADMIN.Range("TEC_Current_ID").value
        
        .cmbProfessionnel.value = .lstData.List(.lstData.ListIndex, 1)
        .cmbProfessionnel.Enabled = False
        '.cmbProfessionnel.Locked = True
        
        .txtDate.value = Format(.lstData.List(.lstData.ListIndex, 2), "dd/mm/yyyy")
        .txtDate.Enabled = False
        '.txtDate.Locked = True
        
        .txtClient.value = .lstData.List(.lstData.ListIndex, 3)
        savedClient = .txtClient.value
        'Debug.Print "Double click on a entry - " & savedClient
        wshAdmin.Range("TEC_Client_ID").value = GetID_FromClientName(savedClient)
        'Debug.Print "Client_ID - " & wshADMIN.Range("TEC_Client_ID").value
         
        .txtActivite.value = .lstData.List(.lstData.ListIndex, 4)
        savedActivite = .txtActivite.value
        
        .txtHeures.value = Format(.lstData.List(.lstData.ListIndex, 5), "#0.00")
        savedHeures = .txtHeures.value
        
        .txtCommNote.value = .lstData.List(.lstData.ListIndex, 6)
        savedCommNote = .txtCommNote.value
        
        .chbFacturable.value = CBool(.lstData.List(.lstData.ListIndex, 7))
        savedFacturable = .chbFacturable.value
        
    End With
    
    With frmSaisieHeures
        frmSaisieHeures.cmdClear.Enabled = True
        frmSaisieHeures.cmdAdd.Enabled = False
        frmSaisieHeures.cmdUpdate.Enabled = False
        frmSaisieHeures.cmdDelete.Enabled = True
    End With
    
    rmv_state = rmv_modeAffichage
    
End Sub

