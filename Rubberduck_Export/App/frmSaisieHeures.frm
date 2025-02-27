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

    oEventHandler.List = rg.Value

End Property

Private Sub icoGraph_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    MsgBox "Affichage des statistiques à compléter"

End Sub

Private Sub lstNomClient_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim i As Long
    With Me.lstNomClient
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Me.txtClient.Value = .List(i, 0)
                wshAdmin.Range("TEC_Client_ID").Value = GetID_From_Client_Name(Me.txtClient.Value)
                Exit For
            End If
        Next i
    End With

End Sub

'******************************************* Execute when UserForm is displayed
Sub UserForm_Activate()

    Call Client_List_Import_All
    
    Dim lastUsedRow As Long
    lastUsedRow = wshClientDB.Range("A999999").End(xlUp).row
    frmSaisieHeures.ListData = wshClientDB.Range("A1:J" & lastUsedRow)
    
    With oEventHandler
        Set .SearchListBox = lstNomClient
        Set .SearchTextBox = txtClient
        .MaxRows = 10
        .ShowAllMatches = False
        .CompareMethod = vbTextCompare
    End With

    cmbProfessionnel.Value = ""
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
    
    If Me.name = "frmSaisieHeures" Then
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

    wshAdmin.Range("TEC_Initials").Value = Me.cmbProfessionnel.Value
    wshAdmin.Range("TEC_Prof_ID").Value = GetID_FromInitials(Me.cmbProfessionnel.Value)
    
    If wshAdmin.Range("TEC_Date").Value <> "" Then
        Call TEC_Advanced_Filter_And_Sort
        Call Refresh_ListBox_And_Add_Hours
    End If
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.Value) <> "" And _
        Trim(Me.txtDate.Value) <> "" And _
        Trim(Me.txtClient.Value) <> "" And _
        Trim(Me.txtHeures.Value) <> "" Then
        cmdAdd.Enabled = True
    End If
    
End Sub

Private Sub txtDate_Enter()

    If txtDate.Value = vbNullString Then
        txtDate.Value = Format(CDate(Now()), "dd/mm/yyyy")
    End If

End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If txtDate.Value = vbNullString Then
        txtDate.Value = Format(CDate(Now()), "dd/mm/yyyy")
    End If

    Dim strDate As String
    strDate = txtDate.Value

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
            txtDate.SelLength = Len(txtDate.Value)
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
    
    txtDate.Value = strDate

End Sub

Private Sub txtDateShowError(ErrorCaption As String)

    txtDate.BackColor = rgbPink
    lblDate.ForeColor = rgbRed
    lblDate.Caption = ErrorCaption
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Value)

End Sub

Private Sub txtDate_AfterUpdate()

    txtDate.BackColor = rgbWhite
    lblDate.Caption = "Date *"
    lblDate.ForeColor = Me.ForeColor
    
    wshAdmin.Range("TEC_Date").Value = CDate(Me.txtDate.Value)

    If wshAdmin.Range("TEC_Prof_ID").Value <> "" Then
        Call TEC_Advanced_Filter_And_Sort
        Call Refresh_ListBox_And_Add_Hours
    End If
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.Value) <> vbNullString And _
        Trim(Me.txtDate.Value) <> vbNullString And _
        Trim(Me.txtClient.Value) <> vbNullString And _
        Trim(Me.txtHeures.Value) <> vbNullString Then
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
        If Trim(Me.cmbProfessionnel.Value) <> "" And _
            Trim(Me.txtDate.Value) <> "" And _
            Trim(Me.txtClient.Value) <> "" And _
            Trim(Me.txtHeures.Value) <> "" Then
            cmdAdd.Enabled = True
        End If
    End If
    
    If rmv_state = rmv_modeAffichage Then
        If savedClient <> Me.txtClient.Value Or _
            savedActivite <> Me.txtActivite.Value Or _
            savedHeures <> Me.txtHeures.Value Or _
            savedCommNote <> Me.txtCommNote Or _
            savedFacturable <> Me.chbFacturable Then
            cmdUpdate.Enabled = True
        End If
    End If
    
    'Debug.Print rmv_state
    
    If rmv_state = rmv_modeAffichage Then
        If Me.txtClient.Value <> savedClient Then
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = True
        End If
    End If
    
End Sub

Private Sub txtActivite_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If txtActivite.Value <> savedActivite Then
            cmdUpdate.Enabled = True
        End If
    End If

End Sub

Sub txtHeures_AfterUpdate()

    'Validation des heures saisies
    Dim strHeures As String
    strHeures = Me.txtHeures.Value
    
    If InStr(".", strHeures) Then
        strHeures = Replace(strHeures, ".", ",")
    End If
    
    If IsNumeric(strHeures) = False Then
        MsgBox _
        Prompt:="La valeur saisie ne peut être utilisée comme valeur numérique!", _
        Title:="Validation d'une valeur numérique", _
        Buttons:=vbCritical
        Me.txtHeures.Value = ""
        Me.txtHeures.SetFocus
        Exit Sub
    End If
    
    Me.txtHeures.Value = Format(strHeures, "#0.00")
    
    'Enabled the ADD button if the minimum fields are non empty
    If rmv_state = rmv_modeCreation Then
        If Trim(Me.cmbProfessionnel.Value) <> "" And _
           Trim(Me.txtDate.Value) <> "" And _
           Trim(Me.txtClient.Value) <> "" And _
           Trim(Me.txtHeures.Value) <> "" Then
            cmdAdd.Enabled = True
        End If
    End If

    If rmv_state = rmv_modeAffichage Then
        If Me.txtHeures.Value <> savedHeures Then
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = True
        End If
    End If
    
End Sub

Private Sub chbFacturable_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If Me.chbFacturable.Value <> savedFacturable Then
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = True
        End If
    End If

End Sub

Private Sub txtCommNote_AfterUpdate()

    If rmv_state = rmv_modeAffichage Then
        If Me.txtCommNote.Value <> savedCommNote Then
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

    If wshAdmin.Range("TEC_Current_ID").Value = "" Then
        MsgBox Prompt:="Vous devez choisir un enregistrement à modifier !", _
               Title:="", _
               Buttons:=vbCritical
        Exit Sub
    End If

    ModifieLigneDetail

End Sub

Private Sub cmdDelete_Click()

    If wshAdmin.Range("TEC_Current_ID").Value = "" Then
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
        wshAdmin.Range("TEC_Current_ID").Value = .lstData.List(.lstData.ListIndex, 0)
        'Debug.Print "Sauvegarde de l'ID de l'enregistrement à modifier - " & _
            wshADMIN.Range("TEC_Current_ID").Value
        
        .cmbProfessionnel.Value = .lstData.List(.lstData.ListIndex, 1)
        .cmbProfessionnel.Enabled = False
        '.cmbProfessionnel.Locked = True
        
        .txtDate.Value = Format(.lstData.List(.lstData.ListIndex, 2), "dd/mm/yyyy")
        .txtDate.Enabled = False
        '.txtDate.Locked = True
        
        .txtClient.Value = .lstData.List(.lstData.ListIndex, 3)
        savedClient = .txtClient.Value
        'Debug.Print "Double click on a entry - " & savedClient
        wshAdmin.Range("TEC_Client_ID").Value = GetID_From_Client_Name(savedClient)
        'Debug.Print "Client_ID - " & wshADMIN.Range("TEC_Client_ID").Value
         
        .txtActivite.Value = .lstData.List(.lstData.ListIndex, 4)
        savedActivite = .txtActivite.Value
        
        .txtHeures.Value = Format(.lstData.List(.lstData.ListIndex, 5), "#0.00")
        savedHeures = .txtHeures.Value
        
        .txtCommNote.Value = .lstData.List(.lstData.ListIndex, 6)
        savedCommNote = .txtCommNote.Value
        
        .chbFacturable.Value = CBool(.lstData.List(.lstData.ListIndex, 7))
        savedFacturable = .chbFacturable.Value
        
    End With
    
    With frmSaisieHeures
        frmSaisieHeures.cmdClear.Enabled = True
        frmSaisieHeures.cmdAdd.Enabled = False
        frmSaisieHeures.cmdUpdate.Enabled = False
        frmSaisieHeures.cmdDelete.Enabled = True
    End With
    
    rmv_state = rmv_modeAffichage
    
End Sub

