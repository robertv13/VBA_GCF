VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   9264.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13044
   OleObjectBlob   =   "frmSaisieHeures_v1.4_20230326_1258.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSaisieHeures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oEventHandler As New clsSearchableDropdown '2023-03-21 @ 09:16

'Description: Allows the calling code to set the data
Public Property Let ListData(ByVal rg As Range)
    oEventHandler.List = rg.value
End Property


Private Sub chbFacturable_Click()

End Sub

'******************************************* Execute when UserForm is displayed
Sub Userform_Activate()

    'Import Clients List to have the latest version
    Call ImportClientsList
    frmSaisieHeures.ListData = shImportedClients.Range("A1").CurrentRegion
    
    rmv_state = rmv_modeInitial

    Call RefreshListBox
    
    With oEventHandler
        Set .SearchListBox = lstNomClient
        Set .SearchTextBox = txtClient
        
        .MaxRows = 10
        .ShowAllMatches = False
        .CompareMethod = vbTextCompare
    End With
    
'    Dim sh As Worksheet
'    Set sh = ThisWorkbook.Sheets("Heures")
'    sh.AutoFilterMode = False
'    sh.ListObjects("tCharges").TotalsRowRange.Delete
    
    cmbProfessionnel.SetFocus
      
End Sub

Private Sub Userform_Terminate()
    
    'Save
    ThisWorkbook.Save
    
    'Clean up
    Set oEventHandler = Nothing

End Sub

Public Sub cmbProfessionnel_AfterUpdate()

'    If Me.cmbProfessionnel.value = "" Then
'        Me.cmbProfessionnel.SetFocus
'        Exit Sub
'    End If
    
    Call FilterProfDate
    Call RefreshListBox
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> "" And _
        Trim(Me.txtDate.value) <> "" And _
        Trim(Me.txtClient.value) <> "" And _
        Trim(Me.txtHeures.value) <> "" Then
            cmdAdd.Enabled = True
    End If
    
    txtDate.SetFocus
    
End Sub

Sub txtDate_AfterUpdate()

    'If Me.txtDate.value = "" Then
    '    Me.txtDate.SetFocus
    '    Exit Sub
    'End If
    
    Dim strDate As String
    Dim separateur As String
    separateur = "/"
    
    strDate = Me.txtDate.value
    Dim currentYear, currentMonth, currentDay As Integer
    currentYear = Format(Year(Now()), "0000")
    currentMonth = Format(Month(Now()), "00")
    currentDay = Format(Day(Now()), "00")
    
    If Len(strDate) = 0 Then
        strDate = currentDay & separateur & currentMonth & separateur & _
            currentYear
    ElseIf Len(strDate) <= 2 Then
        strDate = Format(strDate, "00") & separateur & currentMonth & _
            separateur & currentYear
    ElseIf Len(strDate) = 5 Then
        strDate = strDate & separateur & currentYear
    End If
    
    'Validation de la date
    If IsDate(strDate) = False Then
        MsgBox _
            Prompt:="La valeur saisie ne peut être utilisée comme une date valide!", _
            Title:="Validation de la date", _
            Buttons:=vbCritical
        Me.txtDate.value = ""
        Me.txtDate.SetFocus
        Exit Sub
    End If
    
    Me.txtDate.value = strDate

    Call FilterProfDate
    Call RefreshListBox
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> "" And _
        Trim(Me.txtDate.value) <> "" And _
        Trim(Me.txtClient.value) <> "" And _
        Trim(Me.txtHeures.value) <> "" Then
            cmdAdd.Enabled = True
    End If
    
End Sub

Private Sub txtClient_Enter()

    If rmv_state = rmv_modeInitial Then
        rmv_state = rmv_modeCreation
    End If

End Sub

Private Sub txtClient_AfterUpdate()
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel.value) <> "" And _
        Trim(Me.txtDate.value) <> "" And _
        Trim(Me.txtClient.value) <> "" And _
        Trim(Me.txtHeures.value) <> "" Then
            cmdAdd.Enabled = True
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
    If Trim(Me.cmbProfessionnel.value) <> "" And _
        Trim(Me.txtDate.value) <> "" And _
        Trim(Me.txtClient.value) <> "" And _
        Trim(Me.txtHeures.value) <> "" Then
            cmdAdd.Enabled = True
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

    ModifieLigneDetail

End Sub

Private Sub cmdDelete_Click()

    EffaceLigneDetail

End Sub

'****************************************** Get a row and display it in the form
Sub lstData_dblClick(ByVal Cancel As MSForms.ReturnBoolean)

    rmv_state = rmv_modeAffichage
    
    With frmSaisieHeures
        .txtID.value = .lstData.List(.lstData.ListIndex, 0)
        
        .cmbProfessionnel.value = .lstData.List(.lstData.ListIndex, 1)
        .cmbProfessionnel.Enabled = False
        '.cmbProfessionnel.Locked = True
        
        .txtDate.value = Format(.lstData.List(.lstData.ListIndex, 2), "dd/mm/yyyy")
        .txtDate.Enabled = False
        '.txtDate.Locked = True
        
        .txtClient.value = .lstData.List(.lstData.ListIndex, 3)
         savedClient = .txtClient.value
         
        .txtActivite.value = .lstData.List(.lstData.ListIndex, 4)
        savedActivite = .txtActivite.value
        
        .txtHeures.value = Format(.lstData.List(.lstData.ListIndex, 5), "#0.00")
        savedHeures = .txtHeures.value
        
        .txtCommNote.value = .lstData.List(.lstData.ListIndex, 6)
        savedCommNote = .txtCommNote.value
        
        .chbFacturable.value = .lstData.List(.lstData.ListIndex, 7)
        savedFacturable = .chbFacturable.value
        
    End With
    
    With frmSaisieHeures
        frmSaisieHeures.cmdClear.Enabled = True
        frmSaisieHeures.cmdAdd.Enabled = False
        frmSaisieHeures.cmdUpdate.Enabled = False
        frmSaisieHeures.cmdDelete.Enabled = True
    End With
    
    MsgBox _
        savedClient & vbNewLine & _
        savedActivite & vbNewLine & _
        savedHeures & vbNewLine & _
        savedCommNote & vbNewLine & _
        savedFacturable

End Sub

