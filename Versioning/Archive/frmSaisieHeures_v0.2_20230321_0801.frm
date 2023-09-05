VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaisieHeures 
   Caption         =   "Gestion des heures travaillées"
   ClientHeight    =   10092
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13356
   OleObjectBlob   =   "frmSaisieHeures_v0.2_20230321_0801.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSaisieHeures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'******************************************* Execute when UserForm is displayed
Sub UserForm_Activate()

    'Import Clients List to have the latest version
    Call ImportClientsList
    
    'Working Worksheet 'HeuresFiltered'
'    Dim shFiltered As Worksheet
'    Set shFiltered = ThisWorkbook.Sheets("HeuresFiltered")
'    shFiltered.UsedRange.Clear
    'shFiltered.Activate
    
    'Call FilterProfDate
    Call RefreshListBox
    
    cmbProfessionnel.SetFocus
      
End Sub

Public Sub cmbProfessionnel_AfterUpdate()

    If Me.cmbProfessionnel.value = "" Then
        Me.cmbProfessionnel.SetFocus
        Exit Sub
    End If
    
    Call FilterProfDate
    Call RefreshListBox
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel) <> "" And _
        Trim(Me.txtDate) <> "" And _
        Trim(Me.cmbClient) <> "" And _
        Trim(Me.txtHeures) <> "" Then
            cmdAdd.Enabled = True
    End If
    
    txtDate.SetFocus
    
End Sub
Sub txtDate_AfterUpdate()

    If Me.txtDate.value = "" Then
        Me.txtDate.SetFocus
        Exit Sub
    End If
    
    Dim strDate As String
    strDate = Me.txtDate.value
    Dim tmpAnnee, tmpMois, tmpJour As Integer
    tmpAnnee = Format(Year(Now()), "0000")
    tmpMois = Format(Month(Now()), "00")
    tmpJour = Format(Day(Now()), "00")
    
    If Len(strDate) = 0 Then
        strDate = tmpAnnee & "-" & tmpMois & "-" & tmpJour
    ElseIf Len(strDate) = 2 Then
        strDate = strDate & "-" & tmpMois & "-" & tmpAnnee
    ElseIf Len(strDate) = 5 Then
        strDate = strDate & "-" & tmpAnnee
    End If
    
    'Validation de la date
    If IsDate(strDate) = False Then
        MsgBox "La date est invalide!"
        Me.txtDate.SetFocus
        Exit Sub
    End If
    
    Me.txtDate.value = strDate

    Call FilterProfDate
    Call RefreshListBox
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel) <> "" And _
        Trim(Me.txtDate) <> "" And _
        Trim(Me.cmbClient) <> "" And _
        Trim(Me.txtHeures) <> "" Then
            cmdAdd.Enabled = True
    End If
    
End Sub

Private Sub cmbClient_AfterUpdate()
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel) <> "" And _
        Trim(Me.txtDate) <> "" And _
        Trim(Me.cmbClient) <> "" And _
        Trim(Me.txtHeures) <> "" Then
            cmdAdd.Enabled = True
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
        MsgBox "Veuillez saisir une valeur numérique"
        Me.txtHeures.value = ""
        Me.txtHeures.SetFocus
        Exit Sub
    End If
    
    Me.txtHeures.value = Format(strHeures, "#0.00")
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel) <> "" And _
        Trim(Me.txtDate) <> "" And _
        Trim(Me.cmbClient) <> "" And _
        Trim(Me.txtHeures) <> "" Then
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

    With frmSaisieHeures
        .txtID.value = .lstData.List(.lstData.ListIndex, 0)
        .cmbProfessionnel.value = .lstData.List(.lstData.ListIndex, 1)
        .txtDate.value = Format(.lstData.List(.lstData.ListIndex, 2), "dd-mm-yyyy")
        .cmbClient.value = .lstData.List(.lstData.ListIndex, 3)
        .txtActivite.value = .lstData.List(.lstData.ListIndex, 4)
        .txtHeures.value = Format(.lstData.List(.lstData.ListIndex, 5), "#0.00")
        .txtCommNote.value = .lstData.List(.lstData.ListIndex, 6)
        .chbFacturable.value = .lstData.List(.lstData.ListIndex, 7)
    End With

    With frmSaisieHeures
        frmSaisieHeures.cmdClear.Enabled = True
        frmSaisieHeures.cmdAdd.Enabled = False
        frmSaisieHeures.cmdUpdate.Enabled = True
        frmSaisieHeures.cmdDelete.Enabled = True
    End With
    
End Sub

