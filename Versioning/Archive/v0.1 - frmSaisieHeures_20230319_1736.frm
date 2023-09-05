VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaisieHeures 
   Caption         =   "Data Entry Form"
   ClientHeight    =   10092
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13860
   OleObjectBlob   =   "v0.1 - frmSaisieHeures_20230319_1736.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSaisieHeures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClient_AfterUpdate()
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel) <> "" And _
        Trim(Me.txtDate) <> "" And _
        Trim(Me.cmbClient) <> "" And _
        Trim(Me.txtHeures) <> "" Then
            cmdAdd.Enabled = True
    End If

End Sub

Sub cmbProfessionnel_AfterUpdate()

    'MsgBox "Temp - Sub cmbProfessionnel_AfterUpdate()"
    If Me.cmbProfessionnel.value = "" Then
        Me.cmbProfessionnel.SetFocus
        Exit Sub
    End If
    
    If Me.txtDate.value = "" Then
        Me.txtDate.SetFocus
        Exit Sub
    End If
   
    Call FilterProfDate
    Call RefreshData
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel) <> "" And _
        Trim(Me.txtDate) <> "" And _
        Trim(Me.cmbClient) <> "" And _
        Trim(Me.txtHeures) <> "" Then
            cmdAdd.Enabled = True
    End If
    
    Me.txtDate.SetFocus
    
End Sub

'*************************************************************** cmdAdd_Click()
Sub cmdAdd_Click()

    'MsgBox "Temp - Sub cmdAdd_Click()"
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Heures")
    
    Dim lastRow As Long
    lastRow = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    
    'Validations first (one field at a time)
    If Me.cmbProfessionnel.value = "" Then
        MsgBox "Le professionnel est OBLIGATOIRE !", vbCritical
        Me.cmbProfessionnel.SetFocus
        Exit Sub
    End If

    If Me.txtDate.value = "" Then
        MsgBox "La date est OBLIGATOIRE !", vbCritical
        Me.txtDate.SetFocus
        Exit Sub
    End If

    If Me.cmbClient.value = "" Then
        MsgBox "Le nom du client est OBLIGATOIRE !", vbCritical
        Me.cmbClient.SetFocus
    Exit Sub
    End If
    
    If Me.txtHeures.value = "" Then
        MsgBox "Le nombre d'heures est OBLIGATOIRE !", vbCritical
        Me.txtHeures.SetFocus
        Exit Sub
    End If

    'Load the cmb & txt into the 'Heures' worksheet
    sh.Range("A" & lastRow + 1).value = "=row()-1"
    sh.Range("B" & lastRow + 1).value = Me.cmbProfessionnel.value
    sh.Range("C" & lastRow + 1).value = Format(Me.txtDate.value, "dd-mm-yyyy")
    sh.Range("D" & lastRow + 1).value = Me.cmbClient.value
    sh.Range("E" & lastRow + 1).value = Me.txtActivite.value
    sh.Range("F" & lastRow + 1).value = Format(Me.txtHeures.value, "#0.00")
    sh.Range("G" & lastRow + 1).value = Me.txtCommNote.value
    sh.Range("H" & lastRow + 1).value = Me.chbFacturable.value
    sh.Range("I" & lastRow + 1).value = Now
    sh.Range("J" & lastRow + 1).value = False
    sh.Range("K" & lastRow + 1).value = ""

    'Empty the fields after saving
    cmbClient.value = ""
    txtActivite.value = ""
    txtHeures.value = ""
    txtCommNote.value = ""
        
    Call FilterProfDate
    Call RefreshData
    Me.cmbClient.SetFocus
    
End Sub

'************************************************************* cmdClear_Click()
Sub cmdClear_Click()

    'MsgBox "Temp - Sub cmdClear_Click()"
    Me.cmbClient.value = ""
    Me.txtActivite.value = ""
    Me.txtHeures.value = ""
    Me.txtCommNote.value = ""

    Call FilterProfDate
    Call RefreshData
    
    cmdAdd.Enabled = False
    cmdClear.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    
    Me.cmbClient.SetFocus
    
End Sub

'************************************************************ cmdDelete_Click()
Sub cmdDelete_Click()

    'MsgBox "Temp - Sub cmdDelete_Click()"
    If txtID.value = "" Then
        MsgBox "Vous devez choisir un enregistrement à DÉTRUIRE !"
        Exit Sub
    End If
    
    Dim answerYesNo As Integer
    answerYesNo = MsgBox("Êtes-vous certain de vouloir DÉTRUIRE cet enregistrement ? ", _
                                vbYesNo + vbQuestion, "Confirmation de DESTRUCTION")
    If answerYesNo = vbNo Then
        MsgBox "Cet enregistrement sera DÉTRUIT ! "
        Call cmdClear_Click
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Heures")
    
    Dim selectedRow As Long
    selectedRow = Application.WorksheetFunction.Match(CLng(Me.txtID.value), _
                              sh.Range("A:A"), 0)
    
    sh.Range("A" & selectedRow).EntireRow.Delete
    
    'Empty the fields after deleting
    Me.cmbClient.value = ""
    Me.txtActivite.value = ""
    Me.txtHeures.value = ""
    Me.txtCommNote.value = ""

    MsgBox "L'enregistrement a été DÉTRUIT !"

    Call FilterProfDate
    Call RefreshData
    
End Sub

'************************************************************ cmdUpdate_Click()
Sub cmdUpdate_Click()

    'MsgBox "Temp - Sub cmdUpdate_Click()"
    If txtID.value = "" Then
        MsgBox "Vous devez choisir un enregistrement à modifier !"
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Heures")
    
    Dim selectedRow As Long
    selectedRow = Application.WorksheetFunction.Match(CLng(Me.txtID.value), _
                    sh.Range("A:A"), 0)
    
    'Validations first (one field at a time)
    If Me.cmbProfessionnel.value = "" Then
        MsgBox "Le professionnel est OBLIGATOIRE !", vbCritical
        Me.cmbProfessionnel.SetFocus
        Exit Sub
    End If

    If Me.txtDate.value = "" Then
        MsgBox "La date est OBLIGATOIRE !", vbCritical
        Me.txtDate.SetFocus
        Exit Sub
    End If

    If Me.cmbClient.value = "" Then
        MsgBox "Le nom du client est OBLIGATOIRE !", vbCritical
        Me.cmbClient.SetFocus
    Exit Sub
    End If
    
    If Me.txtHeures.value = "" Then
        MsgBox "Le nombre d'heures est OBLIGATOIRE !", vbCritical
        Me.txtHeures.SetFocus
        Exit Sub
    End If

    sh.Range("B" & selectedRow).value = Me.cmbProfessionnel.value
    sh.Range("C" & selectedRow).value = Format(Me.txtDate.value, "dd-mm-yyyy")
    sh.Range("D" & selectedRow).value = Me.cmbClient.value
    sh.Range("E" & selectedRow).value = Me.txtActivite.value
    sh.Range("F" & selectedRow).value = Format(Me.txtHeures.value, "#0.00")
    sh.Range("G" & selectedRow).value = Me.txtCommNote.value
    sh.Range("H" & selectedRow).value = Me.chbFacturable.value
    sh.Range("I" & selectedRow).value = Now
    sh.Range("J" & selectedRow).value = False
    sh.Range("K" & selectedRow).value = ""
   
    'Empty the fields after modifying
    cmbClient.value = ""
    txtActivite.value = ""
    txtHeures.value = ""
    txtCommNote.value = ""
    
    Call FilterProfDate
    Call RefreshData
    Me.cmbClient.SetFocus

End Sub

'*********************************** Select a row and display it in the details
Sub lstData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    'MsgBox "Temp - Sub lstData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)"
    Me.txtID.value = Me.lstData.List(Me.lstData.ListIndex, 0)
    Me.cmbProfessionnel.value = Me.lstData.List(Me.lstData.ListIndex, 1)
    Me.txtDate.value = Format(Me.lstData.List(Me.lstData.ListIndex, 2), "dd-mm-yyyy")
    Me.cmbClient.value = Me.lstData.List(Me.lstData.ListIndex, 3)
    Me.txtActivite.value = Me.lstData.List(Me.lstData.ListIndex, 4)
    Me.txtHeures.value = Format(Me.lstData.List(Me.lstData.ListIndex, 5), "#0.00")
    Me.txtCommNote.value = Me.lstData.List(Me.lstData.ListIndex, 6)
    Me.chbFacturable.value = Me.lstData.List(Me.lstData.ListIndex, 7)

    cmdClear.Enabled = True
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
End Sub

'********************* Reload listBox from HeuresFiltered and reset the buttons
Sub RefreshData()

    'MsgBox "Temp - Sub RefreshData()"
    Dim shFiltered As Worksheet
    Set shFiltered = ThisWorkbook.Sheets("HeuresFiltered")
    
    'Last Row used in column A
    Dim lastRow As Long
    lastRow = Application.WorksheetFunction.CountA(shFiltered.Range("A:A"))
    'MsgBox "Temp - La valeur de lastRow est de " & lastRow
        
    With lstData
        .ColumnHeads = True
        .ColumnCount = 9
        .ColumnWidths = "25; 30; 55; 120; 190; 30; 90; 35; 70"
        
        If lastRow <= 1 Then
            .RowSource = "HeuresFiltered!A2:K2"
            lastRow = 2
        Else
            .RowSource = "HeuresFiltered!A2:K" & lastRow
        End If
    End With

    'Add hours to totalHeures
    Dim nbrRows, i As Integer
    nbrRows = Me.lstData.ListCount
    Dim totalHeures As Double
    
    If nbrRows >= 1 Then
        For i = 0 To nbrRows - 1
            totalHeures = totalHeures + CCur(lstData.List(i, 5))
        Next
        frmSaisieHeures.txtTotalHeures.value = Format(totalHeures, "#0.00")
    End If

    'Disable all buttons
    cmdClear.Enabled = False
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    
    cmbClient.SetFocus
    
End Sub

Sub txtDate_AfterUpdate()

    'MsgBox "Temp - Sub txtDate_AfterUpdate()"
    If Me.cmbProfessionnel.value = "" Then
        Me.cmbProfessionnel.SetFocus
        Exit Sub
    End If
    
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
    
    Me.txtDate.value = strDate

    Call FilterProfDate
    Call RefreshData
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel) <> "" And _
        Trim(Me.txtDate) <> "" And _
        Trim(Me.cmbClient) <> "" And _
        Trim(Me.txtHeures) <> "" Then
            cmdAdd.Enabled = True
    End If
    
    Me.cmbClient.SetFocus
    
End Sub

Sub txtHeures_AfterUpdate()

    'Validation des heures saisies
    Dim strHeures As String
    strHeures = Me.txtHeures.value
    
    If InStr(".", strHeures) Then
        strHeures = Replace(strHeures, ".", ",")
        Me.txtHeures.value = Format(strHeures, "#0.00")
    Else
        Me.txtHeures.value = Format(Me.txtHeures.value, "#0.00")
    End If
    
    'Enabled the ADD button if the minimum fields are non empty
    If Trim(Me.cmbProfessionnel) <> "" And _
        Trim(Me.txtDate) <> "" And _
        Trim(Me.cmbClient) <> "" And _
        Trim(Me.txtHeures) <> "" Then
            cmdAdd.Enabled = True
    End If

End Sub
'******************************************* Execute when UserForm is displayed
Sub UserForm_Activate()

    'Import Clients List
    Call ImportClientsList
    
    'Working worksheet 'HeuresFiltered'
    Dim shFiltered As Worksheet
    Set shFiltered = ThisWorkbook.Sheets("HeuresFiltered")
    shFiltered.UsedRange.Clear
    shFiltered.Activate
    
    Call FilterProfDate
    Call RefreshData
    cmdAdd.Accelerator = "A"
    cmdClear.Accelerator = "E"
    cmdDelete.Accelerator = "D"
    cmdUpdate.Accelerator = "M"
    cmbProfessionnel.SetFocus
      
End Sub

