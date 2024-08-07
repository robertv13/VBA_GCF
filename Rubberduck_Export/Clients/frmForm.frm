VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmForm 
   Caption         =   "Gestion du fichier Clients"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13095
   OleObjectBlob   =   "frmForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EnableEvents As Boolean

Private Sub cmbSearchColumn_Change()

    If Me.EnableEvents = False Then Exit Sub
    
    If Me.cmbSearchColumn.Value = "All" Then
        Call Reset
    Else
        Me.txtSearch.Value = ""
        Me.txtSearch.Enabled = True
        Me.cmdSearch.Enabled = True
    End If

End Sub

Private Sub cmdDelete_Click()
    
    If Selected_List = 0 Then
        MsgBox "No row is selected.", vbOKOnly + vbInformation, "Delete"
        Exit Sub
    End If
    
    Dim i As VbMsgBoxResult
    i = MsgBox("Do you want to delete the selected record?", vbYesNo + vbQuestion, "Confirmation")
    If i = vbNo Then Exit Sub
    
    Dim iRow As Long
    iRow = Application.WorksheetFunction.Match(Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0), _
    ThisWorkbook.Sheets("Database").Range("A:A"), 0)
    
    ThisWorkbook.Sheets("Database").Rows(iRow).Delete
    
    Call Reset
    
    MsgBox "Selected record has been deleted.", vbOKOnly + vbInformation, "Deleted"
    
End Sub

Private Sub cmdEdit_Click()
    
    If Selected_List = 0 Then
        MsgBox "No row is selected.", vbOKOnly + vbInformation, "Edit"
        Exit Sub
    End If
    
    'Code to update the value to respective controls
    Dim sGender As String
    Me.txtRowNumber.Value = Application.WorksheetFunction.Match(Me.lstDatabase.List(Me.lstDatabase.ListIndex, 0), _
    ThisWorkbook.Sheets("Database").Range("A:A"), 0)
    Me.txtID.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 1)
    Me.txtName.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 2)
    sGender = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 3)
    If sGender = "Female" Then
        Me.optFemale.Value = True
    Else
        Me.optMale.Value = True
    End If
    Me.cmbDepartment.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 4)
    Me.txtCity.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 5)
    Me.txtCountry.Value = Me.lstDatabase.List(Me.lstDatabase.ListIndex, 6)
    
    MsgBox "Please make the required changes and click on 'Save' button to update.", vbOKOnly + vbInformation, "Edit"
    
End Sub

Private Sub cmdPrint_Click()
    
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to print the employee details?", vbYesNo + vbInformation, "Print")
    If msgValue = vbNo Then Exit Sub
    
    If ValidatePrintDetails() = True Then
        Call Print_Form
    End If
    
End Sub

Private Sub cmdReset_Click()

    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to reset the form?", vbYesNo + vbInformation, "Confirmation")
    If msgValue = vbNo Then Exit Sub
    
    Call Reset

End Sub

Private Sub cmdSave_Click()
    
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to save the data?", vbYesNo + vbInformation, "Confirmation")
    If msgValue = vbNo Then Exit Sub
    
    If ValidateEntries() = True Then
    
        Call Submit
        Call Reset
    
    End If
    
End Sub

Private Sub cmdSearch_Click()

    If Me.txtSearch.Value = "" Then
        MsgBox "PLease enter the search value.", vbOKOnly + vbInformation, "Search"
        Exit Sub
    End If
    
    Call SearchData
    
End Sub

Private Sub UserForm_Initialize()

    Call Reset

End Sub
