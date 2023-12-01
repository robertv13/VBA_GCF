VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIssueTracker 
   Caption         =   "Risk Management System"
   ClientHeight    =   9165.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19080
   OleObjectBlob   =   "frmIssueTracker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIssueTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCategory_Change()
    If cboProductType = vbNullString Then Exit Sub
    addRangeToCbo makeRangeContains(shtProduct, 1, 2, cboProductType.Value, 1, cboCategory.Value).Offset(0, 2), cboType
    setCtrToNullString txtClassification, txtLevel
End Sub

Private Sub cboProductType_Change()
    If cboCategory = vbNullString Then Exit Sub
    addRangeToCbo makeRangeContains(shtProduct, 1, 2, cboProductType.Value, 1, cboCategory.Value).Offset(0, 2), cboType
    setCtrToNullString txtClassification, txtLevel
End Sub

Private Sub cboType_Change()
    If cboType = vbNullString Then Exit Sub
    With makeRangeContains(shtProduct, 1, 2, cboProductType.Value, 2, cboType.Value)
        txtClassification = .Offset(0, 3)
        txtLevel = .Offset(0, 4)
    End With
End Sub

Private Sub cmdAddToTracker_Click()
    writeToSheet shtIssue, [rngIssueId], cboIdenDay & cboIdenMonth & cboIdenYear, cboIdentifiedBy, _
        txtStaffName, cboProductType, cboContractorName, cboCategory, cboType, txtClassification, _
            txtLevel, cboSeverity, txtDescription, txtActionTaken, cboStatus, cboResoDay & cboResoMonth & cboResoYear
    [rngIssueId] = [rngIssueId] + 1
End Sub

Private Sub UserForm_Initialize()
    'initialize the dates
    makeDateCboAll cboIdenDay, cboIdenMonth, cboIdenYear, True
    makeDateCboAll cboResoDay, cboResoMonth, cboResoYear, , , False

    addRangeToCbo [rngContractorName], cboContractorName
    addRangeToCbo [rngCategory], cboCategory
    addRangeToCbo [rngFilterBy], cboFilterBy
    makeAddRangeToCbo shtHelper, 7, cboSeverity
    makeAddRangeToCbo shtHelper, 9, cboProductType
    makeAddRangeToCbo shtHelper, 11, cboStatus, , 0
    makeAddRangeToCbo shtHelper, 13, cboIdentifiedBy
End Sub
