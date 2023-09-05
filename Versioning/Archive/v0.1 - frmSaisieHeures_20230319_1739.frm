VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaisieHeures 
   Caption         =   "Data Entry Form"
   ClientHeight    =   10092
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13860
   OleObjectBlob   =   "v0.1 - frmSaisieHeures_20230319_1739.frx":0000
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

