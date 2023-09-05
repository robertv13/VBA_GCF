VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportHeures 
   Caption         =   "Exportation des heures saisies"
   ClientHeight    =   3090
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5790
   OleObjectBlob   =   "frmExportHeures_v1.7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportHeures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Activate()
    
    MsgBox "frmExportHeures - UserForm_Activate"
    
    MsgBox Format(ThisWorkbook.Worksheets("Menu").Range("F6"), "dd/mm/yyyy hh:MM:ss")
    Call FilterTimeAndDate

End Sub

Private Sub UserForm_Initialize()

    MsgBox "frmExportHeures - UserForm_Initialize"
    
    frmExportHeures.txtDateLimiteExport.value = _
        Format(ThisWorkbook.Worksheets("Menu").Range("F6"), "dd/mm/yyyy hh:MM:ss")

End Sub

Private Sub UserForm_Terminate()

    MsgBox "frmExportHeures - UserForm_Terminate"
    
    Me.Hide
    Unload Me
    
    wsMenu.Select

End Sub

Private Sub cmdAnnulerExport_Click()

    MsgBox "Annulation de l'exportation"

End Sub

Private Sub cmdExport_Click()

    MsgBox "Exportation des heures vers le fichier principal"

End Sub

'Autofilter after a specific Date and Time - 2023-03-30
Sub FilterTimeAndDate()

    Dim ws As Worksheet
    Dim rng As Range
    Dim strDateTime As String
    
    'Set the MINIMUM date and time
    strDateTime = Format(CDate(ThisWorkbook.Worksheets("Menu").Range("F6")), _
        "mm/dd/yyyy hh:MM:ss")

    'Reference the From Worksheet and Range
    Set ws = ThisWorkbook.Worksheets("HeuresBase")
    Set rng = ws.Range("A1").CurrentRegion
    
    'Turn OFF Autofilter Mode
    ws.AutoFilterMode = False
    
    'Apply AutoFilter magic formula
    With rng
        rng.AutoFilter Field:=9, Criteria1:=">" & strDateTime
        rng.AutoFilter Field:=12, Criteria1:="FAUX"
    End With
    
    frmExportHeures.txtNombreEnregistrements.value = _
        rng.Columns(1).SpecialCells(xlCellTypeVisible).Cells.count - 1
    
    'Once filtered, worksheet should only show the filtered records
    Dim shHTE As Worksheet
    Set shHTE = ThisWorkbook.Sheets("HeuresAExporter")
    shHTE.UsedRange.Clear
    
    Set rng = ws.Range("A1").CurrentRegion
    rng.Select
    rng.Copy shHTE.Range("A1")
    
    shHTE.Activate
    
    ws.Activate
    ws.AutoFilterMode = False
    ws.ShowAllData
    
    frmExportHeures.cmdExport.Enabled = True
    
End Sub

