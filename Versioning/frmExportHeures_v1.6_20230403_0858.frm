VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportHeures 
   Caption         =   "Exportation des heures saisies"
   ClientHeight    =   3090
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5790
   OleObjectBlob   =   "frmExportHeures_v1.6_20230403_0858.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportHeures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub frmExportHeures_Activate()



End Sub
Sub frmExportHeures_Activate()

    MsgBox ThisWorkbook.Worksheets("Menu").Range("F6")
    'Call FilterTimeAndDate

End Sub

Private Sub cmdAnnulerExport_Click()

    MsgBox "Annulation de l'exportation"

End Sub

Private Sub cmdExport_Click()

    MsgBox "Exportation des heures dans le fichier principal"

End Sub

'Autofilter after a specific Date and Time - 2023-03-30
Sub FilterTimeAndDate()

    Dim ws As Worksheet
    Dim rng As Range
    Dim strDateTime As String
    
    'Reference the Worksheet and Range
    Set ws = ThisWorkbook.Worksheets("Heures")
    Set rng = ws.Range("A1:A133")
    
    'Turn OFF Autofilter Mode
    ws.AutoFilterMode = False
    
    'Set the MINIMUM date and time
    strDateTime = Format(cell("F6").value, "mm/dd/yyyy hh:MM:ss")
    
    'Apply AutoFilter magic formula
    Range("A1").AutoFilter Field:=1, Criteria1:=">" & strDateTime
    
    MsgBox ws.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.count - 1
    
End Sub
