VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmzzzExportHeures 
   Caption         =   "Exportation des heures saisies"
   ClientHeight    =   3690
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5790
   OleObjectBlob   =   "frmzzzExportHeures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmzzzExportHeures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    
    'MsgBox "frmExportHeures - UserForm_Activate"
    
    Call FilterTimeAndDate

End Sub

Private Sub UserForm_Initialize()

    'MsgBox "frmExportHeures - UserForm_Initialize"
    
    frmExportHeures.txtDateLimiteExport.value = _
                                              Format(ThisWorkbook.Worksheets("Menu").Range("F6"), "dd/mm/yyyy hh:MM:ss")

End Sub

Private Sub UserForm_Terminate()

    'MsgBox "frmExportHeures - UserForm_Terminate"
    
    Me.Hide
    Unload Me
    
    'wsMenu.Select

End Sub

Private Sub cmdAnnulerExport_Click()

    MsgBox "L'exportation a été annulée à votre demande", vbInformation
    
    Me.Hide
    Unload Me
    
    wshMenu.Select

End Sub

Private Sub cmdExport_Click()

    frmExportHeures.cmdExport.Enabled = False
    frmExportHeures.cmdAnnulerExport.Enabled = False
    
    'MsgBox "Exportation des heures vers le fichier principal"
    
    Dim wsHTE As Worksheet
    Set wsHTE = wshHoursToExport
    wsHTE.Activate
    
    'Setup the destination workbook/worksheet
    Dim wbGCF As Workbook
    Set wbGCF = Workbooks.Open("C:\VBA\GC FISCALITÉ\Excel Ctb GC Fiscalité.xlsm")
    wbGCF.Worksheets("TEC").Activate
    
    'Setup the row to use
    Dim currentRow As Long
    currentRow = ActiveSheet.Cells(Rows.count, 1).End(xlUp).row + 1
    
    Dim rng As Range
    Set rng = wsHTE.Range("A1").CurrentRegion
    
    Dim r As Long
    For r = 2 To rng.Rows.count
        'Debug.Print r & " - " & Cells(r, 1).value & " - " & Cells(r, 2).value & _
        " - " & Format(Cells(r, 3).value, "dd/mm/yyyy") & _
        " - " & Cells(r, 4).value
        ActiveSheet.Cells(currentRow, 1) = rng.Cells(r, 3).value
        ActiveSheet.Cells(currentRow, 2) = rng.Cells(r, 4).value
        ActiveSheet.Cells(currentRow, 3) = rng.Cells(r, 5).value
        ActiveSheet.Cells(currentRow, 4) = rng.Cells(r, 6).value
        ActiveSheet.Cells(currentRow, 5) = rng.Cells(r, 7).value
        ActiveSheet.Cells(currentRow, 6) = rng.Cells(r, 1).value
        
        currentRow = currentRow + 1
        frmExportHeures.txtProgression.value = r - 1
    Next r
    
    'Save changes & Close the workbook
    wbGCF.Close SaveChanges:=True
    
    'Update Date of last export (frmExportHeures & wsMenu)
    frmExportHeures.txtNextExportDate.value = Format(CDate(Now), "dd/mm/yyyy hh:MM:ss")
    wshMenu.Range("F6").value = frmExportHeures.txtNextExportDate.value

    MsgBox "Félicitations - L'exportation des données s'est bien déroulée", vbInformation
    
End Sub

'Autofilter after a specific Date and Time - 2023-03-30
Sub FilterTimeAndDate()

    Dim ws As Worksheet
    Dim rng As Range
    Dim strDateTime As String
    
    'Set the MINIMUM date and time
    strDateTime = Format(CDate(wshMenu.Range("F6")), _
                         "mm/dd/yyyy hh:MM:ss")

    'Reference the From Worksheet and Range
    Set ws = wshBaseHours
    Set rng = ws.Range("A1").CurrentRegion
    
    'Turn OFF Autofilter Mode
    ws.AutoFilterMode = False
    
    'Apply AutoFilter magic formula
    With rng
        rng.AutoFilter Field:=9, Criteria1:=">" & strDateTime
        rng.AutoFilter Field:=12, Criteria1:="FAUX"
    End With
    
    Dim rowsToExport As Long
    rowsToExport = rng.Columns(1).SpecialCells(xlCellTypeVisible).Cells.count - 1
    
    frmExportHeures.txtNombreEnregistrements.value = rowsToExport
        
    If rowsToExport = 0 Then
        MsgBox "Il n'y a aucune donnée à exporter !", vbInformation
        GoTo Done
    Else
        'Once filtered, worksheet should only show the filtered records
        Dim shHTE As Worksheet
        Set shHTE = wshHoursToExport
        shHTE.usedRange.Clear
        
        'Copy to destination worksheet (wshHoursToExport)
        rng.Copy shHTE.Range("A1")
        
        shHTE.Activate
        
        frmExportHeures.cmdExport.Enabled = True
    End If
    
Done:
    ws.Activate
    ws.AutoFilterMode = False
    ws.ShowAllData
    
End Sub


