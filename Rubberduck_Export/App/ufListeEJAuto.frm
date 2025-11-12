VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeEJAuto 
   Caption         =   "Choisir l'entrée récurrente à utiliser"
   ClientHeight    =   5160
   ClientLeft      =   7065
   ClientTop       =   6180
   ClientWidth     =   7155
   OleObjectBlob   =   "ufListeEJAuto.frx":0000
End
Attribute VB_Name = "ufListeEJAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wrappers As Collection

Private Sub UserForm_Initialize()
    
    Dim lastUsedRow As Long
    lastUsedRow = wsdGL_EJ_Recurrente.Cells(wsdGL_EJ_Recurrente.Rows.count, "J").End(xlUp).Row
    If lastUsedRow < 2 Then Exit Sub 'Empty List
    
    With lstEJRecurrente
        .ColumnHeads = True
        .ColumnCount = 2
        .ColumnWidths = "275; 25"
        .RowSource = wsdGL_EJ_Recurrente.Name & "!J2:K" & lastUsedRow
    End With
   
    Call InitialiserSurveillanceForm(Me, wrappers)
    
End Sub

Private Sub lstEJRecurrente_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim rowSelected As Long, DescEJAuto As String, NoEJAuto As Long
    
    rowSelected = lstEJRecurrente.ListIndex
    DescEJAuto = lstEJRecurrente.List(rowSelected, 0)
    NoEJAuto = lstEJRecurrente.List(rowSelected, 1)
    wshGL_EJ.Range("B2").Value = rowSelected
    
    Unload ufListeEJAuto
    
    Call modGL_EJ.ChargerEJRecurrenteDansEJ(DescEJAuto, NoEJAuto)

End Sub

Private Sub UserForm_Terminate()
    
    Unload Me

End Sub

