VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeEJAuto 
   Caption         =   "Choisir l'entrée récurrente à utiliser"
   ClientHeight    =   5325
   ClientLeft      =   7215
   ClientTop       =   6810
   ClientWidth     =   7200
   OleObjectBlob   =   "ufListeEJAuto.frx":0000
End
Attribute VB_Name = "ufListeEJAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
    Dim lastUsedRow As Long
    lastUsedRow = wsdGL_EJ_Recurrente.Cells(wsdGL_EJ_Recurrente.Rows.count, "J").End(xlUp).row
    If lastUsedRow < 2 Then Exit Sub 'Empty List
    
    With lsbEJ_AutoDesc
        .ColumnHeads = True
        .ColumnCount = 2
        .ColumnWidths = "275; 25"
        .RowSource = wsdGL_EJ_Recurrente.Name & "!J2:K" & lastUsedRow
    End With
   
End Sub

Private Sub lsbEJ_AutoDesc_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim rowSelected As Long, DescEJAuto As String, NoEJAuto As Long
    rowSelected = lsbEJ_AutoDesc.ListIndex
    DescEJAuto = lsbEJ_AutoDesc.List(rowSelected, 0)
    NoEJAuto = lsbEJ_AutoDesc.List(rowSelected, 1)
    wshGL_EJ.Range("B2").Value = rowSelected '2024-01-08 @ 13:58
    Unload ufListeEJAuto
    Call Load_JEAuto_Into_JE(DescEJAuto, NoEJAuto)

End Sub

Private Sub UserForm_Terminate()
    
    Unload Me

End Sub

