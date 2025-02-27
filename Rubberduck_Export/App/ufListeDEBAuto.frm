VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeDEBAuto 
   Caption         =   "Choisir le déboursé récurrent"
   ClientHeight    =   4485
   ClientLeft      =   180
   ClientTop       =   690
   ClientWidth     =   7515
   OleObjectBlob   =   "ufListeDEBAuto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListeDEBAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
    Dim lastUsedRow As Long
    lastUsedRow = wshDEB_Récurrent.Range("O999").End(xlUp).row
    If lastUsedRow < 2 Then Exit Sub 'Empty List
    
    With lsbDEB_AutoDesc
        .ColumnHeads = False
        .ColumnCount = 3
        .ColumnWidths = "40; 260; 30"
        .RowSource = wshDEB_Récurrent.Name & "!O2:Q" & lastUsedRow
    End With
   
End Sub

Private Sub lsbDEB_AutoDesc_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim rowSelected As Long, DescDEBAuto As String, NoDEBAuto As Long
    rowSelected = lsbDEB_AutoDesc.ListIndex
    NoDEBAuto = lsbDEB_AutoDesc.List(rowSelected, 0)
    DescDEBAuto = lsbDEB_AutoDesc.List(rowSelected, 1)
    wshDEB_Saisie.Range("B3").Value = rowSelected '2024-06-14 @ 07:23
    Unload ufListeDEBAuto
    Call Load_DEB_Auto_Into_JE(DescDEBAuto, NoDEBAuto)

End Sub

Private Sub UserForm_Terminate()

    Unload Me
    
End Sub

