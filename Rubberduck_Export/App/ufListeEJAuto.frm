VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeEJAuto 
   Caption         =   "Choisir l'entrée récurrente à utiliser"
   ClientHeight    =   4500
   ClientLeft      =   7155
   ClientTop       =   6585
   ClientWidth     =   9000.001
   OleObjectBlob   =   "ufListeEJAuto.frx":0000
End
Attribute VB_Name = "ufListeEJAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private MyListBoxClass2 As CListBoxAlign

Private Sub UserForm_Initialize()
    
    Dim lastUsedRow As Long
    lastUsedRow = wshGL_EJ_Recurrente.Range("K999").End(xlUp).row
    If lastUsedRow < 2 Then Exit Sub 'Empty List
    
    With lsbEJ_Auto_Desc
        .ColumnHeads = True
        .ColumnCount = 2
        .ColumnWidths = "275; 25"
        .RowSource = "GL_EJ_Auto!K2:L" & lastUsedRow
    End With
   
End Sub

Private Sub lsbEJ_Auto_Desc_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim rowSelected As Long, DescEJAuto As String, NoEJAuto As Long
    rowSelected = lsbEJ_Auto_Desc.ListIndex
    DescEJAuto = lsbEJ_Auto_Desc.List(rowSelected, 0)
    NoEJAuto = lsbEJ_Auto_Desc.List(rowSelected, 1)
    wshGL_EJ.Range("B2").value = rowSelected '2024-01-08 @ 13:58
    Unload ufListeEJAuto
    Call Load_JEAuto_Into_JE(DescEJAuto, NoEJAuto)

End Sub

Private Sub UserForm_Terminate()
    Unload Me
    'Clear the class declaration
'    Set MyListBoxClass2 = Nothing
End Sub

