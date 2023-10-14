VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfListeEJAuto 
   Caption         =   "Description des E/J récurrentes"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5085
   OleObjectBlob   =   "usfListeEJAuto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfListeEJAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyListBoxClass As cListBoxAlign

Private Sub lsbDescEJAuto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim RowSelected As Integer
    RowSelected = lsbDescEJAuto.ListIndex
    Dim DescEJAuto As String
    Dim NoEJAuto As Long
    DescEJAuto = lsbDescEJAuto.List(RowSelected, 0)
    NoEJAuto = lsbDescEJAuto.List(RowSelected, 1) + 1
    usfListeEJAuto.Hide
    Call LoadJEAutoIntoJE(DescEJAuto, NoEJAuto)

End Sub

Private Sub UserForm_Initialize()
    
    Dim rowJEAutoDesc As Long
    Set MyListBoxClass = New cListBoxAlign 'declare the class
    
    lsbDescEJAuto.ColumnCount = 2
    rowJEAutoDesc = wshJERecurrente.Range("K9999").End(xlUp).row + 1  'First Empty Row in wshEJRecurrente (Description Section)

    Dim r As Integer
    For r = 2 To rowJEAutoDesc
        With Me.lsbDescEJAuto
            .AddItem
            .List((r - 2), 0) = wshJERecurrente.Range("K" & r).Value
            .List((r - 2), 1) = wshJERecurrente.Range("L" & r).Value
        End With
    Next r

    'Corrige le format des colonnes
    MyListBoxClass.Left Me.lsbDescEJAuto, 1
    MyListBoxClass.Center Me.lsbDescEJAuto, 2
    
End Sub

Private Sub UserForm_Terminate()
    'clear the class declaration
    Set MyListBoxClass = Nothing
End Sub
