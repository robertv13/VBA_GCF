VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfListeEJAuto 
   Caption         =   "Description des E/J récurrentes"
   ClientHeight    =   1920
   ClientLeft      =   7125
   ClientTop       =   6465
   ClientWidth     =   4500
   OleObjectBlob   =   "usfListeEJAuto.frx":0000
End
Attribute VB_Name = "usfListeEJAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyListBoxClass As clsCListboxAlign

Private Sub lsbDescEJAuto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim RowSelected As Integer
    RowSelected = lsbDescEJAuto.ListIndex
    Dim DescEJAuto As String
    Dim NoEJAuto As Long
    DescEJAuto = lsbDescEJAuto.List(RowSelected, 0)
    NoEJAuto = lsbDescEJAuto.List(RowSelected, 1)
    usfListeEJAuto.Hide
    Call LoadJEAutoIntoJE(DescEJAuto, NoEJAuto)

End Sub

Private Sub UserForm_Initialize()
    
    Dim rowJEAutoDesc As Long
    Set MyListBoxClass = New clsCListboxAlign 'declare the class
    
    lsbDescEJAuto.ColumnCount = 2
    rowJEAutoDesc = wshEJRecurrente.Range("L9999").End(xlUp).row + 1  'First Empty Row in wshEJRecurrente (Description Section)

    Dim r As Integer
    For r = 2 To rowJEAutoDesc
        With Me.lsbDescEJAuto
            .AddItem
            .List((r - 2), 0) = wshEJRecurrente.Range("L" & r).value
            .List((r - 2), 1) = wshEJRecurrente.Range("M" & r).value
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

