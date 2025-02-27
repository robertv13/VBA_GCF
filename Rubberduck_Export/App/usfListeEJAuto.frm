VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfListeEJAuto 
   Caption         =   "Description des E/J récurrentes"
   ClientHeight    =   2910
   ClientLeft      =   7125
   ClientTop       =   6465
   ClientWidth     =   5355
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

    Dim RowSelected As Integer, DescEJAuto As String, NoEJAuto As Long
    RowSelected = lsbDescEJAuto.ListIndex
    DescEJAuto = lsbDescEJAuto.List(RowSelected, 0)
    NoEJAuto = lsbDescEJAuto.List(RowSelected, 1)
    wshJE.Range("B2").Value = RowSelected '2024-01-08 @ 13:58
    Unload usfListeEJAuto
    Call LoadJEAutoIntoJE(DescEJAuto, NoEJAuto)

End Sub

Private Sub UserForm_Initialize()
    
    Dim rowJEAutoDesc As Long
    Set MyListBoxClass = New clsCListboxAlign 'declare the class
    
    lsbDescEJAuto.ColumnCount = 2
    rowJEAutoDesc = wshEJRecurrente.Range("L9999").End(xlUp).row  'Last Row Used in wshEJRecurrente (Description Section)

    Dim r As Integer
    For r = 2 To rowJEAutoDesc
        With Me.lsbDescEJAuto
            .AddItem
            .List((r - 2), 0) = wshEJRecurrente.Range("L" & r).Value
            .List((r - 2), 1) = wshEJRecurrente.Range("M" & r).Value
        End With
    Next r

    'Corrige le format des colonnes
    MyListBoxClass.Left Me.lsbDescEJAuto, 1
    MyListBoxClass.Center Me.lsbDescEJAuto, 2
    
End Sub

Private Sub UserForm_Terminate()
    Unload Me
    'Clear the class declaration
    Set MyListBoxClass = Nothing
End Sub

