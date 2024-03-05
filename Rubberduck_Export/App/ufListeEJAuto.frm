VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeEJAuto 
   Caption         =   "Choisir l'entrée récurrente à utiliser"
   ClientHeight    =   4500
   ClientLeft      =   7125
   ClientTop       =   6465
   ClientWidth     =   7200
   OleObjectBlob   =   "ufListeEJAuto.frx":0000
End
Attribute VB_Name = "ufListeEJAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyListBoxClass As clsCListboxAlign

Private Sub UserForm_Initialize()
    
'    Set MyListBoxClass = New clsCListboxAlign 'declare the class
    
    With lsbDescEJAuto
        .ColumnHeads = True
        .ColumnCount = 2
        .ColumnWidths = "275; 25"
        .RowSource = "EJ_Auto!K2:L12"
    End With
    
End Sub

Private Sub UserForm_Activate()

    Dim rowJEAutoDesc As Long
    rowJEAutoDesc = wshEJRecurrente.Range("L999").End(xlUp).row  'Last Row Used in wshEJRecurrente (Description Section)

    Dim r As Integer
    Dim arr() As Variant
    
    ' Resize the array to hold the data
    ReDim arr(1 To rowJEAutoDesc - 1, 1 To 2)
    
    On Error Resume Next
    For r = 2 To rowJEAutoDesc
        Debug.Print "Row: " & r
        Debug.Print "Value in K column: " & wshEJRecurrente.Range("K" & r).value
        Debug.Print "Value in L column: " & wshEJRecurrente.Range("L" & r).value
        
        ' Store values in the array
        arr(r - 1, 1) = wshEJRecurrente.Range("K" & r).value
        arr(r - 1, 2) = wshEJRecurrente.Range("L" & r).value
    Next r
    
    'Assign the entire array to the listbox
    ufListeEJAuto.lsbDescEJAuto.List = arr

End Sub

Private Sub lsbDescEJAuto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim RowSelected As Integer, DescEJAuto As String, NoEJAuto As Long
    RowSelected = lsbDescEJAuto.ListIndex
    DescEJAuto = lsbDescEJAuto.List(RowSelected, 0)
    NoEJAuto = lsbDescEJAuto.List(RowSelected, 1)
    wshJE.Range("B2").value = RowSelected '2024-01-08 @ 13:58
    Unload ufListeEJAuto
    Call Load_JEAuto_Into_JE(DescEJAuto, NoEJAuto)

End Sub

Private Sub UserForm_Terminate()
    Unload Me
    'Clear the class declaration
    Set MyListBoxClass = Nothing
End Sub

