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

Private MyListBoxClass2 As CListBoxAlign

Private Sub UserForm_Initialize()
    
    With lsbDescEJAuto
        .ColumnHeads = True
        .ColumnCount = 2
        .ColumnWidths = "275; 25"
        .RowSource = "EJ_Auto!K2:L12"
    End With
    
    'Class (clsCListboxAlign) to align column within a lisbox
    MyListBoxClass2.Left Me.lsbDescEJAuto, 1
    MyListBoxClass2.Right Me.lsbDescEJAuto, 2
    
End Sub

Private Sub UserForm_Activate()

    Dim rowJEAutoDesc As Long
    rowJEAutoDesc = wshGL_EJ_Recurrente.Range("L999").End(xlUp).row  'Last Row Used in wshGL_EJ_Recurrente (Description Section)

    Dim r As Integer
    Dim arr() As Variant
    
    ' Resize the array to hold the data
    ReDim arr(1 To rowJEAutoDesc - 1, 1 To 2)
    
    On Error Resume Next
    For r = 2 To rowJEAutoDesc
        'Store values in the array
        arr(r - 1, 1) = wshGL_EJ_Recurrente.Range("K" & r).value
        arr(r - 1, 2) = Pad_A_String(wshGL_EJ_Recurrente.Range("L" & r).value, " ", 2, "L")
    Next r
    On Error GoTo 0
    
    'Assign the entire array to the listbox
    ufListeEJAuto.lsbDescEJAuto.List = arr

End Sub

Private Sub lsbDescEJAuto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim rowSelected As Integer, DescEJAuto As String, NoEJAuto As Long
    rowSelected = lsbDescEJAuto.ListIndex
    DescEJAuto = lsbDescEJAuto.List(rowSelected, 0)
    NoEJAuto = lsbDescEJAuto.List(rowSelected, 1)
    wshGL_EJ.Range("B2").value = rowSelected '2024-01-08 @ 13:58
    Unload ufListeEJAuto
    Call Load_JEAuto_Into_JE(DescEJAuto, NoEJAuto)

End Sub

Private Sub UserForm_Terminate()
    Unload Me
    'Clear the class declaration
    Set MyListBoxClass2 = Nothing
End Sub

