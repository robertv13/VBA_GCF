VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeDEBAuto 
   Caption         =   "Choisir le déboursé récurrent parmi la liste"
   ClientLeft      =   30
   ClientTop       =   -15
   OleObjectBlob   =   "ufListeDEBAuto.frx":0000
End
Attribute VB_Name = "ufListeDEBAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    
    Dim ws As Worksheet
    Set ws = wsdDEB_Recurrent
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "P").End(xlUp).Row
    If lastUsedRow < 2 Then Exit Sub 'Empty List
    Dim arr() As Variant
    ReDim arr(1 To (lastUsedRow - 1), 1 To 4) As Variant
    
    Application.ScreenUpdating = False
    
    With wsdDEB_Recurrent
        Dim i As Long
        For i = 2 To lastUsedRow
            arr(i - 1, 1) = .Range("P" & i).Value      'Deb Récurrent Auto
            arr(i - 1, 2) = .Range("Q" & i).Value      'Description
            arr(i - 1, 3) = Format$(.Range("R" & i).Value, "#,##0.00")     'Montant
            arr(i - 1, 3) = Space(10 - Len(arr(i - 1, 3))) & arr(i - 1, 3)
            arr(i - 1, 4) = .Range("S" & i).Value      'Date
        Next i
    End With
    
    'Nettoyer le listBox et le charger
    ufListeDEBAuto.lsbDEB_AutoDesc.Clear
    
    With ufListeDEBAuto.lsbDEB_AutoDesc
        .ColumnHeads = False
        .ColumnCount = 4
        .ColumnWidths = "30; 287; 65; 35"
        .MultiSelect = fmMultiSelectMulti
        .List = arr
    End With
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Private Sub lsbDEB_AutoDesc_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim rowSelected As Long
    rowSelected = lsbDEB_AutoDesc.ListIndex
    Dim noDEBAuto As Long
    noDEBAuto = lsbDEB_AutoDesc.List(rowSelected, 0)
    Dim descDEBAuto As String
    descDEBAuto = lsbDEB_AutoDesc.List(rowSelected, 1)
    
    wshDEB_Saisie.Range("B3").Value = rowSelected '2024-06-14 @ 07:23
    
    Unload ufListeDEBAuto
    
    Call ChargerDEBRecurrentDansSaisie(descDEBAuto, noDEBAuto)

End Sub

Private Sub UserForm_Terminate()

    Unload Me
    
End Sub

