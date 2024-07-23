VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeProjetsFacture 
   Caption         =   "Factures à préparer"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8730.001
   OleObjectBlob   =   "ufListeProjetsFacture.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListeProjetsFacture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lsbProjetsFacture_Click()

End Sub

Private Sub UserForm_Initialize()
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).row
    If lastUsedRow < 2 Then Exit Sub 'Empty List
    
    Dim arr() As Variant
    ReDim arr(1 To lastUsedRow - 1, 1 To 4)
    
    'Populate the array with non-contiguous columns
    Dim i As Integer
    For i = 2 To lastUsedRow
        arr(i - 1, 1) = ws.Cells(i, 2).value 'nomClient
        arr(i - 1, 2) = ws.Cells(i, 4).value 'date
        arr(i - 1, 3) = Fn_Pad_A_String(Format(ws.Cells(i, 5).value, "#,##0.00$"), " ", 11, "L") 'Honoraires
        arr(i - 1, 4) = ws.Cells(i, 1).value 'ProjetID
    Next i
    
    'Sort the list
    Call Array_2D_Bubble_Sort(arr)
    
    'Clear the ListBox
    Me.lsbProjetsFacture.clear

    With lsbProjetsFacture
        .ColumnHeads = True
        .ColumnCount = 4
        .ColumnWidths = "225; 68; 90; 15"
    End With
        
    'Populate the ListBox with the array
    Dim j As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        Me.lsbProjetsFacture.AddItem
        For j = LBound(arr, 2) To UBound(arr, 2)
            Me.lsbProjetsFacture.List(i - 1, j - 1) = arr(i, j)
        Next j
    Next i

    'Cleanup Memory
    Set ws = Nothing
    
End Sub

Private Sub lsbProjetsFacture_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '2024-07-21 @ 16:38

    Dim rowSelected As Integer
    Dim nomClient As String, dte As String
    Dim honorairesTotal As Double
    Dim projetID As Long
    
    rowSelected = lsbProjetsFacture.ListIndex
    nomClient = lsbProjetsFacture.List(rowSelected, 0)
    dte = lsbProjetsFacture.List(rowSelected, 1)
    honorairesTotal = lsbProjetsFacture.List(rowSelected, 2)
    projetID = lsbProjetsFacture.List(rowSelected, 3)
    
    wshFAC_Brouillon.Range("B51").value = nomClient
    wshFAC_Brouillon.Range("B52").value = projetID
    wshFAC_Brouillon.Range("B53").value = dte
    wshFAC_Brouillon.Range("B54").value = honorairesTotal
    
    Unload ufListeProjetsFacture

End Sub

Private Sub UserForm_Terminate()
    
    Unload ufListeProjetsFacture
    
End Sub


