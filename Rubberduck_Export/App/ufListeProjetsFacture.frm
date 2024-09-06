VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeProjetsFacture 
   Caption         =   "Factures à préparer"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8700.001
   OleObjectBlob   =   "ufListeProjetsFacture.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListeProjetsFacture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).Row
    If lastUsedRow < 2 Then Exit Sub 'Empty List
    
    Dim arr() As Variant
    ReDim arr(1 To lastUsedRow - 1, 1 To 4)
    
    'Populate the array with non-contiguous columns
    Dim i As Long, nbRows As Long
    For i = 2 To lastUsedRow
        If UCase(ws.Cells(i, 26).value) <> "VRAI" Then 'Exclude those projects with isDetruite set to True
            nbRows = nbRows + 1
            arr(nbRows, 1) = ws.Cells(i, 2).value 'nomClient
            arr(nbRows, 2) = ws.Cells(i, 4).value 'date
            arr(nbRows, 3) = Fn_Pad_A_String(Format$(ws.Cells(i, 5).value, "#,##0.00$"), " ", 11, "L") 'Honoraires
            arr(nbRows, 4) = ws.Cells(i, 1).value 'ProjetID
        End If
    Next i
    
    If nbRows > 0 Then
    
        Call Array_2D_Resizer(arr, nbRows, UBound(arr, 2))
    
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
        Dim j As Long
        For i = LBound(arr, 1) To UBound(arr, 1)
            Me.lsbProjetsFacture.AddItem
            For j = LBound(arr, 2) To UBound(arr, 2)
                Me.lsbProjetsFacture.List(i - 1, j - 1) = arr(i, j)
            Next j
        Next i
    Else
        Unload Me
    End If
    
    'Cleanup Memory
    Set ws = Nothing
    Exit Sub

End Sub

Private Sub lsbProjetsFacture_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '2024-07-21 @ 16:38

    Dim rowSelected As Long
    Dim nomClient As String, dte As Date
    Dim honorairesTotal As Double
    Dim projetID As Long
    
    rowSelected = lsbProjetsFacture.ListIndex
    nomClient = lsbProjetsFacture.List(rowSelected, 0)
    dte = CDate(lsbProjetsFacture.List(rowSelected, 1))
    Debug.Print "lsbProjetsFacture_DblClick_70   dte = "; dte; "   "; TypeName(dte)
    honorairesTotal = lsbProjetsFacture.List(rowSelected, 2)
    projetID = lsbProjetsFacture.List(rowSelected, 3)
    
    Application.EnableEvents = False
    
    wshFAC_Brouillon.Range("B51").value = nomClient
    wshFAC_Brouillon.Range("B52").value = projetID
    Debug.Print "lsbProjetsFacture_DblClick_78   dte = "; dte; "   "; TypeName(dte)
    wshFAC_Brouillon.Range("B53").value = dte
    Debug.Print "lsbProjetsFacture_DblClick_80   wshFAC_Brouillon.Range(""B53"").value = "; wshFAC_Brouillon.Range("B53").value; "   "; TypeName(wshFAC_Brouillon.Range("B53").value)
    wshFAC_Brouillon.Range("B54").value = honorairesTotal
    
    Application.EnableEvents = True
    
    Unload ufListeProjetsFacture

End Sub

Private Sub UserForm_Terminate()
    
    Unload ufListeProjetsFacture
    
End Sub

