VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufNonBillableTime 
   Caption         =   "Temps non facturable pour ce client - Veuillez s�lectionner les lignes � convertir en temps FACTURABLE"
   ClientHeight    =   4068
   ClientLeft      =   96
   ClientTop       =   276
   ClientWidth     =   8448.001
   OleObjectBlob   =   "ufNonBillableTime.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufNonBillableTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    'D�finir la couleur de fond du UserForm en utilisant le code RGB (198,224,190)
    Me.BackColor = RGB(198, 224, 190)
    
    'D�finir la couleur de fond du bouton en utilisant le code RGB (118,181,75)
    btnConvertir.BackColor = RGB(118, 181, 75)
    
End Sub

Private Sub btnConvertir_Click()

    Dim tecID As Long
    
    'Y a-t-il des lignes s�lectionn�es (donc � convertir en temps Facturable) ?
    Dim i As Integer, nbLigneS�lectionn�e As Integer
    For i = 0 To lsbNonBillable.ListCount - 1
        If lsbNonBillable.Selected(i) Then
            nbLigneS�lectionn�e = nbLigneS�lectionn�e + 1
            
            tecID = lsbNonBillable.List(i, 0)
            
            Call Convertir_NF_en_Facturable_Dans_BD(tecID)
            Call Convertir_NF_en_Facturable_Locally(tecID)
            
            Debug.Print "#096 - La ligne # " & i + 1 & " a �t� s�lectionn� - " & lsbNonBillable.List(i, 0)
        End If
    Next i

    'Informer du nombre de ligne convertie
    MsgBox "J'ai converti " & nbLigneS�lectionn�e & " ligne(s) en temps facturable", vbOKOnly + vbInformation, _
           "Sur une possibilit� de " & lsbNonBillable.ListCount & " ligne(s)..."
    
    'La conversion est termin�e
    Unload Me
    
End Sub

Private Sub lsbNonBillable_Change()

    Dim selectedCount As Integer, i As Integer

    selectedCount = 0

    For i = 0 To lsbNonBillable.ListCount - 1
        If lsbNonBillable.Selected(i) Then
            selectedCount = selectedCount + 1
        End If
    Next i

    btnConvertir.Visible = (selectedCount > 0)
    
End Sub

