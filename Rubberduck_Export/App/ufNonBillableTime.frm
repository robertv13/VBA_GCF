VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufNonBillableTime 
   Caption         =   "Temps non facturable pour ce client - Veuillez sélectionner les lignes à convertir en temps FACTURABLE"
   ClientHeight    =   4050
   ClientLeft      =   -30
   ClientTop       =   -255
   ClientWidth     =   13200
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

    'Définir la couleur de fond du UserForm en utilisant le code RGB (198,224,190)
    Me.BackColor = RGB(198, 224, 190)
    
    'Définir la couleur de fond du bouton en utilisant le code RGB (118,181,75)
    shpConvertir.BackColor = RGB(118, 181, 75)
    
End Sub

Private Sub shpConvertir_Click()

    Call ConvertirNFenFacturable
    
End Sub

Private Sub ConvertirNFenFacturable()

    Dim tecID As Long
    
    'Y a-t-il des lignes sélectionnées (donc à convertir en temps Facturable) ?
    Dim i As Integer, nbLigneSélectionnée As Integer
    For i = 0 To lstNonBillable.ListCount - 1
        If lstNonBillable.Selected(i) Then
            nbLigneSélectionnée = nbLigneSélectionnée + 1
            
            tecID = lstNonBillable.List(i, 0)
            
            Call ConvertirNFenFacturableBDMaster(tecID)
            Call ConvertirNFenFacturableBDLocale(tecID)
            
            Debug.Print "#096 - La ligne # " & i + 1 & " a été sélectionné - " & lstNonBillable.List(i, 0)
        End If
    Next i

    'Informer du nombre de ligne convertie
    MsgBox "J'ai converti " & nbLigneSélectionnée & " ligne(s) en temps facturable", vbOKOnly + vbInformation, _
           "Sur une possibilité de " & lstNonBillable.ListCount & " ligne(s)..."
    
    'La conversion est terminée
    Unload Me
    
End Sub

Private Sub lstNonBillable_Change()

    Dim selectedCount As Integer, i As Integer

    selectedCount = 0

    For i = 0 To lstNonBillable.ListCount - 1
        If lstNonBillable.Selected(i) Then
            selectedCount = selectedCount + 1
        End If
    Next i

    shpConvertir.Visible = (selectedCount > 0)
    
End Sub

