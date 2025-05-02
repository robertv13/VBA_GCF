VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListe…critureGL 
   Caption         =   "Liste des Ècritures"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   OleObjectBlob   =   "ufListe…critureGL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListe…critureGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lsbListe…critureGL_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim ligne As Long
    
    ligne = lsbListe…critureGL.ListIndex
    
    If ligne <> -1 Then
        wshGL_EJ.Range("B3").value = lsbListe…critureGL.List(ligne, 0)
    End If
    
    Unload ufListe…critureGL
    
End Sub

