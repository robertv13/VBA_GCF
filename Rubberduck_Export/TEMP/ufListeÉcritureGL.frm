VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListe…critureGL 
   Caption         =   "Liste des Ècritures"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10035
   OleObjectBlob   =   "ufListe…critureGL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListe…critureGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lbListe…critureGL_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim ligne As Long
    
    ligne = lbListe…critureGL.ListIndex
    
    If ligne <> -1 Then
        wshGL_EJ.Range("B3").Value = lbListe…critureGL.List(ligne, 0)
    End If
    
    Unload ufListe…critureGL
    
End Sub

