﻿VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeEcritureGL 
   Caption         =   "Liste des écritures"
   ClientLeft      =   -60
   ClientTop       =   -255
   OleObjectBlob   =   "ufListeEcritureGL.frx":0000
End
Attribute VB_Name = "ufListeEcritureGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lsbListeEcritureGL_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim ligne As Long
    
    ligne = lsbListeEcritureGL.ListIndex
    
    If ligne <> -1 Then
        wshGL_EJ.Range("B3").Value = lsbListeEcritureGL.List(ligne, 0)
    End If
    
    Unload ufListeEcritureGL
    
End Sub


