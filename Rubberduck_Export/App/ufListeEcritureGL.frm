VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListeEcritureGL 
   Caption         =   "Liste des écritures"
   ClientHeight    =   7320
   ClientLeft      =   -60
   ClientTop       =   -255
   ClientWidth     =   13500
   OleObjectBlob   =   "ufListeEcritureGL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListeEcritureGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wrappers As Collection

Private Sub UserForm_Initialize()

    Call InitialiserSurveillanceForm(Me, wrappers)
    
End Sub

Private Sub lstListeEcritureGL_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim ligne As Long
    
    ligne = lstListeEcritureGL.ListIndex
    
    If ligne <> -1 Then
        wshGL_EJ.Range("B3").Value = lstListeEcritureGL.List(ligne, 0)
    End If
    
    Unload ufListeEcritureGL
    
End Sub

