VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufFraisDivers 
   Caption         =   "Frais divers pour ce client"
   ClientHeight    =   1170
   ClientLeft      =   -30
   ClientTop       =   -240
   ClientWidth     =   7755
   OleObjectBlob   =   "ufFraisDivers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufFraisDivers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wrappers As Collection

Private Sub UserForm_Initialize() '2025-11-02 @ 10:51

    'Approximation : centré dans la fenêtre Excel
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    'Décalage vertical vers le bas (˜ 5 à 6 lignes Excel)
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2 + 100

    Call InitialiserSurveillanceForm(Me, wrappers)
    
End Sub
