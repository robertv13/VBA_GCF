VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufExitDisable 
   Caption         =   "Sortie non autorisée via le fermeture d'Excel"
   ClientHeight    =   2310
   ClientLeft      =   -30
   ClientTop       =   -255
   ClientWidth     =   7755
   OleObjectBlob   =   "ufExitDisable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufExitDisable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub shpOK_Click() '2024-07-05 @ 06:25

    Me.Hide

End Sub

Private Sub UserForm_Initialize()

    Me.Label1.Caption = "Pour quitter cette application, vous devez OBLIGATOIREMENT" & vbCrLf & vbCrLf & _
                     "utiliser l'option prévue à cet effet (en bas à gauche, du menu principal)"

End Sub
