VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContactFrm 
   Caption         =   "Add/Edit Product"
   ClientHeight    =   5565
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "ContactFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContactFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBtn_Click()
    Unload Me
End Sub

Private Sub SaveBtn_Click()
    Contact_SaveUpdate
End Sub

