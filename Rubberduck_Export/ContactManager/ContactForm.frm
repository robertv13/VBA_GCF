VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContactForm 
   Caption         =   "Contact Manager"
   ClientHeight    =   8520.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18060
   OleObjectBlob   =   "ContactForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContactForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearFilterButton_Click()

    Contact_ClearFilter
    
End Sub

Private Sub cmdBrowse_Click()

    Contact_BrowsePicture

End Sub

Private Sub cmdClearPicture_Click()

    Contact_ClearPicture

End Sub

Private Sub cmdDeleteContact_Click()
    
    Contact_Delete
    
End Sub

Private Sub cmdNewContact_Click()

    Contact_New

End Sub

Private Sub cmdSaveContact_Click()

    Contact_SaveUpdate

End Sub

Private Sub ContactActive_Click()

    ContactListLoad

End Sub

Private Sub ContactList_Click()

    Contact_Load

End Sub

Private Sub ContactSearch_Change()

    ContactListLoad

End Sub

Private Sub UserForm_Click()

End Sub
