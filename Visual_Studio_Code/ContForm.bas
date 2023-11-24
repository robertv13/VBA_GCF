Option Explicit

Private Sub ClearFilterButton_Click()
    ContactClearFilter
End Sub

Private Sub cmdBrowse_Click()
    ContactBrowsePicture
End Sub

Private Sub cmdClearPicture_Click()
    ContactClearPicture
End Sub

Private Sub cmdDeleteContact_Click()
    ContactDelete
End Sub

Private Sub cmdDeleteInteraction_Click()
    InteractionDelete
End Sub

Private Sub cmdNewContact_Click()
    ContactNew
End Sub

Private Sub cmdNewInteraction_Click()
    InteractionNew
End Sub

Private Sub cmdSaveContact_Click()
    ContactSaveUpdate
End Sub

Private Sub cmdSaveInteraction_Click()
    InteractionSaveUpdate
End Sub

Private Sub ContactActive_Click()
    ContactListLoad
End Sub

Private Sub ContactList_Click()
    ContactLoad
End Sub

Private Sub ContactSearch_Change()
    ContactListLoad
End Sub

Private Sub Inter6_Change()
    Inter6.Value = Format(Inter6.Value, "[$-en-US]h:mm AM/PM;@")
End Sub

Private Sub Inter7_Change()
    Inter7.Value = Format(Inter7.Value, "h:mm;@")
End Sub

Private Sub InterList_Click()
    InteractionLoad
End Sub
