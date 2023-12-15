Attribute VB_Name = "modAdmin"
Option Explicit

Dim PicFolder As FileDialog

Sub BrowseForMainSharedFolder()
    Dim SharedFolder As FileDialog
    Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With SharedFolder
        .Title = "Choisir le répertoire de données partagées, selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show <> -1 Then GoTo NotSelected
'        If InStr(.SelectedItems(1), "Dropbox") = 0 Then '2023-12-15 @ 07:28
'            MsgBox "Veuillez vous assurer de choisir un répertoire à l'intérieur de Dropbox."
'            Exit Sub
'        End If
        wshAdmin.Range("F3").value = .SelectedItems(1) 'Full Folder Path
NotSelected:
    End With
End Sub

Sub BrowseForEmployeePicFolder()
    Set PicFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With PicFolder
        .Title = "Browse Product Picture Folder"
        .AllowMultiSelect = False
        If .show <> -1 Then GoTo NoSelection
        wshAdmin.Range("C4").value = .SelectedItems(1)
    End With
NoSelection:
End Sub

Sub BrowseForProductPicFolder()
    Set PicFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With PicFolder
        .Title = "Browse Product Picture Folder"
        .AllowMultiSelect = False
        If .show <> -1 Then GoTo NoSelection
        wshAdmin.Range("C5").value = .SelectedItems(1)
    End With
NoSelection:
End Sub

Sub BrowseForUserPicFolder()
    Set PicFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With PicFolder
        .Title = "Browse User Picture Folder"
        .AllowMultiSelect = False
        If .show <> -1 Then GoTo NoSelection
        wshAdmin.Range("C6").value = .SelectedItems(1)
    End With
NoSelection:
End Sub

