Attribute VB_Name = "modAdmin"
Option Explicit

Sub BrowseForMainSharedFolder()
    Dim SharedFolder As FileDialog
    Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With SharedFolder
        .Title = "Choisir le répertoire de données partagées, selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show <> -1 Then GoTo NotSelected
        wshAdmin.Range("F3").value = .SelectedItems(1) 'Full path for shared data files
NotSelected:
    End With
End Sub

Sub BrowseForFacturesPDFFolder()
    Dim PDFFolder As FileDialog
    Set PDFFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With PDFFolder
        .Title = "Choisir le répertoire des copies de facture (PDF), selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show <> -1 Then GoTo NoSelection
        wshAdmin.Range("F4").value = .SelectedItems(1) 'Full path for Invoice PDF directory
    End With
NoSelection:
End Sub

