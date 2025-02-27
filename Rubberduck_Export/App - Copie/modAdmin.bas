Attribute VB_Name = "modAdmin"
Option Explicit

Sub ADMIN_DataFiles_Folder_Selection() '2024-03-28 @ 14:10

    Dim SharedFolder As FileDialog
    Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With SharedFolder
        .Title = "Choisir le répertoire de données partagées, selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F5").Value = .selectedItems(1)
        End If
    End With
    
End Sub

Sub ADMIN_PDF_Folder_Selection() '2024-03-28 @ 14:10

    Dim PDFFolder As FileDialog
    Set PDFFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With PDFFolder
        .Title = "Choisir le répertoire des copies de facture (PDF), selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F6").Value = .selectedItems(1)
        End If
    End With

End Sub

