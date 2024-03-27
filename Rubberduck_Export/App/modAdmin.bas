Attribute VB_Name = "modAdmin"
Option Explicit

Sub BrowseForMainSharedFolder()
    Dim SharedFolder As FileDialog
    Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With SharedFolder
        .Title = "Choisir le r�pertoire de donn�es partag�es, selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show <> -1 Then GoTo NotSelected
        wshAdmin.Range("F5").value = .SelectedItems(1)
NotSelected:
    End With
'    FolderSharedData.value = "P:\Admin-GC\GC Fiscalit� Plus Inc\Informatique RMV\GC_FISCALIT�\DataFiles"
    
End Sub

Sub BrowseForFacturesPDFFolder()
    Dim PDFFolder As FileDialog
    Set PDFFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With PDFFolder
        .Title = "Choisir le r�pertoire des copies de facture (PDF), selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show <> -1 Then GoTo NoSelection
        wshAdmin.Range("F6").value = .SelectedItems(1)
    End With
NoSelection:
'    FolderPDFInvoice = "P:\Admin-GC\GC Fiscalit� Plus Inc\Informatique RMV\GC_FISCALIT�\Factures_PDF"
End Sub

