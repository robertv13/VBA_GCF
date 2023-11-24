Attribute VB_Name = "AppSyncMacros"
Option Explicit

Sub UpdateFile()
    Dim FileName As String
    Dim FilePath As String
    Dim LongFileName As String
    Dim fso As Object
    Dim oFile As Object
    Dim TCFile As String
    Dim SyncText As String
    Dim CurrentUser As String
    Dim SheetName As String
    Dim CellAddress As String
    Dim CellText As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    CurrentUser = Sheet2.Range("CurrentUser").Value
    If Dir(Sheet1.Range("SharedFolder").Value, vbDirectory) = Empty Then Exit Sub
    FilePath = Sheet1.Range("SharedFolder").Value & "\" & CurrentUser & "\"
    If Dir(FilePath, vbDirectory) = Empty Then
        fso.createfolder (FilePath) 'Add User Folder if needed
        Exit Sub 'Nothing to sync
    End If
    Sheet2.Range("B3").Value = True 'Set syncing to TRUE (bringing in changes)
    FileName = Dir(FilePath & "*.txt")
    Do While Len(FileName) > 0 'Start of Loop
        LongFileName = FilePath & FileName
        Open LongFileName For Input As #1
        Line Input #1, SyncText
        Close #1
        SheetName = Left(SyncText, InStr(SyncText, ",") - 1)
        CellAddress = Mid(SyncText, InStr(SyncText, ",") + 1, InStr(SyncText, ":") - InStr(SyncText, ",") - 1) '
        CellText = Right(SyncText, Len(SyncText) - InStr(SyncText, ":"))
        ThisWorkbook.Sheets(SheetName).Range(CellAddress).Value = CellText
        Kill (LongFileName)
        FileName = Dir() 'Clear the current file name
    Loop
    Sheet2.Range("B3").Value = False 'Set syncing to FALSE
    Set fso = Nothing
    Set oFile = Nothing
End Sub

Sub BrowseSharedFolder()
    Dim SharedFolder As FileDialog
    Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With SharedFolder
        .Title = "Select a Shared Folder"
        If .Show <> -1 Then GoTo NoSelection
        Sheet1.Range("P3").Value = .SelectedItems(1)
NoSelection:
    End With
End Sub
