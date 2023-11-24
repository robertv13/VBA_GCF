Option Explicit

Dim LastRow As Long, LastResultRow As Long, SelRow As Long
Dim ContRow As Long, ContCol As Long, ContID As Long
Dim ContName As String
Dim ContField As Control

Sub OpenContForm()
    ContactListLoad
    ContForm.picField.Picture = Nothing
    'wshContDB.Range("T3:V99999").ClearContents
    wshInterDB.Range("Q3:W99999").ClearContents
    ContForm.Show
End Sub

Sub ContactListLoad()
    With wshContDB
        LastRow = .Range("A99999").End(xlUp).Row
        If LastRow < 4 Then Exit Sub
        If ContForm.ContactActive.Value = True Then
            .Range("O3").Value = True
        Else
            .Range("O3").Value = "<>"
        End If
        .Range("P3").Value = "*" & ContForm.ContactSearch.Value & "*"
        .Range("A3:L" & LastRow).AdvancedFilter xlFilterCopy, _
                                                CriteriaRange:=.Range("O2:P3"), _
                                                CopyToRange:=.Range("T2:V2"), _
                                                Unique:=True
        LastResultRow = .Range("T99999").End(xlUp).Row
        If LastResultRow < 3 Then Exit Sub
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshContDB.Range("T3"), SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .SetRange wshContDB.Range("T3:V" & LastResultRow)
            .Apply
        End With
NoSort:
    End With
End Sub

Sub ContactClearFilter()
    ContForm.ContactSearch.Value = ""
End Sub

Sub ContactNew()
    With ContForm
        For ContCol = 1 To 11
            Set ContField = .Controls("Field" & ContCol)
            ContField.Value = ""
        Next ContCol
        .picField.Picture = Nothing
        .Field2.SetFocus
        .Field10.Value = True
        .ContactList.Value = ""
    End With
End Sub

Sub ContactLoad()
    With ContForm
        SelRow = .ContactList.ListIndex + 3
        ContRow = wshContDB.Range("V" & SelRow).Value
        If ContRow = 0 Then Exit Sub
        For ContCol = 1 To 11
            Set ContField = .Controls("Field" & ContCol)
            ContField.Value = wshContDB.Cells(ContRow, ContCol)
        Next ContCol
    End With
    ContactShowPicture
    InteractionListLoad 'Load Interactions
End Sub

Sub ContactShowPicture()
    Dim PicPath As String
    With ContForm
        If .Field11.Value = Empty Then Exit Sub
        PicPath = .Field11.Value 'Picture Path
        If Dir(PicPath, vbDirectory) = "" Then Exit Sub
        On Error Resume Next
        .picField.Picture = LoadPicture(PicPath)
        On Error GoTo 0
    End With
End Sub

Sub ContactBrowsePicture()
    Dim PicFile As FileDialog
    Set PicFile = Application.FileDialog(msoFileDialogFilePicker)
    With PicFile
        .Title = "Select a Contact Picture"
        .Filters.Clear
        .Filters.Add "Select a JPG Picture", "*.jpg, .jpeg, 1"
        If .Show <> -1 Then GoTo NoSelection
        ContForm.Field11.Value = .SelectedItems(1) 'Full Picture Path
        ContactShowPicture
NoSelection:
    End With
End Sub

Sub ContactClearPicture()
    ContForm.Field11.Value = ""
    ContForm.picField.Picture = Nothing
End Sub

Sub ContactSaveUpdate()
    With ContForm
        If .Field2.Value = Empty Then
            MsgBox "Please make sure to add in a Contact Name before saving"
            Exit Sub
        End If
        ContName = .Field2.Value 'Set Contact Name in memory
        If .Field1.Value = "" Then
            ContRow = wshContDB.Range("A99999").End(xlUp).Row + 1 'First Available Row
            On Error Resume Next
            ContID = Application.WorksheetFunction.Max(wshContDB.Range("ContID")) + 1
                                                            'Next available Contact ID
            On Error GoTo 0
            If ContID = 0 Then ContID = 1 'Set Default on No Data
            .Field1.Value = ContID
            wshContDB.Range("A" & ContRow).Value = ContID 'Set Contact ID
            wshContDB.Range("L" & ContRow).Value = "=ROW()" 'Set Row Number
            .ContactSearch.Value = "" 'Clear Any Search Filters
        Else                                     'Existing contact
            SelRow = .ContactList.ListIndex + 3  'Selected Index Item + 3
            ContRow = wshContDB.Range("V" & SelRow).Value 'Existing Contact DB
        End If
        If ContRow = 0 Then Exit Sub
        For ContCol = 2 To 11
            Set ContField = .Controls("Field" & ContCol)
            wshContDB.Cells(ContRow, ContCol).Value = ContField.Value 'Save to ContactDB
        Next ContCol
        ContactListLoad 'Reload Contact List
        SelRow = 0
        On Error Resume Next
        SelRow = wshContDB.Range("T:T").Find(ContName, , xlValues, xlWhole).Row
        On Error GoTo 0
        If SelRow <> 0 Then .ContactList.ListIndex = SelRow - 3 'Set Selected Contact
        MsgBox "Contact Saved"
    End With
End Sub

Sub ContactDelete()
    If MsgBox("Are you sure you want to delete this Contact?", vbYesNo, "Delete Contact") _
        = vbNo Then Exit Sub
    With ContForm
        If .Field1.Value = "" Then GoTo NotSaved 'Not yet recorded !
        SelRow = .ContactList.ListIndex + 3 'Selected Index Item + 3
        ContRow = wshContDB.Range("V" & SelRow).Value 'Existing Contact DB
        wshContDB.Range(ContRow & ":" & ContRow).EntireRow.Delete 'Delete Contact Row
NotSaved:
        ContactNew
        ContactListLoad 'Reload Contact List
    End With
End Sub
