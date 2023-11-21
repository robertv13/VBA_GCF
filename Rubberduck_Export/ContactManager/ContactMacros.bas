Attribute VB_Name = "ContactMacros"
Option Explicit

Dim LastRow As Long, LastResultRow As Long, SelRow As Long, ContactRow As Long, ContactCol As Long, ContactID As Long
Dim ContactName As String
Dim ContactField As Control

Sub OpenContactForm()
    ContactListLoad
    ContactForm.picField.Picture = Nothing
    wshInteractionsDB.Range("Q3:W99999").ClearContents
    ContactForm.Show
End Sub

Sub ContactListLoad()
    With wshContactsDB
        LastRow = .Range("A99999").End(xlUp).Row
        If LastRow < 4 Then Exit Sub
        If ContactForm.ContactActive.Value = True Then
            .Range("O3").Value = True
        Else
            .Range("O3").Value = "<>"
        End If
        .Range("P3").Value = "*" & ContactForm.ContactSearch.Value & "*"
        .Range("A3:L" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("O2:P3"), CopyToRange:=.Range("T2:V2"), Unique:=True
        LastResultRow = .Range("T99999").End(xlUp).Row
        If LastResultRow < 3 Then Exit Sub
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshContactsDB.Range("T3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange wshContactsDB.Range("T3:V" & LastResultRow)
            .Apply
        End With
NoSort:
    End With
End Sub

Sub Contact_ClearFilter()

    ContactForm.ContactSearch.Value = ""

End Sub

Sub Contact_New()

    With ContactForm
        For ContactCol = 1 To 11
            Set ContactField = .Controls("Field" & ContactCol)
            ContactField.Value = ""
        Next ContactCol
        .picField.Picture = Nothing
        .Field2.SetFocus
        .Field10.Value = True
        .ContactList.Value = ""
    End With

End Sub

Sub Contact_Load()

    With ContactForm
        SelRow = .ContactList.ListIndex + 3
        ContactRow = wshContactsDB.Range("V" & SelRow).Value
        If ContactRow = 0 Then Exit Sub
        For ContactCol = 1 To 11
            Set ContactField = .Controls("Field" & ContactCol)
            ContactField.Value = wshContactsDB.Cells(ContactRow, ContactCol)
        Next ContactCol
    End With
    Contact_ShowPicture

End Sub

Sub Contact_ShowPicture()

    Dim PicPath As String
    With ContactForm
        If .Field11.Value = Empty Then Exit Sub
        PicPath = .Field11.Value 'Picture Path
        If Dir(PicPath, vbDirectory) = "" Then Exit Sub
        On Error Resume Next
            .picField.Picture = LoadPicture(PicPath)
        On Error GoTo 0
    End With

End Sub

Sub Contact_BrowsePicture()

    Dim PicFile As FileDialog
    Set PicFile = Application.FileDialog(msoFileDialogFilePicker)
    With PicFile
        .Title = "Select a Contact Picture"
        .Filters.Clear
        .Filters.Add "Select a JPG Picture", "*.jpg, .jpeg, 1"
        If .Show <> -1 Then GoTo NoSelection
        ContactForm.Field11.Value = .SelectedItems(1) 'Full Picture Path
        Contact_ShowPicture
NoSelection:
    End With

End Sub

Sub Contact_ClearPicture()

    ContactForm.Field11.Value = ""
    ContactForm.picField.Picture = Nothing
    
End Sub

Sub Contact_SaveUpdate()

    With ContactForm
        If .Field2.Value = Empty Then
            MsgBox "Please make sure to add in a Contact Name before saving"
            Exit Sub
        End If
        ContactName = .Field2.Value 'Set Contact Name in memory
        If .Field1.Value = "" Then
            ContactRow = wshContactsDB.Range("A99999").End(xlUp).Row + 1 'First Available Row
            On Error Resume Next
            ContactID = Application.WorksheetFunction.Max(wshContactsDB.Range("ContactID")) + 1  'Next available Contact ID
            On Error GoTo 0
            If ContactID = 0 Then ContactID = 1 'Set Default on No Data
            .Field1.Value = ContactID
            wshContactsDB.Range("A" & ContactRow).Value = ContactID 'Set Contact ID
            wshContactsDB.Range("L" & ContactRow).Value = "=ROW()" 'Set Row Number
            .ContactSearch.Value = "" 'Clear Any Search Filters
        Else 'Existing contact
            SelRow = .ContactList.ListIndex + 3 'Selected Index Item + 3
            ContactRow = wshContactsDB.Range("V" & SelRow).Value 'Existing Contact DB
        End If
        If ContactRow = 0 Then Exit Sub
        For ContactCol = 2 To 11
            Set ContactField = .Controls("Field" & ContactCol)
            wshContactsDB.Cells(ContactRow, ContactCol).Value = ContactField.Value 'Save to ContactDB
        Next ContactCol
        ContactListLoad 'Reload Contact List
        SelRow = 0
        On Error Resume Next
            SelRow = wshContactsDB.Range("T:T").Find(ContactName, , xlValues, xlWhole).Row
        On Error GoTo 0
        If SelRow <> 0 Then .ContactList.ListIndex = SelRow - 3 'Set Selected Contact
        MsgBox "Contact Saved"
    End With

End Sub
