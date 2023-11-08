Attribute VB_Name = "Contact_Macros"
Option Explicit
Dim ContactRow As Long, ContactCol As Long, LastRow As Long, LastResultRow As Long
Dim ContactFld As control

Sub Contact_AddNew()
    Schedule.Range("M5").ClearContents           'Clear Selected Contact Row and Contact DB Row
    ContactFrm.Show
End Sub

Sub Contact_Edit()
    With Schedule
        If .Range("B13").Value = Empty Then
            MsgBox "Please select on a Contact to edit"
            Exit Sub
        End If
        ContactRow = .Range("B13").Value         'Contact Row
        With ContactFrm
            For ContactCol = 2 To 8
                Set ContactFld = .Controls("Field" & ContactCol - 1)
                ContactFld.Value = ContactsDB.Cells(ContactRow, ContactCol).Value 'Add Contact Data
            Next ContactCol
            .Show                                'Display User Form
        End With
    End With
End Sub

Sub Contact_SaveUpdate()
    'Check For Required field
    With ContactFrm
        If .Field1.Value = Empty Then
            MsgBox "Please make sure to add a Contact Name before saving"
            Exit Sub
        End If
        If Schedule.Range("B13").Value = Empty Then 'New Contact
            ContactRow = ContactsDB.Range("A9999").End(xlUp).Row + 1 'First Avail. Row
            ContactsDB.Range("A" & ContactRow).Value = Schedule.Range("B14").Value 'Next Contact ID
            ContactsDB.Range("I" & ContactRow).Value = "=Row()"
        Else                                     'Existing Contact.
            ContactRow = Schedule.Range("B13").Value 'Contact Row
        End If
        For ContactCol = 2 To 8
            Set ContactFld = .Controls("Field" & ContactCol - 1)
            ContactsDB.Cells(ContactRow, ContactCol).Value = ContactFld.Value
        Next ContactCol
        Schedule.Range("M5").Value = .Field1.Value 'Set Name (in case of changes)
    End With
    Unload ContactFrm
    Contact_SortNames                            'Resort Names
End Sub

Sub Contact_SortNames()
    With ContactsDB
        LastRow = .Range("A99999").End(xlUp).Row
        If LastRow < 4 Then Exit Sub
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        .Range("B3:B" & LastRow).AdvancedFilter xlFilterCopy, , CopyToRange:=.Range("L2"), Unique:=True
        LastResultRow = .Range("L99999").End(xlUp).Row
        If LastResultRow < 4 Then Exit Sub       'No Sort On 1 Row Of data
        'Sort By Name & remove Blank Rows
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=ContactsDB.Range("L3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Sort
            .SetRange ContactsDB.Range("L3:L" & LastResultRow) 'Set Range
            .Apply                               'Apply Sort
        End With
    End With
End Sub

Sub Contact_Delete()
    If MsgBox("Are you sure you want to delete this Contact?", vbYesNo, "Delete Contact") = vbNo Then Exit Sub
    If Schedule.Range("B13").Value = Empty Then Exit Sub
    ContactRow = Schedule.Range("B13").Value     'Contact Row
    For ContactCol = 2 To 9
        ContactsDB.Cells(ContactRow, ContactCol).ClearContents
    Next ContactCol
    Schedule.Range("M5").ClearContents
    Contact_SortNames
End Sub

