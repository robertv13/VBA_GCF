Attribute VB_Name = "InteractionMacros"
Option Explicit

Dim LastRow As Long, LastResultRow As Long, SelRow As Long
Dim InterRow As Long, InterCol As Long, InterID As Long
Dim InterName As String
Dim InterField As Control

Sub InteractionListLoad()
    If ContForm.Field1.Value = "" Then Exit Sub 'Exit on No Contact ID
    With wshInterDB
        LastRow = .Range("A99999").End(xlUp).Row
        If LastRow < 4 Then Exit Sub
        .Range("M3").Value = ContForm.Field1 'Set Contact ID Criteria
        .Range("A3:I" & LastRow).AdvancedFilter xlFilterCopy, _
                                                CriteriaRange:=.Range("M2:M3"), _
                                                CopyToRange:=.Range("Q2:W2"), _
                                                Unique:=True
        LastResultRow = .Range("Q99999").End(xlUp).Row
        If LastResultRow < 3 Then Exit Sub
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshInterDB.Range("Q3"), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlDescending, _
                            DataOption:=xlSortTextAsNumbers
            .SortFields.Add Key:=wshInterDB.Range("R3"), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlDescending, _
                            DataOption:=xlSortTextAsNumbers
            .SetRange wshInterDB.Range("Q3:W" & LastResultRow) 'Set Range
            .Apply
        End With
NoSort:
        .Calculate
    End With
End Sub

Sub InteractionLoad()
    With ContForm
        SelRow = .InterList.ListIndex + 3 'Selected Row from ListBox
        InterRow = wshInterDB.Range("W" & SelRow).Value 'InteractionDB Row
        If InterRow = 0 Then Exit Sub
        Application.ScreenUpdating = False
        .Inter1.Value = wshInterDB.Cells(InterRow, 1).Value 'Interaction ID
        For InterCol = 3 To 8
            Set InterField = .Controls("Inter" & InterCol)
            InterField.Value = wshInterDB.Cells(InterRow, InterCol).Value
        Next InterCol
        .Inter6.Value = Format(.Inter6.Value, "[$-en-US]h:mm AM/PM;@")
        .Inter7.Value = Format(.Inter7.Value, "h:mm;@")
        Application.ScreenUpdating = True
    End With
End Sub

Sub InteractionNew()
    With ContForm
        .Inter1.Value = "" 'Interaction ID
        For InterCol = 3 To 8
            Set InterField = .Controls("Inter" & InterCol)
            InterField.Value = ""
        Next InterCol
        .Inter3.SetFocus 'Interaction get Focus
        .InterList.Value = ""
    End With
End Sub

Sub InteractionSaveUpdate()
    With ContForm
        If .Inter3.Value = "" Then
            MsgBox "Please make sure to add in an Interaction Name before saving"
            Exit Sub
        End If
        InterName = .Inter3.Value 'Set Interaction Name in memory
        If .Inter1.Value = "" Then 'Check Interaction ID
            InterRow = wshInterDB.Range("A99999").End(xlUp).Row + 1 'First Available Row
            On Error Resume Next
            InterID = Application.WorksheetFunction.Max(wshInterDB.Range("InterID")) + 1 'Next available Interaction ID
            On Error GoTo 0
            If InterID = 0 Then InterID = 1 'Set Default on No Data
            .Inter1.Value = InterID
            wshInterDB.Range("A" & InterRow).Value = InterID 'Set Interaction ID
            wshInterDB.Range("B" & InterRow).Value = .Field1.Value 'Set Contact ID
            wshInterDB.Range("I" & InterRow).Value = "=ROW()" 'Set Row Number
        Else 'Existing Interaction
            SelRow = .InterList.ListIndex + 3 'Selected Index Item + 3
            InterRow = wshInterDB.Range("W" & SelRow).Value 'Existing Interaction DB
        End If
        If InterRow = 0 Then Exit Sub
        For InterCol = 3 To 8
            Set InterField = .Controls("Inter" & InterCol)
            wshInterDB.Cells(InterRow, InterCol).Value = InterField.Value 'Save to InteractionDB
        Next InterCol
        InteractionListLoad 'Reload Interaction List
        wshInterDB.Calculate
        SelRow = 0
        On Error Resume Next
        SelRow = wshInterDB.Range("Q:Q").Find(InterName, , xlValues, xlWhole).Row
        On Error GoTo 0
        If SelRow <> 0 Then .InterList.ListIndex = SelRow - 3 'Set Selected Interaction
        MsgBox "Interaction Saved"
    End With
End Sub

Sub InteractionDelete()
    If MsgBox("Are you sure you want to delete this Interaction?", vbYesNo, _
        "Delete Interaction") = vbNo Then Exit Sub
    With ContForm
        If .Inter1.Value = "" Then GoTo NotSaved 'Not yet recorded !
        SelRow = .InterList.ListIndex + 3 'Selected Index Item + 3
        InterRow = wshInterDB.Range("W" & SelRow).Value 'Existing Contact DB
        Debug.Print InterRow
        If InterRow = 0 Then GoTo NotSaved
        wshInterDB.Range(InterRow & ":" & InterRow).EntireRow.Delete 'Delete Interaction Row
NotSaved:
        InteractionNew
        InteractionListLoad 'Reload Interaction List
        wshInterDB.Calculate
    End With
End Sub

