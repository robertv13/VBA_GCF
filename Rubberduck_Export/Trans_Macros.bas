Attribute VB_Name = "Trans_Macros"
Option Explicit

Sub ShowSaveBtn()
With Sheet1
    Dim SelRow As Long
    If .Range("B2").Value = Empty Then Exit Sub
    SelRow = .Range("B2").Value
    With .Shapes("SaveBtn")
        .Left = Sheet1.Range("J" & SelRow).Left
        .Top = Sheet1.Range("J" & SelRow).Top
        .IncrementLeft 3
       .IncrementTop 2
        .Visible = msoCTrue
    End With
End With
End Sub

Sub LoadTransactions()
Dim LastResultsRow As Long
Dim ResultsRow As Long
Dim TransRow As Long
Dim LastTransDataRow As Long
Sheet1.Range("B1").Value = True 'Set Load To True
StopCalc
Sheet1.Range("C8:H99999").ClearContents
With Sheet3 'Transaction Sheet
    .Range("V3:AF999999").ClearContents 'Clear Results
    LastTransDataRow = .Range("E999999").End(xlUp).Row
    If LastTransDataRow < 5 Then GoTo NoData
    .Range("D4:N" & LastTransDataRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("Q2:T4"), CopyToRange:=.Range("V2:AF2"), Unique:=False
    LastResultsRow = .Range("W999999").End(xlUp).Row
    'Sort List Based on Date
    .Sort.SortFields.Clear
    .Sort.SortFields.Add Key:=.Range("W3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With .Sort
            .SetRange Range("V3:AF" & LastResultsRow)
            .Apply
    End With
    TransRow = 8
    For ResultsRow = 3 To LastResultsRow Step 2
        Sheet1.Range("C" & TransRow & ":H" & TransRow).Value = .Range("V" & ResultsRow & ":AA" & ResultsRow).Value
        Sheet1.Range("D" & TransRow + 1 & ":H" & TransRow + 1).Value = .Range("AB" & ResultsRow + 1 & ":AF" & ResultsRow + 1).Value
        TransRow = TransRow + 2
    Next ResultsRow
NoData:
    Sheet1.Range("B1").Value = False 'Set Load To False
    ResetCalc
End With
End Sub
