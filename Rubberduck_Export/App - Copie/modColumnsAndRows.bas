Attribute VB_Name = "modColumnsAndRows"
Option Explicit

Sub AutofitAllUsedColumns()
    Dim x As Integer
    For x = 1 To ActiveSheet.UsedRange.Columns.Count
        Columns(x).EntireColumn.AutoFit
    Next x
End Sub

Sub CodeDivers()
    'Autofit sur colonne A & B
    Columns("A:B").EntireColumn.AutoFit
End Sub

Sub CopyPasteOneColumn()
    'Copy and Paste Column
    Range("A:A").Copy Range("B:B")
    Application.CutCopyMode = False
End Sub

Sub CutPasteOneColumn()
    'Cut and Paste Column
    Range("A:A").Cut Range("B:B")
    Application.CutCopyMode = False
End Sub

Public Sub CountMyColumns()
    'Count Selection's number of Columns
    MsgBox Selection.Columns.Count
End Sub

Public Sub CountMyRows()
    'Count Selection's number of Rows
    MsgBox Selection.Rows.Count
End Sub

Sub DeleteBlankRows()
    'Delete blank rows
    Dim x As Long
    With ActiveSheet
        For x = .Cells.SpecialCells(xlCellTypeLastCell).Row To 1 Step -1
            If WorksheetFunction.CountA(.Rows(x)) = 0 Then
                ActiveSheet.Rows(x).Delete
            End If
        Next
    End With
End Sub

Sub Delete_Rows(Data_Range As Range, Text As String)
    'Delete rows from 'Data_range' where cell in column A equals 'Text' - Need to be called
    Dim Row_Counter As Integer
    For Row_Counter = Data_Range.Rows.Count To 1 Step -1
        If Data_Range Is Nothing Then
            Exit Sub
            If UCase(Left(Data_Range.Cells(Row_Counter, 1).Value, Len(Text))) = UCase(Text) Then
                Data_Range.Cells(Row_Counter, 1).EntireRow.Delete
            End If
        Next Row_Counter
    End Sub

Public Sub StartEndMerge()
    'Return the first column AND last column of a selected Range
    Dim StartColumn As Integer
    Dim EndColumn As Integer
    'Assign Variables
    StartColumn = ActiveCell.Column
    EndColumn = Selection.Columns.Count + StartColumn - 1
    'Show Results
    MsgBox "Start Column " & StartColumn
    MsgBox "End Column " & EndColumn

End Sub

Public Sub AfterLast()
    'Do something after last row (empty cell) in column 'B'
    ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp).Offset(1, 0).Value = "FirstEmpty"
End Sub

Sub LastCol()                                    'BUG !!!
    'What is the last column ?
    Dim ColRow As Integer
    ColRow = ActiveSheet.UsedRange.Col.Count
    MsgBox ColRow

End Sub

Public Sub ActiveColumn()
    'Which column is selected ?
    MsgBox ActiveCell.Column
End Sub

Public Sub ActiveRow()
    'Which row is selected ?
    MsgBox ActiveCell.Row
End Sub

Public Sub LoopColumn()
    'Parse the entire column 'A', looking for a value 'FindMe'
    Dim c As Range
    For Each c In Range("A:A")
        If c.Value = "FindMe" Then
            MsgBox "FindMe found at " & c.Address
        End If
    Next c
End Sub

Public Sub LoopRow()
    'Parse an entire row '7', looking for a value 'FindMe'
    Dim c As Range
    For Each c In Range("7:7")
        If c.Value = "FindMe" Then
            MsgBox "FindMe found at " & c.Address
        End If
    Next c
End Sub

Public Sub ScrollToColumn()
    'Scroll to a column '5'
    ActiveWindow.ScrollColumn = 5
End Sub

Public Sub ScrollToRow()
    'Scroll to a column '10'
    ActiveWindow.ScrollRow = 10
End Sub

Sub SelectMultiColumns()
    'Select multi columns, non-Contiguous
    Range("A:A, C:C, E:E, G:G").Select
End Sub

Sub ColumnWidth()
    'Set Columns width to 30 for columns 'A' to 'E'
    Columns("A:E").ColumnWidth = 30
End Sub

Sub RowHeight()
    'Set Rows height to 30, for Rows '1' to '5'
    Rows("1:5").RowHeight = 30
End Sub

