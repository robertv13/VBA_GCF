Attribute VB_Name = "modUtils"
Option Explicit

Sub ListWorksheets()
    
    Dim ws As Worksheet
    
    'Loop through all worksheets in the active workbook
    For Each ws In ThisWorkbook.Sheets
        'Print the name of each worksheet to the Immediate Window
        Debug.Print "Name = " & ws.Name & " - CName = " & ws.CodeName
    Next ws
    
End Sub

Sub ListSortedWorksheetsToUtils() '2023-12-02 @ 14:40
    
    Dim ws As Worksheet
    Dim wshNames() As String
    Dim i As Integer
    
    'Clear the worksheets 'Utils' -OR- create it"
    On Error Resume Next
    ThisWorkbook.Sheets("Utils").Delete
    On Error GoTo 0
    
    'Add the new worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.count))
    ws.Name = "Utils"
    
    'Resize the array to the number of worksheets in the active workbook
    ReDim wshNames(1 To ThisWorkbook.Sheets.count)
    
    'Store the names of all worksheets in the array
    For Each ws In ThisWorkbook.Sheets
        i = i + 1
        wshNames(i) = ws.Name
        Debug.Print wshNames(i)
    Next ws
    
    'Sort the array alphabetically
    SortArrayAlphabetically wshNames
    
    'Print or use the sorted list in the "Utils" worksheet
    ThisWorkbook.Sheets("Utils").Cells(1, 1).value = "name"
    ThisWorkbook.Sheets("Utils").Cells(1, 2).value = "codeName"
    
    For i = LBound(wshNames) + 1 To UBound(wshNames)
        'Output to the "Utils" worksheet starting from cell A1
        ThisWorkbook.Sheets("Utils").Cells(i, 1).value = wshNames(i)
        ThisWorkbook.Sheets("Utils").Cells(i, 2).value = ThisWorkbook.Sheets(wshNames(i)).CodeName
    Next i
    
    Sheets("Utils").Columns("A:B").AutoFit
    
End Sub

Sub SortArrayAlphabetically(arr() As String) '2023-12-02 @ 14:40
    Dim i As Long, j As Long
    Dim temp As String
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                'Swap elements if they are in the wrong order
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Sub GetAllRecordsFromAClosedWorkbook()
    Dim sourceWorkbook As Workbook
    Dim outputSheet As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim targetRow As Long
    Dim targetCol As Long

    'Set the source workbook (change the path as needed)
    Set sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                        "GCF_BD_Entrée.xlsx" '2023-12-19 @ 07:21

    'Set the output sheet
    Set outputSheet = ThisWorkbook.Sheets.Add
    outputSheet.Name = "Records"

    'Initialize target row
    targetRow = 1

    'Loop through all worksheets in the source workbook
    For Each ws In sourceWorkbook.Sheets
        'Find the last row and last column in the current worksheet
        lastRow = WorksheetFunction.Min(ws.Cells(ws.Rows.count, "A").End(xlUp).row, 25)
        lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

        'Copy data to the output sheet
'        ws.Range(ws.Cells(1, 1)).value = ws.Name
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy outputSheet.Cells(targetRow, 1)

        'Update target row for the next worksheet
        targetRow = targetRow + lastRow + 1
    Next ws

    'Close the source workbook without saving changes
    sourceWorkbook.Close SaveChanges:=False
    
End Sub





