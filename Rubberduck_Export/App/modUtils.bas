Attribute VB_Name = "modUtils"
Option Explicit

Sub ListWorksheets()
    
    Dim ws As Worksheet
    
    'Loop through all worksheets in the active workbook
    For Each ws In ThisWorkbook.Sheets
        Dim spaces As String
        spaces = Space(22 - Len(ws.codeName))
        'Print the name of each worksheet to the Immediate Window
        Debug.Print "codeName = " & ws.codeName & spaces & "Name = " & ws.Name; ""
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
        ThisWorkbook.Sheets("Utils").Cells(i, 2).value = ThisWorkbook.Sheets(wshNames(i)).codeName
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

Sub ListAllProceduresAndAllFunctions()
    Dim VBComp As Object
    Dim vbCodeMod As Object
    Dim LineNum As Long
    Dim ProcName As String
    
    'Loop through all components (modules, forms, etc.) in the workbook
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        'Debug.Print vbComp.Name
        'Check if the component is a code module
        If VBComp.Type = 1 Then
            ' Get the code module for the component
            Set vbCodeMod = VBComp.CodeModule
            
            'Loop through all lines in the code module
            For LineNum = 1 To vbCodeMod.CountOfLines
                'Check if the line contains a procedure or function
                If Left(Trim(vbCodeMod.Lines(LineNum, 1)), 1) = "Sub" Or Left(Trim(vbCodeMod.Lines(LineNum, 1)), 3) = "Function" Then
                    'Extract the procedure or function name
                    ProcName = Mid(Trim(vbCodeMod.Lines(LineNum, 1)), 4)
                    
                    'Print the name to the Immediate Window
                    Debug.Print VBComp.Name & ": " & Trim(ProcName)
                End If
            Next LineNum
        End If
    Next VBComp
End Sub




