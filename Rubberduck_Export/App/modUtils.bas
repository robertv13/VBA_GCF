Attribute VB_Name = "modUtils"
Option Explicit

Sub List_All_Worksheets()
    
    Dim ws As Worksheet
    
    'Loop through all worksheets in the active workbook
    For Each ws In ThisWorkbook.Sheets
        Dim spaces As String
        spaces = Space(25 - Len(ws.codeName))
        'Print the name of each worksheet to the Immediate Window
        Debug.Print "codeName = " & ws.codeName & spaces & "Name = " & ws.name; ""
    Next ws
    
End Sub

Sub List_All_Formulas() '2024-02-19 @ 07:41 - ChatGPT
    'Set a reference to the current workbook
    Dim wb As Workbook: Set wb = ThisWorkbook
    
    'Prepare existing worksheet to receive data
    Dim lastUsedRow As Long, r As Long, c As Long
    lastUsedRow = wshDocFormules.Range("E99999").End(xlUp).row 'Last used row
    If lastUsedRow > 1 Then wshDocFormules.Range("A2:G" & lastUsedRow).ClearContents
    
    'Create an Array to receive the formulas informations
    Dim OutputArray(1499, 7) As Variant
    
    'Loop through each worksheet
    Dim ws As Worksheet
    Dim codeName As String, name As String, usedRange As String, cellsCount As String
    r = 0
    For Each ws In wb.Sheets
        If ws.codeName = "wshDocNamedRange" Or _
            ws.codeName = "wshDocFormules" Then
                GoTo Continue_for_each_ws
        End If
        Debug.Print r; ws.name; Tab(20); ws.codeName; Tab(45); Now()
        'Save information for this worksheet
        codeName = ws.codeName
        name = ws.name
        usedRange = ws.usedRange.Address
        cellsCount = ws.usedRange.count
        'Loop through all cells in the used range
        Dim cell As Range
        For Each cell In ws.usedRange
            'Does the cell contain a Formula
            If Left(cell.formula, 1) = "=" Then
                'Write formula information to the destination worksheet
                OutputArray(r, 0) = codeName
                OutputArray(r, 1) = name
                OutputArray(r, 2) = usedRange
                OutputArray(r, 3) = cellsCount
                OutputArray(r, 4) = cell.Address
                OutputArray(r, 5) = "'=" & Mid(cell.formula, 2) 'Add ' to preserve formulas
                OutputArray(r, 6) = Now() 'Timestamp
                OutputArray(r, 7) = "=ROW()"
                r = r + 1 'Move to the next row
            End If
        Next cell
Continue_for_each_ws:
    Next ws
    'Transfer the array data to the worksheet
    With wshDocFormules
        .Range(.Cells(2, 1), .Cells(r + 1, 8)).value = OutputArray
    End With

End Sub

Sub List_All_Subs_And_Functions() '2024-02-17 @ 16:53
    
    Dim VBComp As Object
    
    'Setup the worksheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("DocSubs&Functions")

    'Prepare the output worksheet
    Dim r As Long
    r = ws.Range("D9999").End(xlUp).row 'Last Used Row
    ws.Range("A2:F" & r).ClearContents
    r = 2
    
    Dim posSub As Integer, posFunction As Integer, posExitSub As Integer, posREM As Integer
    Dim lineNum As Long
    Dim trimLineOfCode As String, remarks As String
    'Loop through all VBcomponents (modules, class and forms) in the workbook
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        'Check if the component is a userForm (1), a module (2) or a class module (3)
        If VBComp.Type <= 3 Then
            ws.Cells(r, 1).value = VBComp.name
            ws.Cells(r, 2).value = VBComp.Type
            'Get the code module for the component
            Dim vbCodeMod As Object
            Set vbCodeMod = VBComp.CodeModule
            Debug.Print vbCodeMod.name
            'Loop through all lines in the code module
            For lineNum = 1 To vbCodeMod.CountOfLines
                'Check if the line contains 'Sub' or 'Function'
                trimLineOfCode = Trim(vbCodeMod.Lines(lineNum, 1))
                posSub = InStr(trimLineOfCode, "Sub ")
                posFunction = InStr(trimLineOfCode, "Function ")
                posExitSub = InStr(trimLineOfCode, "Exit Sub")
                If (posSub <> 0 Or posFunction <> 0) And posExitSub = 0 Then
                    ws.Cells(r, 3).value = lineNum
                    posREM = InStr(trimLineOfCode, ") '")
                    If posREM > 0 Then
                        remarks = Trim(Mid(trimLineOfCode, posREM + 1))
                        trimLineOfCode = Trim(Left(trimLineOfCode, posREM))
                    End If
                    ws.Cells(r, 4).value = trimLineOfCode
                    If remarks <> "" Then ws.Cells(r, 5).value = "'" & remarks
                    ws.Cells(r, 6).value = Now()
                    r = r + 1
                    remarks = ""
                End If
            Next lineNum
        End If
        r = r + 1
    Next VBComp
End Sub

Sub List_All_Shapes_Properties()
    Dim ws As Worksheet
    Dim Shp As Shape
    
    ' Set the worksheet (change "Sheet1" to your sheet's name)
    Set ws = ActiveSheet
    
    Dim r As Integer
    r = 2
    ws.Range("D" & r).value = "Type"
    ws.Range("E" & r).value = "Shape Name"
    ws.Range("F" & r).value = "ZOrder"
    ws.Range("G" & r).value = "Top"
    ws.Range("H" & r).value = "Left"
    ws.Range("I" & r).value = "Width"
    ws.Range("J" & r).value = "Height"
    
    r = 3
    'Loop through all shapes on the worksheet
    For Each Shp In ws.Shapes
        ws.Range("D" & r).value = Shp.Type
        ws.Range("E" & r).value = Shp.name
        ws.Range("F" & r).value = Shp.ZOrderPosition
        ws.Range("G" & r).value = Shp.Top
        ws.Range("H" & r).value = Shp.Left
        ws.Range("I" & r).value = Shp.Width
        ws.Range("J" & r).value = Shp.Height
        r = r + 1
    Next Shp
End Sub

Sub List_All_Named_Ranges() '2024-02-18 @ 07:23 - From ChatGPT
    
    Application.ScreenUpdating = False
    
    Dim nr As name
    
    'Define the old workbook name and the new workbook name
    Dim wbName As String
    wbName = ThisWorkbook.name

    'Setup and prepare the output worksheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("DocNamedRanges")
    Dim r As Long
    r = ws.Range("A9999").End(xlUp).row 'Last Used Row
    ws.Range("A2:F" & r).ClearContents
    r = 2
    
    'Loop through each named range in the workbook
    For Each nr In ThisWorkbook.Names
        Debug.Print nr.name
        ws.Cells(r, 1).value = nr.name
        ws.Cells(r, 2).value = "'" & nr.value
        ws.Cells(r, 3).value = "'" & nr.Parent
        ws.Cells(r, 4).value = Now()
        r = r + 1
'        'Check if the named range refers to the old workbook
'        If InStr(1, nr.RefersTo, oldWorkbookName) > 0 Then
'            ' Replace the old workbook name with the new workbook name
'            nr.RefersTo = Replace(nr.RefersTo, oldWorkbookName, newWorkbookName)
'        End If
    Next nr
    
    Application.ScreenUpdating = True

End Sub

Sub Array_Bubble_Sort(arr() As String)
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

Sub Protect_Unprotect_Worksheet()
    Dim password As String
    password = "GCmfp"

    'Unprotect the worksheet with the password
    ActiveSheet.Unprotect password:=password

    'Your code to modify cells goes here

    'Protect the worksheet again with the password
    ActiveSheet.Protect password:=password
End Sub

Sub Add_Columns_To_Active_Worksheet()
    Dim colToAdd As Integer
    colToAdd = 5
    
    'Set the worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Find the last column with data
    Dim lastColumn As Integer
    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    'Add columns to the right of the last column
    ws.Columns(lastColumn + 1).Resize(, colToAdd).Insert Shift:=xlToRight
    
    Debug.Print colToAdd & " columns added to the worksheet."
End Sub

