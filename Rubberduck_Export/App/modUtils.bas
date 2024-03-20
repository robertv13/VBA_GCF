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
    lastUsedRow = wshzDocFormules.Range("E99999").End(xlUp).row 'Last used row
    If lastUsedRow > 1 Then wshzDocFormules.Range("A2:G" & lastUsedRow).ClearContents
    
    'Create an Array to receive the formulas informations
    Dim OutputArray(1499, 7) As Variant
    
    'Loop through each worksheet
    Dim ws As Worksheet
    Dim codeName As String, name As String, usedRange As String, cellsCount As String
    r = 0
    For Each ws In wb.Sheets
        If ws.codeName = "wshzDocNamedRange" Or _
            ws.codeName = "wshzDocFormules" Then
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
    With wshzDocFormules
        .Range(.Cells(2, 1), .Cells(r + 1, 8)).value = OutputArray
    End With

End Sub

Sub List_All_Subs_And_Functions() '2024-03-15 @ 21:26
    
    Dim timerStart As Double: timerStart = Timer

    Dim VBComp As Object

    Dim posSub As Integer, posFunction As Integer, posExitSub As Integer, posExitFunction As Integer, posSpace As Integer
    Dim posREM As Integer, posParam As Integer, scope As String, sType As String
    Dim lineNum As Long
    Dim trimLineOfCode As String, saveLineOfCode As String, remarks As String, params As String
    Dim arr() As Variant
    ReDim arr(1 To 300, 1 To 9)
    'Loop through all VBcomponents (modules, class and forms) in the workbook
    Dim oName As String, oType As String, r As Integer
    r = 1
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        'Check if the component is a userForm (1), a module (2) or a class module (3)
        If VBComp.Type <= 3 Then
            oName = VBComp.name
            oType = VBComp.Type
            Select Case oType
                Case 1
                    oType = "1_Module"
                Case 2
                    oType = "2_Class"
                Case 3
                    oType = "3_userform"
                Case Else
                    oType = String(10, "*")
            End Select
            'Get the code module for the component
            Dim vbCodeMod As Object
            Set vbCodeMod = VBComp.CodeModule
            'Loop through all lines in the code module
            For lineNum = 1 To vbCodeMod.CountOfLines
                'Check if the line contains 'Sub' or 'Function'
                saveLineOfCode = Trim(vbCodeMod.Lines(lineNum, 1))
                trimLineOfCode = Trim(vbCodeMod.Lines(lineNum, 1))
                posSub = InStr(trimLineOfCode, "Sub ")
                posFunction = InStr(trimLineOfCode, "Function ")
                posExitSub = InStr(trimLineOfCode, "Exit Sub")
                posExitFunction = InStr(trimLineOfCode, "Exit Function")
                If (posSub <> 0 Or posFunction <> 0) And posExitSub = 0 And posExitFunction = 0 Then
                    arr(r, 1) = oType
                    arr(r, 2) = oName
                    arr(r, 3) = lineNum
                    posREM = InStr(trimLineOfCode, ") '")
                    If posREM > 0 Then
                        remarks = Trim(Mid(trimLineOfCode, posREM + 2))
                        trimLineOfCode = Trim(Left(trimLineOfCode, posREM))
                    End If
                    posParam = InStr(trimLineOfCode, "(")
                    If posParam > 0 Then
                        params = Trim(Mid(trimLineOfCode, posParam))
                        trimLineOfCode = Trim(Left(trimLineOfCode, posParam - 1))
                    End If
                    
                    If InStr(trimLineOfCode, "Sub ") > 1 Or InStr(trimLineOfCode, "Function ") > 1 Then
                        posSpace = InStr(trimLineOfCode, " ")
'                        On Error Resume Next
                        scope = Left(trimLineOfCode, posSpace - 1)
                        trimLineOfCode = Trim(Mid(trimLineOfCode, posSpace + 1))
                    Else
                        scope = ""
                    End If
                    arr(r, 4) = scope
'                    On Error GoTo 0
'                    If Trim(scope) = "Sub" Or Trim(scope) = "Function" Then scope = ""
                    If InStr(trimLineOfCode, "Sub ") = 1 Then
                        sType = "Sub"
                        trimLineOfCode = Trim(Mid(trimLineOfCode, 5))
                    Else
                        If InStr(trimLineOfCode, "Function ") = 1 Then
                            sType = "Function"
                            trimLineOfCode = Trim(Mid(trimLineOfCode, 10))
                        End If
                    End If
                    arr(r, 5) = sType
                    arr(r, 6) = trimLineOfCode
                    If params <> "" Then arr(r, 7) = params
                    If remarks <> "" Then arr(r, 8) = "'" & remarks
                    arr(r, 9) = Format(Now(), "yyyy-mm-dd hh:mm:ss")
                    params = ""
                    remarks = ""
                    r = r + 1
                End If
            Next lineNum
        End If
    Next VBComp
    r = r - 1
    Set vbCodeMod = Nothing
    
    'Prepare the output worksheet
    Dim lastUsedRow As Long
    lastUsedRow = wshzDocSubsAndFunctions.Range("A999").End(xlUp).row 'Last Used Row
    wshzDocSubsAndFunctions.Range("A2:I" & lastUsedRow).ClearContents

    Dim numColumns As Long
    numColumns = UBound(arr, 2)
    Dim minArray() As Variant
    ReDim minArray(1 To r, 1 To numColumns)
    
    'Copy the data from arr to minArray
    Dim i As Integer, j As Integer
    For i = 1 To r
        For j = 1 To numColumns
            minArray(i, j) = arr(i, j)
        Next j
    Next i
    
    'Sort the array based on column 1 then column 1
    Call BubbleSort_2D_Array(minArray)
    
    'Transfer the array to the worksheet
    wshzDocSubsAndFunctions.Range("A2").Resize(UBound(minArray, 1), UBound(minArray, 2)).value = minArray

    Call Output_Timer_Results("List_All_Subs_And_Functions()", timerStart)

End Sub

Sub List_All_Shapes_Properties()
    Dim ws As Worksheet
    Dim shp As Shape
    
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
    For Each shp In ws.Shapes
        ws.Range("D" & r).value = shp.Type
        ws.Range("E" & r).value = shp.name
        ws.Range("F" & r).value = shp.ZOrderPosition
        ws.Range("G" & r).value = shp.Top
        ws.Range("H" & r).value = shp.Left
        ws.Range("I" & r).value = shp.width
        ws.Range("J" & r).value = shp.Height
        r = r + 1
    Next shp
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

Sub List_All_Conditional_Formatting() '2024-02-24 @ 07:36
    
    Dim timerStart As Double: timerStart = Timer

    Dim r As Integer: r = 1
    Dim numRows As Integer, numCols As Integer
    numRows = 100
    numCols = 5
    Dim arr(1 To 100, 1 To 5) As String

    'Loop through each worksheet
    Dim ws As Worksheet
    Dim fc As FormatCondition
    Dim rng As Range
    Dim wsName As String
    For Each ws In ThisWorkbook.Worksheets
        wsName = ws.name
        'Loop through each conditional formatting rule within the worksheet
        For Each fc In ws.Cells.FormatConditions
            'Debug.Print Tab(5); "Type: " & TypeName(fc)
            arr(r, 1) = wsName
            arr(r, 2) = fc.AppliesTo.Address
            Set rng = fc.AppliesTo
            arr(r, 3) = rng.Cells.count
            arr(r, 4) = fc.Formula1
            arr(r, 5) = Now()
            r = r + 1
        Next fc
    Next ws
    
    Set fc = Nothing
    Set ws = Nothing
    Set rng = Nothing
    
    'Setup and prepare the output worksheet
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Sheets("DocConditionalFormatting")
    Dim lastUsedRow As Long
    lastUsedRow = wsOutput.Range("A999").End(xlUp).row 'Last Used Row
    wsOutput.Range("A2:F" & lastUsedRow).ClearContents
    
    wsOutput.Range("A2").Resize(numRows, numCols).value = arr
    
    Set wsOutput = Nothing
    
    Call Output_Timer_Results("List_All_Conditional_Formatting()", timerStart)
 
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

Sub BubbleSort_2D_Array(ByRef arr() As Variant) 'ChatGPT - 2024-02-26 @ 11:40
    
    Dim i As Long, j As Long
    Dim numRows As Long, numCols As Long
    Dim temp As Variant
    Dim sorted As Boolean

    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)

    'Bubble Sort Algorithm
    Dim c As Integer
    For i = 1 To numRows - 1
        sorted = True
        For j = 1 To numRows - i
            'Compare column 2 first
            If arr(j, 1) > arr(j + 1, 1) Then
                'Swap rows
                For c = 1 To numCols
                    temp = arr(j, c)
                    arr(j, c) = arr(j + 1, c)
                    arr(j + 1, c) = temp
                Next c
                sorted = False
            ElseIf arr(j, 1) = arr(j + 1, 1) Then
                'Column 1 values are equal, then compare column2 values
                If arr(j, 2) > arr(j + 1, 2) Then
                    'Swap rows
                    For c = 1 To numCols
                        temp = arr(j, c)
                        arr(j, c) = arr(j + 1, c)
                        arr(j + 1, c) = temp
                    Next c
                    sorted = False
                End If
            End If
        Next j
        'If no swaps were made, the array is sorted
        If sorted Then Exit For
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

Sub Build_Worksheet_Columns() '2024-02-28 @ 06:40

    Dim arr(1 To 20, 1 To 2) As Variant
    Dim output(1 To 200, 1 To 5) As Variant
    Dim r As Long
    r = 0
    r = r + 1: arr(r, 1) = "AR_Entête": arr(r, 2) = "A2:J2"
    r = r + 1: arr(r, 1) = "ClientsImportés": arr(r, 2) = "A1:J1"
    r = r + 1: arr(r, 1) = "Doc_ConditionalFormatting": arr(r, 2) = "A1:E1"
    r = r + 1: arr(r, 1) = "Doc_Formules": arr(r, 2) = "A1:H1"
    r = r + 1: arr(r, 1) = "Doc_Log_Appli": arr(r, 2) = "A1:C1"
    r = r + 1: arr(r, 1) = "Doc_Named_Ranges": arr(r, 2) = "A1:B1"
    r = r + 1: arr(r, 1) = "Doc_Subs_&_Functions": arr(r, 2) = "A1:G1"
    r = r + 1: arr(r, 1) = "Documentation": arr(r, 2) = "A1:E1"
    r = r + 1: arr(r, 1) = "Encaissements_Entête": arr(r, 2) = "A3:F3"
    r = r + 1: arr(r, 1) = "Encaissements_Détail": arr(r, 2) = "A3:F3"
    r = r + 1: arr(r, 1) = "Factures": arr(r, 2) = "A3:T3"
    r = r + 1: arr(r, 1) = "FacturesLignes": arr(r, 2) = "A3:G3"
    r = r + 1: arr(r, 1) = "GL_Trans": arr(r, 2) = "A1:J1"
    r = r + 1: arr(r, 1) = "EJ_Auto": arr(r, 2) = "C1:J1"
    r = r + 1: arr(r, 1) = "Invoice List": arr(r, 2) = "A2:J2"
    r = r + 1: arr(r, 1) = "TEC_Local": arr(r, 2) = "A2:P2"
    r = 1
    Dim i As Long, colNo As Integer
    For i = 1 To UBound(arr, 1)
        If arr(i, 1) = "" Then Exit For
        Dim rng As Range, rngAddress As String, cell As Range
        Set rng = Sheets(arr(i, 1)).Range(arr(i, 2))
        colNo = 0
        For Each cell In rng
            colNo = colNo + 1
            output(r, 2) = arr(i, 1)
            output(r, 3) = Chr(64 + colNo)
            output(r, 4) = colNo
            output(r, 5) = cell.value
            r = r + 1
        Next cell
    Next i
    
    'Setup and prepare the output worksheet
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Sheets("Doc_TableLayouts")
    Dim lastUsedRow As Long
    lastUsedRow = wsOutput.Range("A999").End(xlUp).row 'Last Used Row
    wsOutput.Range("A2:F" & lastUsedRow + 1).ClearContents
    
    wsOutput.Range("A2").Resize(r, 5).value = output
    
    Set rng = Nothing
    Set cell = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub Output_Timer_Results(subName As String, t As Double)

    Dim modeOper As Integer
    modeOper = 2 '2024-02-26 @ 16:42
    
    'modeOper = 1 - Dump to immediate Window
    If modeOper = 1 Then
        Dim l As Integer: l = Len(subName)
        Debug.Print vbNewLine & String(40 + l, "*") & vbNewLine & _
        Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & subName & " = " _
        & Format(Timer - t, "##0.0000") & " secondes" & vbNewLine & String(40 + l, "*")
    End If

    'modeOper = 2 - Dump to worksheet
    If modeOper = 2 Then
        With wshzDocLogAppli
            Dim lastUsedRow As Long
            lastUsedRow = .Range("A9999").End(xlUp).row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("B" & lastUsedRow).value = subName
            If t Then
                .Range("C" & lastUsedRow).value = Timer - t
            End If
        End With
    End If

End Sub

Sub Reorganize_Tests_And_Todos_Worksheet() '2024-03-02 @ 15:21

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshzDocTests_And_Todos
    Dim rng As Range, lastUsedRow As Long
    lastUsedRow = ws.Range("A999").End(xlUp).row
    Set rng = ws.Range("A1:E" & lastUsedRow)
    
    With ws.ListObjects("tblTests_And_Todo").Sort
        .SortFields.clear
        .SortFields.Add2 _
            Key:=Range("tblTests_And_Todo[Statut]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            Key:=Range("tblTests_And_Todo[Module]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            Key:=Range("tblTests_And_Todo[Priorité]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            Key:=Range("tblTests_And_Todo[TimeStamp]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim tbl As ListObject: Set tbl = ws.ListObjects("tblTests_And_Todo")
    Dim rowToMove As Range

    'Move completed item ($D = a) to the bottom of the list
    Dim i As Long, lastRow As Long
    i = 2

    While ws.Range("D2").value = "a"
        Set rowToMove = tbl.ListRows(1).Range
        lastRow = tbl.ListRows.count
        rowToMove.Cut Destination:=tbl.DataBodyRange.Rows(lastRow + 1)
        tbl.ListRows(1).delete
    Wend

    ws.Calculate
    
    Set ws = Nothing
    Set rng = Nothing
    Set tbl = Nothing
    Set rowToMove = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Sub Test_Lookup_Data_In_A_Range()

    Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
    Dim myInfo() As Variant
    
    Dim searchString As String
    searchString = "Gestion MAROB inc."
    
    myInfo = Lookup_Data_In_A_Range(rng, 1, searchString, 3)
    
    Set rng = Nothing

End Sub

Sub Test_Array_To_Range() '2024-03-18 @ 17:34

    Dim ws As Worksheet
    Set ws = Feuil2
    
    Dim arr() As Variant
    ReDim arr(1 To 1000, 1 To 20)
    
    Dim i As Integer, j As Integer
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            arr(i, j) = "i = " & i & " and j = " & j & " - *********"
        Next j
    Next i
    
    ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
    
End Sub

