Attribute VB_Name = "modDev_Utils"
Option Explicit

'@TODO - Enlever la procédure ci-dessous - 2025-07-07 @ 15:49
'Sub Add_Columns_To_Active_Worksheet()
'
'    Dim colToAdd As Long
'    colToAdd = 5
'
'    'Set the worksheet
'    Dim ws As Worksheet: Set ws = ActiveSheet
'
'    'Find the last column with data
'    Dim lastColumn As Long
'    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
'
'    'Add columns to the right of the last column
'    ws.Columns(lastColumn + 1).Resize(, colToAdd).Insert Shift:=xlToRight
'
'    Debug.Print "#026 - " & colToAdd & " columns added to the worksheet."
'
'    'Libérer la mémoire
'    Set ws = Nothing
'
'End Sub
'
Sub TrierTableau2DBubble(ByRef arr() As Variant) '2024-06-23 @ 07:05
    
    Dim i As Long, j As Long, numRows As Long, numCols As Long
    Dim temp As Variant
    Dim sorted As Boolean
    
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    'Bubble Sort Algorithm
    Dim c As Long, cProcess As Long
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

Sub RedimensionnerTableau2D(ByRef inputArray As Variant, ByVal nRows As Long, ByVal nCols As Long)
    
    Dim oRows As Long, oCols As Long
    
    'Get the original dimensions of the input array
    oRows = UBound(inputArray, 1)
    oCols = UBound(inputArray, 2)
    
    'Ensure the new dimensions are within the original array's bounds
    If nRows > oRows Then nRows = oRows
    If nCols > oCols Then nCols = oCols
    
    'Create a new array with the specified dimensions
    Dim tempArray() As Variant
    ReDim tempArray(LBound(inputArray, 1) To nRows, LBound(inputArray, 2) To nCols)
    
    ' Copy the relevant data from the input array to the new array
    Dim i As Long, j As Long
    For i = LBound(inputArray, 1) To nRows
        For j = LBound(inputArray, 2) To nCols
            tempArray(i, j) = inputArray(i, j)
        Next j
    Next i
    
    ' Assign the trimmed array back to the input array
    inputArray = tempArray
    
End Sub

Sub TrierTableauBubble(MyArray() As String) '2024-07-02 @ 15:18
    
    Dim i As Long, j As Long
    Dim temp As Variant
    For i = LBound(MyArray) To UBound(MyArray) - 1
        For j = i + 1 To UBound(MyArray)
            If MyArray(i) > MyArray(j) Then
                temp = MyArray(j)
                MyArray(j) = MyArray(i)
                MyArray(i) = temp
            End If
        Next j
    Next i
    
End Sub

'CommentOut - Pas utilisé -2025-07-14 @ 09:48
'Sub Check_Invoice_Template()
'
'    Dim ws As Worksheet: Set ws = wsdADMIN
'    Dim firstUsedRow As Long, lastUsedRow As Long
'    firstUsedRow = 12
'    lastUsedRow = ws.Cells(ws.Rows.count, "Z").End(xlUp).Row
'    Dim rng As Range
'    Set rng = ws.Range("Z" & firstUsedRow & ":AA" & lastUsedRow)
'
'    'First - Determine which templates are used
'    Dim arr As Variant
'    Dim strTemplates As String
'    Dim i As Long, j As Long
'    For i = 1 To lastUsedRow - firstUsedRow + 1
'        If Not rng.Cells(i, 2) = vbNullString Then
'            arr = Split(rng.Cells(i, 2), ",")
'            For j = 0 To UBound(arr)
'                strTemplates = strTemplates & Trim$(arr(j)) & "-" & i & "|"
'            Next j
'        End If
'    Next i
'
'    'Second - Sort all the found templates
'    Dim tt() As String
'    tt = Split(strTemplates, "|")
'    Call TrierTableauBubble(tt)
'
'    'Third - Prepare the worksheet to receive information
'    Call EffacerEtRecreerWorksheet("Gabarits_Facture")
'
'    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Gabarits_Facture")
'    wsOutput.Range("A1").Value = "Gabarit"
'    wsOutput.Range("B1").Value = "Code"
'    wsOutput.Range("C1").Value = "Service"
'    Dim outputRow As Long: outputRow = 1
'
'    'Third - Build the list of services associated to each template (First Letter)
'    Dim rowNo As Long
'    Dim template As String, oldTemplate As String
'
'    With wsOutput
'        For i = 0 To UBound(tt)
'            If tt(i) <> vbNullString Then
'                template = Left$(tt(i), 1)
'                If template <> oldTemplate Then
'                    outputRow = outputRow + 2
'                    .Range("A" & outputRow).Value = "Gabarit '" & template & "'"
'                    oldTemplate = template
'                End If
'                rowNo = Mid$(tt(i), InStr(1, tt(i), "-") + 1)
'                outputRow = outputRow + 1
'                .Range("B" & outputRow).Value = tt(i)
'                .Range("C" & outputRow).Value = rng.Cells(rowNo, 1)
'            End If
'        Next i
'        wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit
'    End With
'
'    With wsOutput.Range("A1:C1")
'        With .Interior
'            .Pattern = xlSolid
'            .PatternColorIndex = xlAutomatic
'            .Color = 12611584
'            .TintAndShade = 0
'            .PatternTintAndShade = 0
'        End With
'
'        With .Font
'            .ThemeColor = xlThemeColorDark1
'            .TintAndShade = 0
'            .size = 10
'            .Italic = True
'            .Bold = True
'        End With
'    End With
'
'    With wsOutput.Range("A2:A" & outputRow)
'        .Font.Bold = True
'    End With
'
'    'Cleaning - 2024-07-02 @ 20:12
'    Set rng = Nothing
'    Set ws = Nothing
'    Set wsOutput = Nothing
'
'End Sub
'
Sub List_Worksheets_From_Closed_Workbook_All() '2024-07-14 @ 07:02
    
    Call EffacerEtRecreerWorksheet("X_Feuilles_du_Classeur")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Feuilles_du_Classeur")
    wsOutput.Range("A1").Value = "Feuille"
    wsOutput.Range("B1").Value = "CodeName"
    wsOutput.Range("C1").Value = "TimeStamp"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"), RGB(0, 112, 192))

    'Specify the full path and name of the closed workbook
    Dim wbPath As String
    wbPath = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER.xlsx"
    
    'Open the workbook in read-only mode
    Dim wb As Workbook: Set wb = Workbooks.Open(fileName:=wbPath, ReadOnly:=True)
    
    Dim wbName As String
    wbName = wb.Name
    
    'Loop through each worksheet in the workbook and add its name to the immediate window
    Dim arr() As Variant
    ReDim arr(1 To 100, 1 To 3)
    Dim ws As Worksheet
    Dim timeStamp As String
    Dim r As Long
    Dim f As Long
    For Each ws In wb.Worksheets
        r = r + 1
        arr(r, 1) = ws.Name
        arr(r, 2) = ws.CodeName
        timeStamp = Format$(Now(), "dd-mm-yyyy hh:mm:ss")
        arr(r, 3) = timeStamp
    Next ws
    
    'Close the workbook without saving changes
    wb.Close SaveChanges:=False
    
    Call RedimensionnerTableau2D(arr, r, UBound(arr, 2))
    
    Call TrierTableau2DBubble(arr)
    
    For r = 1 To UBound(arr, 1)
        wsOutput.Cells(r + 1, 1) = arr(r, 1)
        wsOutput.Cells(r + 1, 2) = arr(r, 2)
        wsOutput.Cells(r + 1, 3) = arr(r, 3)
        f = f + 1
    Next r
    
    wsOutput.Columns.AutoFit
    
   'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = r + 2
    wsOutput.Range("A" & lastUsedRow).Value = "*** " & Format$(f, "###,##0") & _
                                    " feuilles pour le workbook '" & wbName & "' ***"
    
    lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "A").End(xlUp).Row
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Liste des feuilles d'un classeur"
    Dim header2 As String: header2 = wbName
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    ThisWorkbook.Worksheets("X_Feuilles_du_Classeur").Activate
    
    'Libérer la mémoire
    Set rngToPrint = Nothing
    Set wb = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
End Sub

'@Description ("Saisie des chaines, construction du tableau des lignes de code")
Sub RechercherCodeProjet() '2024-10-26 @ 10:41
    
    'Declare lineOfCode() as variant
    Dim allLinesOfCode As Variant
    ReDim allLinesOfCode(1 To 45000, 1 To 4)
    
    'Allows up to 3 search strings
    Dim search1 As String, search2 As String, search3 As String
    search1 = InputBox("Enter the search string ? ", "Search1")
    search2 = InputBox("Enter the search string ? ", "Search2")
    search3 = InputBox("Enter the search string ? ", "Search3")
    
    'Loop through all VBcomponents (modules, class and forms) in the active workbook
    Dim LineNum As Long
    Dim lignesLues As Long
    Dim indiceTableau As Long
    
    Dim vbComp As Object
    Dim oType As String
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
        Case 1
            oType = "1_Module"
        Case 2
            oType = "2_Class"
        Case 3
            oType = "3_userform"
        Case 100
            oType = "0_Worksheet"
        Case Else
            oType = "9_?????"
            Stop
        End Select
        
        'Get the code module for the component
        Dim vbCodeMod As Object: Set vbCodeMod = vbComp.codeModule
        
        'Loop through all lines in the code module to save all the lines in memory
        For LineNum = 1 To vbCodeMod.CountOfLines
            lignesLues = lignesLues + 1
            If Trim$(vbCodeMod.Lines(LineNum, 1)) <> vbNullString Then
                indiceTableau = indiceTableau + 1
                allLinesOfCode(indiceTableau, 1) = oType
                allLinesOfCode(indiceTableau, 2) = vbComp.Name
                allLinesOfCode(indiceTableau, 3) = LineNum
                allLinesOfCode(indiceTableau, 4) = Trim$(vbCodeMod.Lines(LineNum, 1))
            End If
        Next LineNum
    Next vbComp
    
    'At this point allLinesOfCode contains all non-empty lines of code of the application - 2025-06-18 @ 14:48
    
    Call RedimensionnerTableau2D(allLinesOfCode, indiceTableau, UBound(allLinesOfCode, 2))
    
    Call Search_Every_Lines_Of_Code(allLinesOfCode, lignesLues, search1, search2, search3)
    
    'Libérer la mémoire
    Set vbComp = Nothing
    Set vbCodeMod = Nothing
    
End Sub

Sub List_Conditional_Formatting_All() '2024-06-23 @ 18:37

    'Work in memory
    Dim arr() As Variant
    ReDim arr(1 To 100, 1 To 7)
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim area As Range
    Dim ruleIndex As Long
    Dim cf As FormatCondition
    Dim i As Long
    
    'Loop through each worksheet in the current workbook
    For Each ws In ThisWorkbook.Worksheets
        
        On Error Resume Next
        'Attempt to get the range with conditional formatting
        Application.EnableEvents = False
        Set rng = ws.usedRange.SpecialCells(xlCellTypeAllFormatConditions)
        Application.EnableEvents = True
        On Error GoTo 0
        
        'Check if rng is not nothing, which means there are conditional formatting rules
        If Not rng Is Nothing Then
            'Loop through each area in the range
            For Each area In rng.Areas
                Debug.Print "#027 - " & ws.Name & " - " & area.FormatConditions.count
                ' Loop through each conditional formatting rule in the area
                For ruleIndex = 1 To area.FormatConditions.count
                    Set cf = area.FormatConditions(ruleIndex)
                    i = i + 1
                    arr(i, 1) = ws.Name & Chr$(0) & area.Address
                    arr(i, 2) = ws.Name
                    arr(i, 3) = area.Address
                    arr(i, 4) = cf.Type
                    arr(i, 5) = cf.Formula1
                    
                    On Error Resume Next
                    If cf.Type = xlCellValue And (cf.Operator = xlBetween Or cf.Operator = xlNotBetween) Then
                        arr(i, 6) = cf.Formula2
                    End If
                    On Error GoTo 0
                    
                    arr(i, 7) = Format$(Now(), "yyyy-mm-dd hh:mm")
                Next ruleIndex
            Next area
        End If
        
        'Reset the range variable for the next worksheet
    Next ws
    
    Call RedimensionnerTableau2D(arr, i, UBound(arr, 2))
    Call TrierTableau2DBubble(arr)

    'Setup and prepare the output worksheet
    Dim wsOutput As Worksheet: Set wsOutput = wshzDocConditionalFormatting
    Dim lastUsedRow As Long
    lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "A").End(xlUp).Row
    If lastUsedRow > 1 Then
        wsOutput.Range("A2:F" & lastUsedRow).ClearContents
    End If
    
    'Assign array to range
    wsOutput.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    wsOutput.Range("A:A").EntireColumn.Hidden = True 'Do not show the SortKey
   
    MsgBox "J'ai trouvé " & i & " Conditional Formatting"
    
    'Libérer la mémoire
    Set area = Nothing
    Set cf = Nothing
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing

End Sub

Sub List_Data_Validations_All() '2024-07-15 @ 06:52

    'Prepare the result worksheet (wsOutput)
    Call EffacerEtRecreerWorksheet("Doc_Data_Validations")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Doc_Data_Validations")
    wsOutput.Cells(1, 1).Value = "SortKey"
    wsOutput.Cells(1, 2).Value = "Worksheet"
    wsOutput.Cells(1, 3).Value = "CellAddress"
    wsOutput.Cells(1, 4).Value = "ValidationType"
    wsOutput.Cells(1, 5).Value = "Formula1"
    wsOutput.Cells(1, 6).Value = "Formula2"
    wsOutput.Cells(1, 7).Value = "Operator"
    wsOutput.Cells(1, 8).Value = "TimeStamp"
    
    Call Make_It_As_Header(wsOutput.Range("A1:H1"), RGB(0, 112, 192))
    
    'Create the Array to store results in memory
    Dim arr() As Variant
    ReDim arr(1 To 5000, 1 To 8)
    
    ' Loop through each worksheet in the workbook
    Dim dvType As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim timeStamp As String
    Dim X As Long: X = 1
    Dim xAnalyzed As Long
    For Each ws In ThisWorkbook.Worksheets
        'Loop through each cell in the worksheet
        For Each cell In ws.usedRange
            'Check if the cell has data validation
            xAnalyzed = xAnalyzed + 1

            On Error Resume Next
            dvType = vbNullString
            dvType = cell.Validation.Type
            On Error GoTo 0
            
            If dvType <> vbNullString And dvType <> "0" Then
                'Write the data validation details to the output sheet
                arr(X, 1) = ws.Name & Chr$(0) & cell.Address 'Sort Key
                arr(X, 2) = ws.Name
                arr(X, 3) = cell.Address
                arr(X, 4) = dvType
                Select Case dvType
                    Case "2"
                        arr(X, 4) = "Min/Max"
                    Case "3"
                        arr(X, 4) = "Liste"
                    Case Else
                        arr(X, 4) = dvType
                End Select
                On Error Resume Next
                arr(X, 5) = "'" & cell.Validation.Formula1
                On Error GoTo 0
                
                On Error Resume Next
                arr(X, 6) = "'" & cell.Validation.Formula2
                On Error GoTo 0
                
                On Error Resume Next
                arr(X, 7) = "'" & cell.Validation.Operator
                On Error GoTo 0
                
                timeStamp = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
                arr(X, 8) = timeStamp

                'Increment the output row counter
                X = X + 1
            End If
        Next cell
    Next ws

    If X > 1 Then
    
        X = X - 1
        
        Call RedimensionnerTableau2D(arr, X, UBound(arr, 2))
        
        Call TrierTableau2DBubble(arr)
        
        'Array to Worksheet
        Dim outputRow As Long: outputRow = 2
        wsOutput.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
        wsOutput.Range("A:A").EntireColumn.Hidden = True 'Do not show the sortKey
        wsOutput.Columns(4).HorizontalAlignment = xlCenter
        wsOutput.Columns(7).HorizontalAlignment = xlCenter
        wsOutput.Columns(8).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        Dim lastUsedRow As Long
        lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "B").End(xlUp).Row
        Dim j As Long, oldWorksheet As String
        oldWorksheet = wsOutput.Range("B" & lastUsedRow).Value
        For j = lastUsedRow To 2 Step -1
            If wsOutput.Range("B" & j).Value <> oldWorksheet Then
                wsOutput.Rows(j + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                oldWorksheet = wsOutput.Range("B" & j).Value
            End If
        Next j
        
        'Since we might have inserted new row, let's update the lastUsedRow
        lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "B").End(xlUp).Row
        With wsOutput.Range("B2:H" & lastUsedRow)
            On Error Resume Next
            ActiveSheet.Cells.FormatConditions.Delete
            On Error GoTo 0
        
            .FormatConditions.Add Type:=xlExpression, Formula1:= _
                "=ET($B2<>"""";MOD(LIGNE();2)=1)"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        
        wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit

    End If

    'AutoFit the columns for better readability
    wsOutput.Columns.AutoFit
    
    'Result print setup - 2024-07-15 @ 09:22
    lastUsedRow = lastUsedRow + 2
    wsOutput.Range("B" & lastUsedRow).Value = "*** " & Format$(xAnalyzed, "###,##0") & _
                                    " cellules analysées dans l'application ***"
    Dim header1 As String: header1 = "Cells Data Validations"
    Dim header2 As String: header2 = "All worksheets"
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsOutput, wsOutput.Range("B2:H" & lastUsedRow), _
                           header1, _
                           header2, _
                           "$1:$1", _
                           "L")
    
    'Libérer la mémoire
    Set cell = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    MsgBox "Data validation list were created in worksheet: " & wsOutput.Name
    
End Sub

Sub EffacerEtRecreerWorksheet(sheetName As String)

    Dim ws As Worksheet
    Dim wsExists As Boolean

    'Check if the worksheet exists
    wsExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            wsExists = True
            Exit For
        End If
    Next ws

    'If the worksheet exists, delete it
    If wsExists Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    'Create a new worksheet with the specified name
    Set ws = ThisWorkbook.Worksheets.Add(Before:=wshMenu)
    ws.Name = sheetName
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub List_Formulas_All() '2024-06-22 @ 15:42
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    
    'Prepare existing worksheet to receive data
    Dim lastUsedRow As Long
    lastUsedRow = wshzDocFormulas.Cells(wshzDocFormulas.Rows.count, "A").End(xlUp).Row 'Last used row
    If lastUsedRow > 1 Then wshzDocFormulas.Range("A2:G" & lastUsedRow).ClearContents
    
    'Create an Array to receive the formulas informations
    Dim outputArray() As Variant
    ReDim outputArray(1 To 30000, 1 To 8)
    
    'Loop through each worksheet
    Dim ws As Worksheet
    Dim Name As String, usedRange As String, cellsCount As String
    For Each ws In wb.Sheets
        If ws.CodeName = "wshzDocNamedRange" Or _
            ws.CodeName = "wshzDocFormules" Then
                GoTo nextIteration
        End If
        'Save information for this worksheet
        Name = ws.Name
        usedRange = ws.usedRange.Address
        cellsCount = ws.usedRange.count
        'Loop through all cells in the used range
        Dim cell As Range
        Dim i As Long
        For Each cell In ws.usedRange
            'Does the cell contain a Formula
            If Left$(cell.formula, 1) = "=" Then
                'Write formula information to the destination worksheet
                i = i + 1
                If i Mod 50 Then Application.StatusBar = "J'ai traité " & i & " formules"
                outputArray(i, 1) = ws.CodeName & Chr$(0) & cell.Address
                outputArray(i, 2) = ws.CodeName
                outputArray(i, 3) = Name
                outputArray(i, 4) = usedRange
                outputArray(i, 5) = cellsCount
                outputArray(i, 6) = cell.Address
                outputArray(i, 7) = "'=" & Mid$(cell.formula, 2) 'Add ' to preserve formulas
                outputArray(i, 8) = Format$(Now(), "yyyy-mm-dd hh:mm") 'TimeStamp
            End If
        Next cell
nextIteration:
    Next ws
    
    Call RedimensionnerTableau2D(outputArray, i, UBound(outputArray, 2))
    Call TrierTableau2DBubble(outputArray)
    
    'Transfer the array data to the worksheet
    wshzDocFormulas.Range("A2").Resize(UBound(outputArray, 1), UBound(outputArray, 2)).Value = outputArray
    wshzDocFormulas.Range("A:A").EntireColumn.Hidden = True 'Do not show the outputArray

    MsgBox "J'ai trouvé " & Format$(i, "#,##0") & " formules"
    
    'Libérer la mémoire
    Set cell = Nothing
    Set wb = Nothing
    Set ws = Nothing

End Sub

Function HandleComments(ByVal codeLine As String, action As String) As String '2024-06-30 @ 10:45
    
    'R as action will remove the comments
    'U as action will UPPERCASE the comments
    
    Dim inString As Boolean: inString = False
    Dim codePart As String, commentPart As String
    
    Debug.Assert action = "R" Or action = "U"
    
    Dim i As Long, char As String
    For i = 1 To Len(codeLine)
        char = Mid$(codeLine, i, 1)
        
        'Toggle inString flag if a double quote is encountered
        If char = """" Then
            inString = Not inString
        End If
        
        'If the current character is ' and we are not within a string...
        If char = "'" Then
            If Not inString Then
                commentPart = Mid$(codeLine, i)
                Exit For
            Else
                codePart = codePart & char
            End If
        Else
            codePart = codePart & char
        End If
    Next i
    
    'Take action - R remove the comment from the code, L uppercase the comment
    If action = "R" Then
        commentPart = vbNullString
    Else
        commentPart = Trim$(UCase$(commentPart))
    End If
    
    HandleComments = codePart & commentPart
    
End Function

Sub List_All_Shapes_Properties() '2024-08-07 @ 19:37

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    
    Dim rng As Range
    Dim row As Long, col As Long
    row = ActiveCell.row
    col = ActiveCell.Column
    
    Application.EnableEvents = False
    
    Dim r As Long
    r = row
    ws.Cells(r, col).Value = "Type"
    ws.Cells(r, col + 1).Value = "Shape Name"
    ws.Cells(r, col + 2).Value = "ZOrder"
    ws.Cells(r, col + 3).Value = "Top"
    ws.Cells(r, col + 4).Value = "Left"
    ws.Cells(r, col + 5).Value = "Width"
    ws.Cells(r, col + 6).Value = "Height"
    
    'Loop through all shapes on the worksheet
    Dim shp As Shape
    r = row + 1
    For Each shp In ws.Shapes
        ws.Cells(r, col).Value = shp.Type
        ws.Cells(r, col + 1).Value = shp.Name
        ws.Cells(r, col + 2).Value = shp.ZOrderPosition
        ws.Cells(r, col + 3).Value = shp.Top
        ws.Cells(r, col + 4).Value = shp.Left
        ws.Cells(r, col + 5).Value = shp.Width
        ws.Cells(r, col + 6).Value = shp.Height
        r = r + 1
    Next shp
    
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub List_All_Tables()

    'Loop through each worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        'Loop through each ListObject (table) in the worksheet
        Dim lo As ListObject
        For Each lo In ws.ListObjects
            Debug.Print "#028 - Sheet: " & ws.Name; Tab(40); "Table: " & lo.Name & vbCrLf
        Next lo
    Next ws
    
    'Libérer la mémoire
    Set lo = Nothing
    Set ws = Nothing
    
End Sub

Sub List_Named_Ranges_All() '2024-06-23 @ 07:40
    
    'Setup and clear the output worksheet
    Dim ws As Worksheet: Set ws = wshzDocNamedRange
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    ws.Range("A2:I" & lastUsedRow).ClearContents
    
    'Loop through each named range in the workbook
    Dim arr() As Variant
    ReDim arr(1 To 300, 1 To 9)
    Dim i As Long
    Dim nr As Name
    Dim rng As Range
    Dim timeStamp As String
    For Each nr In ThisWorkbook.Names
        i = i + 1
        arr(i, 1) = UCase$(nr.Name) & Chr$(0) & UCase$(nr.RefersTo) 'Sort Key
        arr(i, 2) = nr.Name
        arr(i, 3) = "'" & nr.RefersTo
        If InStr(nr.RefersTo, "#REF!") Then
            arr(i, 4) = "'#REF!"
        End If
        
        'Check if the name refers to a range
        On Error Resume Next
        Set rng = nr.RefersToRange
        On Error GoTo 0
        
        If Not rng Is Nothing Then
            arr(i, 5) = rng.Worksheet.Name
            arr(i, 6) = rng.Address
        End If
         
        arr(i, 7) = nr.Comment
        If nr.Visible = False Then
            arr(i, 8) = nr.Visible
        End If
        timeStamp = Format$(Now(), "dd-mm-yyyy hh:mm:ss")
        arr(i, 9) = timeStamp
    Next nr
    
    Call RedimensionnerTableau2D(arr, i, UBound(arr, 2))
    Call TrierTableau2DBubble(arr)
    
    'Transfer the array data to the worksheet
    wshzDocNamedRange.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    wshzDocNamedRange.Range("A:A").EntireColumn.Hidden = True 'Do not show the outputArray
    
    'Result print setup - 2024-07-14 2 07:10
    If i > 1 Then
        Dim header1 As String: header1 = "List all Named Ranges"
        Dim header2 As String: header2 = vbNullString
        Call modAppli_Utils.MettreEnFormeImpressionSimple(wshzDocNamedRange, wshzDocNamedRange.Range("B2:I" & i), _
                               header1, _
                               header2, _
                               "$1:$1", _
                               "L")
    End If
   
    MsgBox "J'ai trouvé " & i & " named ranges"
    
    'Libérer la mémoire
    Set nr = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
End Sub

Sub shp_Reorganize_Tests_And_Todos_Click()

    Call Reorganize_Tests_And_Todos_Worksheet

End Sub

Sub Reorganize_Tests_And_Todos_Worksheet() '2024-03-02 @ 15:21

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshzDocTests_And_Todos
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Dim rng As Range: Set rng = ws.Range("A1:E" & lastUsedRow)
    
    With ws.ListObjects("tblTests_And_Todo").Sort
        Application.EnableEvents = False
        .SortFields.Clear
        .SortFields.Add2 _
            key:=ActiveSheet.Range("tblTests_And_Todo[Statut]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            key:=ActiveSheet.Range("tblTests_And_Todo[Module]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            key:=ActiveSheet.Range("tblTests_And_Todo[Priorité]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            key:=ActiveSheet.Range("tblTests_And_Todo[TimeStamp]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
        .Apply
        Application.EnableEvents = True
    End With
    
    Dim tbl As ListObject: Set tbl = ws.ListObjects("tblTests_And_Todo")
    Dim rowToMove As Range

    'Move completed item ($D = a) to the bottom of the list
    Dim i As Long, lastRow As Long
    i = 2

    Application.EnableEvents = False
    
    While ws.Range("D2").Value = "a"
        Set rowToMove = tbl.ListRows(1).Range
        lastRow = tbl.ListRows.count
        rowToMove.Cut Destination:=tbl.DataBodyRange.Rows(lastRow + 1)
        tbl.ListRows(1).Delete
    Wend

    ws.Calculate
    
    Application.EnableEvents = False
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set rng = Nothing
    Set rowToMove = Nothing
    Set tbl = Nothing
    Set ws = Nothing
    
End Sub

Sub Search_Every_Lines_Of_Code(arr As Variant, lignesLues As Long, search1 As String, search2 As String, search3 As String)

    'Declare arr() to keep results in memory
    Dim arrResult() As Variant
    ReDim arrResult(1 To 3000, 1 To 7)

    Dim saveLineOfCode As String, trimmedLineOfCode As String, procedureName As String
    Dim timeStamp As String
    Dim X As Long, xr As Long
    For X = LBound(arr, 1) To UBound(arr, 1)
        trimmedLineOfCode = arr(X, 4)
        saveLineOfCode = trimmedLineOfCode
        
        'Handle comments (second parameter is either Remove or Uppercase)
        If InStr(1, trimmedLineOfCode, "'") <> 0 Then
            trimmedLineOfCode = HandleComments(trimmedLineOfCode, "U")
        End If
        
        If trimmedLineOfCode <> vbNullString Then
            'Is this a procedure (Sub) declaration line ?
            If InStr(trimmedLineOfCode, "Sub ") <> 0 Then
                If InStr(trimmedLineOfCode, "End Sub") = 0 And _
                    InStr(trimmedLineOfCode, "Sub = ") = 0 And _
                    InStr(trimmedLineOfCode, "Sub As ") = 0 And _
                    InStr(trimmedLineOfCode, "Exit Sub") = 0 Then
                        procedureName = Mid$(saveLineOfCode, InStr(trimmedLineOfCode, "Sub "))
                End If
            End If
            
            If InStr(trimmedLineOfCode, "End Sub") = 1 Then
                procedureName = vbNullString
            End If

            'Is this a function declaration line ?
            If InStr(trimmedLineOfCode, "Function ") <> 0 Then
                If InStr(trimmedLineOfCode, "End Function") = 0 And _
                    InStr(trimmedLineOfCode, "Function = ") = 0 And _
                    InStr(trimmedLineOfCode, "Function As ") = 0 And _
                    InStr(trimmedLineOfCode, "Exit Function") = 0 Then
                        procedureName = Mid$(saveLineOfCode, InStr(trimmedLineOfCode, "Function "))
                End If
            End If
            
            If InStr(trimmedLineOfCode, "End Function") = 1 Then
                procedureName = vbNullString
            End If
            
            'Do we find the search1 or search2 or sreach3 strings in this line of code ?
            If (search1 <> vbNullString And InStr(trimmedLineOfCode, search1) <> 0) Or _
                (search2 <> vbNullString And InStr(trimmedLineOfCode, search2) <> 0) Or _
                (search3 <> vbNullString And InStr(trimmedLineOfCode, search3) <> 0) Then
                'Found an occurence
                xr = xr + 1
                arrResult(xr, 2) = arr(X, 1) 'oType
                arrResult(xr, 3) = arr(X, 2) 'oName
                arrResult(xr, 4) = arr(X, 3) 'LineNum
                arrResult(xr, 5) = procedureName
                arrResult(xr, 6) = "'" & saveLineOfCode
                timeStamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
                arrResult(xr, 7) = timeStamp
                arrResult(xr, 1) = UCase$(arr(X, 1)) & Chr$(0) & UCase$(arr(X, 2)) & Chr$(0) & Format$(arr(X, 3), "0000") & Chr$(0) & procedureName 'Future sort key
            End If
        End If
    Next X

    'Prepare the result worksheet
    Call EffacerEtRecreerWorksheet("X_Doc_Search_Utility_Results")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Doc_Search_Utility_Results")
    wsOutput.Range("A1").Value = "SortKey"
    wsOutput.Range("B1").Value = "Type"
    wsOutput.Range("C1").Value = "ModuleName"
    wsOutput.Range("D1").Value = "LineNo"
    wsOutput.Range("E1").Value = "ProcedureName"
    wsOutput.Range("F1").Value = "Code"
    wsOutput.Range("G1").Value = "TimeStamp"
    
    Call Make_It_As_Header(wsOutput.Range("A1:G1"), RGB(0, 112, 192))
    
    'Is there anything to show ?
    If xr > 0 Then
    
        'Data starts at row 2
        Dim r As Long: r = 2

        Call RedimensionnerTableau2D(arrResult, xr, UBound(arrResult, 2))
        
        'Sort the 2D array based on column 1
        Call TrierTableau2DBubble(arrResult)
    
        'Transfer the array to the worksheet
        wsOutput.Range("A2").Resize(UBound(arrResult, 1), UBound(arrResult, 2)).Value = arrResult
        wsOutput.Range("A:A").EntireColumn.Hidden = True 'Do not show the sortKey
        wsOutput.Columns(4).HorizontalAlignment = xlCenter
        wsOutput.Columns(7).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        Dim lastUsedRow As Long
        lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "B").End(xlUp).Row
        Dim j As Long, oldProcedure As String
        oldProcedure = wsOutput.Range("C" & lastUsedRow).Value & wsOutput.Range("E" & lastUsedRow).Value
        For j = lastUsedRow To 2 Step -1
            If wsOutput.Range("C" & j).Value & wsOutput.Range("E" & j).Value <> oldProcedure Then
                wsOutput.Rows(j + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                oldProcedure = wsOutput.Range("C" & j).Value & wsOutput.Range("E" & j).Value
            End If
        Next j
        
        'Since we might have inserted new row, let's update the lastUsedRow
        lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "B").End(xlUp).Row
        With wsOutput.Range("B2:G" & lastUsedRow)
            On Error Resume Next
            ActiveSheet.Cells.FormatConditions.Delete
            On Error GoTo 0
        
            .FormatConditions.Add Type:=xlExpression, Formula1:= _
                "=(MOD(LIGNE();2)=1)"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
            End With
            .FormatConditions(1).StopIfTrue = False
        End With
        
        wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit
    End If
    
    'Result print setup - 2024-07-14 2 06:24
    lastUsedRow = lastUsedRow + 2
    wsOutput.Range("B" & lastUsedRow).Value = "*** " & Format$(lignesLues, "###,##0") & " lignes de code dans l'application ***"
    Dim header1 As String: header1 = "Search Utility Results"
    Dim header2 As String
    header2 = "Searched strings '" & search1 & "'"
    If search2 <> vbNullString Then header2 = header2 & " '" & search2 & "'"
    If search3 <> vbNullString Then header2 = header2 & " '" & search3 & "'"
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsOutput, wsOutput.Range("B2:G" & lastUsedRow), _
                           header1, _
                           header2, _
                           "$1:$1", _
                           "L")
    
    'Display the final message
    If xr Then
        MsgBox "J'ai trouvé " & xr & " lignes avec les chaines '" & search1 & "'" & vbNewLine & _
                vbNewLine & "après avoir analysé un total de " & _
                Format$(lignesLues, "#,##0") & " lignes de code"
    Else
        MsgBox "Je n'ai trouvé aucune occurences avec les chaines '" & search1 & "'" & vbNewLine & _
                vbNewLine & "après avoir analysé un total de " & _
                Format$(lignesLues, "#,##0") & " lignes de code"
    End If
    
    'Libérer la mémoire
    Set wsOutput = Nothing
    
End Sub

Sub List_All_Columns() '2024-08-09 @ 11:52

    Dim colType As String
    
    'Erase & create a worksheet for the report
    Call EffacerEtRecreerWorksheet("Liste des Colonnes")
    Dim reportSheet As Worksheet: Set reportSheet = ThisWorkbook.Worksheets("Liste des Colonnes")
    
    'Add headers to the report
    With reportSheet
        .Cells(1, 1).Value = "Nom de la feuille"
        .Cells(1, 2).Value = "No. col."
        .Cells(1, 3).Value = "Lettre col."
        .Cells(1, 4).Value = "Nom Col."
        .Cells(1, 5).Value = "Type données"
        .Cells(1, 6).Value = "Largeur"
    End With
    
    Dim outputRow As Long
    outputRow = 2
    
    'Loop through each worksheet
    Dim ws As Worksheet
    Dim col As Range
    For Each ws In ThisWorkbook.Worksheets
        'Loop through each column in the worksheet
        Dim i As Long
        For i = 1 To ws.Range("A1").CurrentRegion.Columns.count
            Set col = ws.Cells(1, i).EntireColumn
            
            colType = Fn_Get_Column_Type(col)
            
            'Output the information to the report
            With reportSheet
                .Cells(outputRow, 1).Value = ws.Name
                .Cells(outputRow, 2).Value = i
                .Cells(outputRow, 3).Value = Replace(col.Address(False, False), "1", vbNullString)
                .Cells(outputRow, 4).Value = ws.Cells(1, i).Value
                .Cells(outputRow, 5).Value = colType
                .Cells(outputRow, 6).Value = col.ColumnWidth
            End With
            
            outputRow = outputRow + 1
        Next i
    Next ws
    
    'Sort the report by worksheet name and column number
    With reportSheet.Sort
        .SortFields.Clear
        .SortFields.Add key:=ActiveSheet.Range("A2"), Order:=xlAscending ' Worksheet name
        .SortFields.Add key:=ActiveSheet.Range("B2"), Order:=xlAscending ' Column number
        .SetRange ActiveSheet.Range("A1:F" & outputRow - 1)
        .Header = xlYes
        .Apply
    End With
    
    'Libérer la mémoire
    Set col = Nothing
    Set reportSheet = Nothing
    Set ws = Nothing
    
    MsgBox "Le rapport des colonnes a été généré avec succès !", vbInformation
    
End Sub

Sub List_All_Macros_Used_With_Objects() '2024-11-26 @ 20:14
    
    'Prepare the result worksheet
    Call EffacerEtRecreerWorksheet("Doc_All_Macros_Used_With_Object")

    Dim wsOutputSheet As Worksheet
    Set wsOutputSheet = ThisWorkbook.Worksheets("Doc_All_Macros_Used_With_Object")
    
    wsOutputSheet.Cells(1, 1).Value = "Worksheet"
    wsOutputSheet.Cells(1, 2).Value = "Object Type"
    wsOutputSheet.Cells(1, 3).Value = "Object Name"
    wsOutputSheet.Cells(1, 4).Value = "Macro Name"
    
    Call Make_It_As_Header(wsOutputSheet.Range("A1:D1"), RGB(0, 112, 192))

    Dim outputRow As Long
    outputRow = 2 'Start writing from the second row

    'Iterate through each worksheet in the workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim macroName As String

    For Each ws In ThisWorkbook.Worksheets
        'Skip the output sheet to avoid listing its own shapes
        If ws.Name <> "Doc_All_Macros_Used_With_Object" Then
            'Check for macros assigned to shapes
            For Each shp In ws.Shapes
                On Error Resume Next
                macroName = shp.OnAction
                On Error GoTo 0
                If macroName <> vbNullString Then
                    wsOutputSheet.Cells(outputRow, 1).Value = ws.Name
                    wsOutputSheet.Cells(outputRow, 2).Value = "Shape"
                    wsOutputSheet.Cells(outputRow, 3).Value = shp.Name
                    wsOutputSheet.Cells(outputRow, 4).Value = macroName
                    outputRow = outputRow + 1
                End If
            Next shp

            'Check for macros assigned to ActiveX controls
            Dim obj As OLEObject
            For Each obj In ws.OLEObjects
                On Error Resume Next
                If TypeOf obj.Object Is MSForms.CommandButton Then
                    macroName = obj.Object.OnClick
                ElseIf TypeOf obj.Object Is MSForms.ComboBox Then
                    macroName = obj.Object.OnChange
                ElseIf TypeOf obj.Object Is MSForms.ListBox Then
                    macroName = obj.Object.OnClick
                End If
                On Error GoTo 0
                If macroName <> vbNullString Then
                    wsOutputSheet.Cells(outputRow, 1).Value = ws.Name
                    wsOutputSheet.Cells(outputRow, 2).Value = "ActiveX Control"
                    wsOutputSheet.Cells(outputRow, 3).Value = obj.Name
                    wsOutputSheet.Cells(outputRow, 4).Value = macroName
                    outputRow = outputRow + 1
                End If
            Next obj
        End If
    Next ws

    'Autofit columns for better readability
    wsOutputSheet.Columns("A:D").AutoFit
    outputRow = outputRow - 1 'Did not use the last line
    
    'Sort the results, based on column 1, 2, 3 & 4
    If outputRow > 2 Then
        'Sort the data by columns 1, 2, 3, and 4
        With wsOutputSheet.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsOutputSheet.Range("A2:A" & outputRow - 1), Order:=xlAscending
            .SortFields.Add key:=wsOutputSheet.Range("B2:B" & outputRow - 1), Order:=xlAscending
            .SortFields.Add key:=wsOutputSheet.Range("C2:C" & outputRow - 1), Order:=xlAscending
            .SortFields.Add key:=wsOutputSheet.Range("D2:D" & outputRow - 1), Order:=xlAscending
            .SetRange wsOutputSheet.Range("A1:D" & outputRow - 1)
            .Header = xlYes
            .Apply
        End With
    End If
    
    'Set conditional formatting for the worksheet (alternate colors)
    outputRow = wsOutputSheet.Cells(wsOutputSheet.Rows.count, "A").End(xlUp).Row
    Dim rngArea As Range: Set rngArea = wsOutputSheet.Range("A2:D" & outputRow)
    Call modAppli_Utils.AppliquerConditionalFormating(rngArea, 1, RGB(173, 216, 230)) 'There are blankrows to account for
    
    outputRow = wsOutputSheet.Cells(wsOutputSheet.Rows.count, "A").End(xlUp).Row
    Dim rngToPrint As Range: Set rngToPrint = wsOutputSheet.Range("A2:D" & outputRow)
    Dim header1 As String: header1 = "Liste des macros associées à des contrôles"
    Dim header2 As String: header2 = ThisWorkbook.Name
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsOutputSheet, rngToPrint, header1, header2, "$1:$1", "P")
    
    MsgBox "La liste des macros assignées à des contrôles est dans " & _
                vbNewLine & vbNewLine & "la feuille 'Doc_All_Macros_Used_With_Object'.", vbInformation
                
    'Libérer la mémoire
    Set obj = Nothing
    Set wsOutputSheet = Nothing
    Set rngArea = Nothing
    Set rngToPrint = Nothing
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub List_Subs_And_Functions_All() '2024-11-26 @ 20:02
    
    Dim ws As Worksheet: Set ws = wshzDocSubsAndFunctions
    
    Dim posProcedure As Long, posExitProcedure As Long
    Dim posFonction As Long, posExitFonction As Long
    Dim posSpace As Long, posREM As Long, posParam As Long
    Dim scope As String, sType As String
    
    'Loop through all VBcomponents (modules, class and forms) in the active workbook
    Dim vbComp As Object
    Dim oType As String
    Dim arr() As Variant
    ReDim arr(1 To 1500, 1 To 10)
    Dim trimmedLineOfCode As String, savedLineOfCode As String, remarks As String, params As String
    Dim LineNum As Long, lread As Long
    Dim i As Long

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        'Check if the component is a userForm (1), a module (2) or a class module (3)
        If vbComp.Type <= 100 Then
            Select Case vbComp.Type
                Case 1
                    oType = "1_Module"
                Case 2
                    oType = "2_Class"
                Case 3
                    oType = "3_userform"
                Case 100
                    oType = "0_Worksheet"
                Case Else
                    oType = vbComp.Type & "_?????"
            End Select
            'Get the code module for the component
            Dim vbCodeMod As Object: Set vbCodeMod = vbComp.codeModule
            'Loop through all lines in the code module
            For LineNum = 1 To vbCodeMod.CountOfLines
                lread = lread + 1
                'Check if the line contains 'Sub' or 'Function' without beeing a Remark line
                savedLineOfCode = Trim$(vbCodeMod.Lines(LineNum, 1))
                trimmedLineOfCode = Trim$(vbCodeMod.Lines(LineNum, 1))
                'Remove comments
                If InStr(1, trimmedLineOfCode, "'") Then
                    trimmedLineOfCode = HandleComments(trimmedLineOfCode, "U")
                End If
                
                posProcedure = InStr(trimmedLineOfCode, "Sub ")
                If posProcedure Then
                    If posProcedure = InStr(trimmedLineOfCode, "Sub = ") Or _
                        posProcedure = InStr(trimmedLineOfCode, "Sub As ") Then
                        posProcedure = 0
                    End If
                End If
                posFonction = InStr(trimmedLineOfCode, "Function ")
                If posFonction Then
                    If posFonction = InStr(trimmedLineOfCode, "Function = ") Or _
                        posFonction = InStr(trimmedLineOfCode, "Function As ") Then
                        posFonction = 0
                    End If
                End If
                posExitProcedure = InStr(trimmedLineOfCode, "Exit Sub")
                posExitFonction = InStr(trimmedLineOfCode, "Exit Function")
                If (posProcedure <> 0 Or posFonction <> 0) And posExitProcedure = 0 And posExitFonction = 0 Then
                    i = i + 1
                    arr(i, 2) = oType
                    arr(i, 3) = vbComp.Name
                    arr(i, 4) = LineNum
                    'Goback to savedLineOfCode
                    trimmedLineOfCode = Trim$(vbCodeMod.Lines(LineNum, 1))
                    posREM = InStr(trimmedLineOfCode, ") '")
                    If posREM > 0 Then
                        remarks = Trim$(Mid$(trimmedLineOfCode, posREM + 2))
                        trimmedLineOfCode = Trim$(Left$(trimmedLineOfCode, posREM))
                    End If
                    posParam = InStr(trimmedLineOfCode, "(")
                    If posParam > 0 Then
                        params = Trim$(Mid$(trimmedLineOfCode, posParam))
                        trimmedLineOfCode = Trim$(Left$(trimmedLineOfCode, posParam - 1))
                    End If
                    
                    If InStr(trimmedLineOfCode, "Sub ") > 1 Or InStr(trimmedLineOfCode, "Function ") > 1 Then
                        posSpace = InStr(trimmedLineOfCode, " ")
                        scope = Left$(trimmedLineOfCode, posSpace - 1)
                        trimmedLineOfCode = Trim$(Mid$(trimmedLineOfCode, posSpace + 1))
                    Else
                        scope = vbNullString
                    End If
                    arr(i, 5) = scope
                    If InStr(trimmedLineOfCode, "Sub ") = 1 Then
                        sType = "Sub"
                        trimmedLineOfCode = Trim$(Mid$(trimmedLineOfCode, 5))
                    Else
                        If InStr(trimmedLineOfCode, "Function ") = 1 Then
                            sType = "Function"
                            trimmedLineOfCode = Trim$(Mid$(trimmedLineOfCode, 10))
                        End If
                    End If
                    arr(i, 6) = sType
                    arr(i, 7) = trimmedLineOfCode
                    arr(i, 1) = UCase$(oType) & Chr$(0) & UCase$(vbComp.Name) & Chr$(0) & UCase$(trimmedLineOfCode) 'Future sort key
                    If params <> "()" Then arr(i, 8) = params
                    If remarks <> vbNullString Then arr(i, 9) = remarks
                    arr(i, 10) = Format$(Now(), "yyyy-mm-dd hh:mm")
                    params = vbNullString
                    remarks = vbNullString
                End If
            Next LineNum
        End If
    Next vbComp
    
    'Prepare the output worksheet
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    ws.Range("A2:J" & lastUsedRow).ClearContents

    Call RedimensionnerTableau2D(arr, i, UBound(arr, 2))
    
    'Sort the 2D array based on column 1
    Call TrierTableau2DBubble(arr)
    
    'Transfer the array to the worksheet
    ws.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    ws.Range("A:A").EntireColumn.Hidden = True 'Do not show the sortKey
    
    ws.Activate
    
    MsgBox "J'ai trouvé " & i & " lignes Sub or Function" & vbNewLine & _
                vbNewLine & "après avoir analysé un total de " & _
                Format$(lread, "#,##0") & " Lignes de code"
    
    'Libérer la mémoire
    Set vbComp = Nothing
    Set vbCodeMod = Nothing
    Set ws = Nothing
    
End Sub

Sub Test_Array_To_Range() '2024-03-18 @ 17:34

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(2)
    
    Dim arr() As Variant
    ReDim arr(1 To 1000, 1 To 20)
    
    Dim i As Long, j As Long
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            arr(i, j) = "i = " & i & " and j = " & j & " - *********"
        Next j
    Next i
    
    ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub TestRedimensionnerTableau2D()
    Dim originalArray() As Variant
    Dim i As Long, j As Long
    
    'Dimension the original array to a fixed size (e.g., 10 rows and 5 columns)
    ReDim originalArray(1 To 10, 1 To 5)
    
    'Fill the original array with some example data
    For i = 1 To 10
        For j = 1 To 5
            originalArray(i, j) = "R" & i & "C" & j
        Next j
    Next i
    
    'Output the original array to the immediate window
    Debug.Print "#030 - Original Array:"
    For i = 1 To 10
        For j = 1 To 5
            Debug.Print "#031 - " & originalArray(i, j);
        Next j
        Debug.Print "#032"
    Next i
    
    ' Trim the array to 6 rows and 3 columns
    Call RedimensionnerTableau2D(originalArray, 6, 3)
    
    ' Output the trimmed array to the immediate window
    Debug.Print "#033 - Trimmed Array:"
    For i = 1 To 6
        For j = 1 To 3
            Debug.Print "#034 - " & originalArray(i, j);
        Next j
        Debug.Print "#035"
    Next i
End Sub

Sub Toggle_A1_R1C1_Reference()

    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If

End Sub

Sub List_Worksheets_From_Current_Workbook_All() '2024-07-24 @ 10:14
    
    Call EffacerEtRecreerWorksheet("X_Feuilles_du_Classeur")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("X_Feuilles_du_Classeur")
    wsOutput.Range("A1").Value = "Feuille"
    wsOutput.Range("B1").Value = "CodeName"
    wsOutput.Range("C1").Value = "TimeStamp"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"), RGB(0, 112, 192))

    'Loop through all worksheets in the active workbook
    Dim arr() As Variant
    ReDim arr(1 To 100, 1 To 3)
    
    Dim ws As Worksheet
    Dim timeStamp As String
    Dim i As Long
    For Each ws In ThisWorkbook.Sheets
        i = i + 1
        arr(i, 1) = ws.Name
        arr(i, 2) = ws.CodeName
        timeStamp = Format$(Now(), "dd-mm-yyyy hh:mm:ss")
        arr(i, 3) = timeStamp
    Next ws
    
    Call RedimensionnerTableau2D(arr, i, UBound(arr, 2))
    
    Call TrierTableau2DBubble(arr)
    
    Dim f As Long
    For i = 1 To UBound(arr, 1)
        wsOutput.Cells(i + 1, 1) = arr(i, 1)
        wsOutput.Cells(i + 1, 2) = arr(i, 2)
        wsOutput.Cells(i + 1, 3) = arr(i, 3)
        f = f + 1
    Next i
    
    wsOutput.Columns.AutoFit
    
   'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = i + 2
    wsOutput.Range("A" & lastUsedRow).Value = "*** " & Format$(f, "###,##0") & _
                                    " feuilles pour le workbook '" & ThisWorkbook.Name & "' ***"
    
    lastUsedRow = wsOutput.Cells(wsOutput.Rows.count, "A").End(xlUp).Row
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Liste des feuilles d'un classeur"
    Dim header2 As String: header2 = ThisWorkbook.Name
    Call modAppli_Utils.MettreEnFormeImpressionSimple(wsOutput, rngToPrint, header1, header2, "$1:$1", "P")
    
    ThisWorkbook.Worksheets("X_Feuilles_du_Classeur").Activate
    
    'Libérer la mémoire
    Set rngToPrint = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing

End Sub

Sub DeterminerOrdreDeTabulation(ws As Worksheet) '2024-06-15 @ 13:58

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modDev_Utils:DeterminerOrdreDeTabulation", ws.CodeName, 0)

    'Clear previous settings AND protect the worksheet
    With ws
        .Protect userInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

    'Collect all unprotected cells
    Dim cell As Range
    Dim unprotectedCells As Range
    Application.ScreenUpdating = False
    For Each cell In ws.usedRange
        If Not cell.Locked Then
'            Debug.Print cell.Address
            If unprotectedCells Is Nothing Then
                Set unprotectedCells = cell
            Else
                Set unprotectedCells = Union(unprotectedCells, cell)
            End If
        End If
    Next cell

    'Sort to ensure cells are sorted left-to-right, top-to-bottom
    If Not unprotectedCells Is Nothing Then
        Dim sortedCells As Range: Set sortedCells = unprotectedCells
        Debug.Print "(" & ws.Name & ") - DeterminerOrdreDeTabulation - Unprotected cells are '" & sortedCells.Address & "' - " & sortedCells.count & " cellule(s) - " & Format$(Now(), "dd/mm/yyyy hh:mm:ss")

        'Enable TAB through unprotected cells
        Application.EnableEvents = False
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    'Libérer la mémoire
    Set cell = Nothing
    Set unprotectedCells = Nothing
    Set sortedCells = Nothing

    Call EnregistrerLogApplication("modDev_Utils:DeterminerOrdreDeTabulation", vbNullString, startTime)

End Sub

Sub EnregistrerLogApplication(ByVal procedureName As String, param As String, Optional ByVal startTime As Double = 0) '2025-02-03 @ 17:17

    'En attendant de trouver la problématique... 2025-06-01 @ 05:06
    If gUtilisateurWindows = vbNullString Then
        gUtilisateurWindows = Fn_Get_Windows_Username
        Debug.Print "Réinitialisation forcée de gUtilisateurWindows - " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    End If
    
    On Error GoTo ErrorHandler
    
    'TimeStamp avec centièmes de seconde
    Dim timeStamp As String
    timeStamp = Format$(Now, "yyyy-mm-dd hh:mm:ss") & "." & Right$(Format$(Timer, "0.00"), 2)
    
    Dim logFile As String
    logFile = wsdADMIN.Range("F5").Value & gDATA_PATH & _
                                    Application.PathSeparator & "LogMainApp.log"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    
    'On laisse une ligne blanche dans le fichier Log
    If Trim$(procedureName) = vbNullString Then
        Print #fileNum, vbNullString
    ElseIf startTime = 0 Then 'On marque le départ d'une procédure/fonction
        Print #fileNum, timeStamp & " | " & _
                        GetNomUtilisateur() & " | " & _
                        ThisWorkbook.Name & " | " & _
                        procedureName & " | " & _
                        param & " | "
    ElseIf startTime < 0 Then 'On enregistre une entrée intermédiaire (au coeur d'un procédure/fonction)
        Print #fileNum, timeStamp & " | " & _
                        GetNomUtilisateur() & " | " & _
                        ThisWorkbook.Name & " | " & _
                        procedureName & " | " & _
                        param & " | "
    Else 'On marque la fin d'une procédure/fonction
        Dim elapsedTime As Double
        elapsedTime = Round(Timer - startTime, 4) 'Calculate elapsed time
        Print #fileNum, timeStamp & " | " & _
                        GetNomUtilisateur() & " | " & _
                        ThisWorkbook.Name & " | " & _
                        procedureName & " | " & _
                        param & " | " & _
                        Format$(elapsedTime, "0.0000") & " secondes" & vbCrLf
    End If
    
    Close #fileNum
    
    Exit Sub
    
ErrorHandler:

    MsgBox "Une erreur est survenue à l'ouverture du fichier 'LogMainApp.log' " & vbNewLine & vbNewLine & _
                "Erreur : " & Err & " = " & Err.description, vbCritical, "Répertoire utilisé '" & wsdADMIN.Range("F5").Value & "'"
    
    'Nettoyage : réactivation des événements, calculs, etc.
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    'Fermeture des classeurs sans sauvegarde si nécessaire
    On Error Resume Next 'Ignorer les erreurs pendant la fermeture des fichiers
    ThisWorkbook.Close SaveChanges:=False

    'Sortir gracieusement de l'application
    Application.Quit
    
End Sub

Sub Test_EnregistrerLogApplication()

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modDev_Utils:Test_EnregistrerLogApplication", vbNullString, 0)

    Call EnregistrerLogApplication("modDev_Utils:Test_EnregistrerLogApplication", vbNullString, startTime)
    
End Sub

Sub Log_Saisie_Heures(oper As String, txt As String, Optional blankline As Boolean = False) '2024-09-14 @ 06:56

    On Error GoTo Error_Handler
    
    'Détermine si cette entrée sera ou non sauvegardée dans le log
    If InStr(oper, "ADD") = 0 And _
        InStr(oper, "UPDATE") = 0 And _
        InStr(oper, "DELETE") = 0 Then
        If gLogSaisieHeuresVeryDetailed = False Then
            Exit Sub
        End If
    End If
    
    'TimeStamp avec les centièmes de secondes
    Dim ms As String
    Dim timeStamp As String
    timeStamp = Format$(Now, "yyyy-mm-dd hh:mm:ss") & "." & Right$(Format$(Timer, "0.00"), 2)
    
    'Path complet du fichier LogSaisieHeures.txt
    Dim logSaisieHeuresFile As String
    logSaisieHeuresFile = wsdADMIN.Range("F5").Value & gDATA_PATH & _
                                Application.PathSeparator & "LogSaisieHeures.log"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logSaisieHeuresFile For Append As #fileNum
    
    If blankline = True Then
        Print #fileNum, vbNullString
    End If
    
    Print #fileNum, timeStamp & " | " & _
                        Left$(GetNomUtilisateur() & Space(19), 19) & " | " & _
                        ThisWorkbook.Name & " | " & _
                        oper & " | " & _
                        txt
    Close #fileNum
    
    Exit Sub
    
Error_Handler:

    MsgBox "Une erreur est survenue : " & Err.description, vbCritical, "Log_Saisie_Heures"
    'Sortir gracieusement de l'application
    Application.Quit 'No save...
    
End Sub

Sub Test_Log_Saisie_Heures()

    Call Log_Saisie_Heures("W", "Test")
    
End Sub

Sub Settrace(Source As String, module As String, procedure As String, variable As String, vType As String) '2024-09-26 @ 10:31

    On Error GoTo Error_Handler
    
    Dim ms As String
    
    Dim settraceFile As String
    settraceFile = wsdADMIN.Range("F5").Value & gDATA_PATH & _
        Application.PathSeparator & "LogSettrace.txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    'Ajoute les millisecondes à la chaîne de temps
    ms = Right$(Format$(Timer, "0.00"), 2) 'Récupère les millisecondes sous forme de texte
    
    Dim timeStamp As String
    timeStamp = Format$(Now, "yyyy-mm-dd hh:mm:ss") & "." & ms
    
    Open settraceFile For Append As #fileNum
    
    Print #fileNum, timeStamp & " | " & _
                    GetNomUtilisateur() & " | " & _
                    Source & " | " & _
                    module & " | " & _
                    procedure & " | " & _
                    variable & " | " & _
                    vType

    Close #fileNum
    
    Exit Sub
    
Error_Handler:

    MsgBox "Une erreur est survenue : " & Err.description, vbCritical, "Log_Settrace"
    'Sortir gracieusement de l'application
    Application.Quit 'No save...
    
End Sub

Sub Test_Settrace()

    Call Settrace("DB.1854", "modDev_Utils", "Test_Settrace", "date = '" & Date & "'", "type = " & "Date")
    
End Sub

Sub SortDelimitedString(ByRef inputString As String, delimiter As String)
    
    'Split the string into components
    Dim components() As String
    components = Split(inputString, delimiter)
    
    'Sort components (simple bubble sort)
    Dim i As Long, j As Long
    Dim intResult As Integer
    Dim temp As String
    For i = LBound(components) To UBound(components) - 1
        For j = i + 1 To UBound(components)
            intResult = StrComp(components(i), components(j), vbTextCompare)
            If intResult = 1 Then
                'Swap components
                temp = components(i)
                components(i) = components(j)
                components(j) = temp
            End If
        Next j
    Next i
    
    'Rejoin the sorted components into a single string
    inputString = Join(components, delimiter)
    If Left$(inputString, 1) = "|" Then
        inputString = Right$(inputString, Len(inputString) - 1)
    End If
    
End Sub

Sub LogMainApp_Analysis() '2025-01-10 @ 17:10

    Dim logFile As String
    logFile = wsdADMIN.Range("F5").Value & Application.PathSeparator & "LogMainApp.log"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Dim currentTime As String
    currentTime = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    
    Open logFile For Input As #fileNum

    Dim strUser As String, strDate As String, strVersion As String, strModule As String
    Dim logline As String
    Dim arrTime(1 To 10) As Long
    Dim ctr As Long
    Do Until EOF(fileNum)
        Line Input #fileNum, logline
        ctr = ctr + 1
        logline = logline & "|"
        Dim arr() As String
        arr = Split(logline, "|")
        If InStr(strUser, arr(0)) = 0 Then
            strUser = strUser + Fn_Pad_A_String(arr(0), " ", 15, "R") & "|"
        End If
        If InStr(strDate, Left$(arr(1), 10)) = 0 Then
            strDate = strDate & Left$(arr(1), 10) & "|"
        End If
        If InStr(strVersion, arr(2)) = 0 Then
            strVersion = strVersion & Fn_Pad_A_String(arr(2), " ", 10, "R") & "|"
        End If
        If InStr(strModule, arr(3)) = 0 Then
            strModule = strModule & Fn_Pad_A_String(arr(3), " ", 65, "R") & "|"
        End If
        Dim subString As String
        Dim e As Double
        If InStr(arr(4), "Temps écoulé: ") > 0 Then
            subString = Mid$(arr(4), InStr(arr(4), "Temps écoulé: ") + 14)
            subString = Replace(subString, ".", ",")
            e = Left$(subString, InStr(subString, " ") - 1)
            If e <= 0.2 Then
                arrTime(1) = arrTime(1) + 1
            ElseIf e <= 0.3 Then
                arrTime(2) = arrTime(2) + 1
            ElseIf e <= 0.4 Then
                arrTime(3) = arrTime(3) + 1
            ElseIf e <= 0.5 Then
                arrTime(4) = arrTime(4) + 1
            ElseIf e <= 0.6 Then
                arrTime(5) = arrTime(5) + 1
            ElseIf e <= 0.7 Then
                arrTime(6) = arrTime(6) + 1
            ElseIf e <= 0.8 Then
                arrTime(7) = arrTime(7) + 1
            ElseIf e <= 0.9 Then
                arrTime(8) = arrTime(8) + 1
            ElseIf e <= 1 Then
                arrTime(9) = arrTime(9) + 1
            Else
                arrTime(10) = arrTime(10) + 1
            End If
            If e > 1 Then Debug.Print "#041 - " & ctr, Format$(e, "0.0000"), arr(2), arr(3)
        End If
    Loop
    
    Debug.Print "#042 - " & arrTime(1), arrTime(2), arrTime(3), arrTime(4), arrTime(5), arrTime(6), arrTime(7), arrTime(8), arrTime(9), arrTime(10)
    
    Call SortDelimitedString(strUser, "|")
    Dim arrUser() As String
    arrUser = Split(strUser, "|")
    
    Call SortDelimitedString(strDate, "|")
    Dim arrDate() As String
    arrDate = Split(strDate, "|")
    
    Call SortDelimitedString(strVersion, "|")
    Dim arrVersion() As String
    arrVersion = Split(strVersion, "|")

    Call SortDelimitedString(strModule, "|")
    Dim arrModule() As String
    arrModule = Split(strModule, "|")
    
    MsgBox "Il y a " & UBound(arrUser, 1) + 1 & " utilisateurs dans le log", vbInformation
    MsgBox "Il y a " & UBound(arrDate, 1) + 1 & " jours dans le log", vbInformation
    MsgBox "Il y a " & UBound(arrVersion, 1) + 1 & " versions d'application dans le log", vbInformation
    MsgBox "Il y a " & UBound(arrModule, 1) + 1 & " modules distincts dans le log", vbInformation
    
    'Close the file
    Close #fileNum
    
End Sub

Sub Test_Fn_Get_A_Row_From_A_Worksheet() '2025-01-13 @ 08:49

    Dim feuille As String
    Dim valeurRecherche As String
    Dim colRecherche As Integer
    Dim resultat As Variant
    Dim i As Long
    
    'Définir la feuille, la valeur à rechercher et la colonne
    feuille = "BD_Clients"
    valeurRecherche = "9299-2585 Québec Inc. [Informat] Marie Guay Isabelle Vigneault"
    colRecherche = 17
    
    'Appeler la fonction de recherche
    resultat = Fn_Get_A_Row_From_A_Worksheet(feuille, valeurRecherche, fClntFMNomClientPlusNomClientSystème)
    
    'Vérifier le résultat
    If IsArray(resultat) Then
        Debug.Print "#080 - Valeur trouvée :"
        For i = LBound(resultat) To UBound(resultat)
            Debug.Print "#081 - " & i; Tab(13); resultat(i)
        Next i
    Else
        MsgBox "Valeur non trouvée", vbInformation
    End If
    
End Sub

