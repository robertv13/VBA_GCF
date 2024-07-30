Attribute VB_Name = "modzDevUtils"
Option Explicit

Sub Add_Columns_To_Active_Worksheet()

    Dim colToAdd As Integer
    colToAdd = 5
    
    'Set the worksheet
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    'Find the last column with data
    Dim lastColumn As Integer
    lastColumn = ws.Cells(1, ws.columns.count).End(xlToLeft).Column
    
    'Add columns to the right of the last column
    ws.columns(lastColumn + 1).Resize(, colToAdd).Insert Shift:=xlToRight
    
    Debug.Print colToAdd & " columns added to the worksheet."
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Sub Array_2D_Bubble_Sort(ByRef arr() As Variant) '2024-06-23 @ 07:05
    
    Dim i As Long, j As Long, numRows As Long, numCols As Long
    Dim temp As Variant
    Dim sorted As Boolean
    
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    'Bubble Sort Algorithm
    Dim c As Integer, cProcess As Long
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

Sub Array_2D_Resizer(ByRef inputArray As Variant, ByVal nRows As Long, ByVal nCols As Long)
    
    Dim oRows As Long, oCols As Long
    
    'Get the original dimensions of the input array
    oRows = UBound(inputArray, 1)
    oCols = UBound(inputArray, 2)
    
    'Ensure the new dimensions are within the original array's bounds
    If nRows > oRows Then nRows = oRows
    If nCols > oCols Then nCols = oCols
    
    'Create a new array with the specified dimensions
    Dim tempArray() As Variant
    ReDim tempArray(1 To nRows, 1 To nCols)
    
    ' Copy the relevant data from the input array to the new array
    Dim i As Long, j As Long
    For i = 1 To nRows
        For j = 1 To nCols
            tempArray(i, j) = inputArray(i, j)
        Next j
    Next i
    
    ' Assign the trimmed array back to the input array
    inputArray = tempArray
    
End Sub

Sub Bubble_Sort_1D_Array(arr() As String)
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

Sub BubbleSort(MyArray() As String) '2024-07-02 @ 15:18 - WellSR.com
    'Sorts a one-dimensional VBA array from smallest to largest 'using the bubble sort algorithm.
    'HOW TO USE: Call BubbleSort(MyArray())
    
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

Sub Build_File_Layouts() '2024-03-26 @ 14:35

    Dim arr(1 To 20, 1 To 2) As Variant
    Dim output(1 To 150, 1 To 5) As Variant
    Dim r As Long
    r = 0
    r = r + 1: arr(r, 1) = "AR_Entête": arr(r, 2) = "A2:J2"
    r = r + 1: arr(r, 1) = "BD_Clients": arr(r, 2) = "A1:J1"
    r = r + 1: arr(r, 1) = "Doc_ConditionalFormatting": arr(r, 2) = "A1:E1"
    r = r + 1: arr(r, 1) = "Doc_Formules": arr(r, 2) = "A1:H1"
    r = r + 1: arr(r, 1) = "Doc_Log_Appli": arr(r, 2) = "A1:C1"
    r = r + 1: arr(r, 1) = "Doc_NamedRanges": arr(r, 2) = "A1:B1"
    r = r + 1: arr(r, 1) = "Doc_Subs&Functions": arr(r, 2) = "A1:G1"
    r = r + 1: arr(r, 1) = "ENC_Entête": arr(r, 2) = "A3:F3"
    r = r + 1: arr(r, 1) = "ENC_Détail": arr(r, 2) = "A3:F3"
    r = r + 1: arr(r, 1) = "FAC_Entête": arr(r, 2) = "A3:T3"
    r = r + 1: arr(r, 1) = "FAC_Détails": arr(r, 2) = "A3:G3"
    r = r + 1: arr(r, 1) = "GL_Trans": arr(r, 2) = "A1:J1"
    r = r + 1: arr(r, 1) = "GL_EJ_Auto": arr(r, 2) = "C1:J1"
    r = r + 1: arr(r, 1) = "Invoice List": arr(r, 2) = "A2:J2"
    r = r + 1: arr(r, 1) = "TEC_Local": arr(r, 2) = "A2:P2"
    r = 1
    Dim i As Long, colNo As Integer
    For i = 1 To UBound(arr, 1)
        If arr(i, 1) = "" Then Exit For
        Dim rng As Range: Set rng = Sheets(arr(i, 1)).Range(arr(i, 2))
        colNo = 0
        Dim cell As Range
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
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rng = Nothing
    Set cell = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub Check_Invoice_Template()

    Dim ws As Worksheet: Set ws = wshAdmin
    Dim firstUsedRow As Long, lastUsedRow As Long
    firstUsedRow = 12
    lastUsedRow = ws.Range("Z9999").End(xlUp).row
    Dim rng As Range
    Set rng = ws.Range("Z" & firstUsedRow & ":AA" & lastUsedRow)
    
    'First - Determine which templates are used
    Dim arr As Variant
    Dim strTemplates As String
    Dim i As Integer, j As Integer
    For i = 1 To lastUsedRow - firstUsedRow + 1
        If Not rng.Cells(i, 2) = "" Then
            arr = Split(rng.Cells(i, 2), ",")
            For j = 0 To UBound(arr)
                strTemplates = strTemplates & Trim(arr(j)) & "-" & i & "|"
            Next j
        End If
    Next i
    
    'Second - Sort all the found templates
    Dim tt() As String
    tt = Split(strTemplates, "|")
    Call BubbleSort(tt)
    
    'Third - Prepare the worksheet to receive information
    Call Erase_And_Create_Worksheet("Gabarits_Facture")
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Gabarits_Facture")
    wsOutput.Range("A1").value = "Gabarit"
    wsOutput.Range("B1").value = "Code"
    wsOutput.Range("C1").value = "Service"
    Dim outputRow As Integer: outputRow = 1
    
    'Third - Build the list of services associated to each template (First Letter)
    Dim rowNo As Integer
    Dim template As String, oldTemplate As String
    
    With wsOutput
        For i = 0 To UBound(tt)
            If tt(i) <> "" Then
                template = Left(tt(i), 1)
                If template <> oldTemplate Then
                    outputRow = outputRow + 2
                    .Range("A" & outputRow).value = "Gabarit '" & template & "'"
                    oldTemplate = template
                End If
                rowNo = Mid(tt(i), InStr(1, tt(i), "-") + 1)
                outputRow = outputRow + 1
                .Range("B" & outputRow).value = tt(i)
                .Range("C" & outputRow).value = rng.Cells(rowNo, 1)
            End If
        Next i
        wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit
    End With
    
    With wsOutput.Range("A1:C1")
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        With .Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .Size = 10
            .Italic = True
            .Bold = True
        End With
    End With
    
    With wsOutput.Range("A2:A" & outputRow)
        .Font.Bold = True
    End With

    'Cleaning - 2024-07-02 @ 20:12
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub List_Worksheets_From_Closed_Workbook_All() '2024-07-14 @ 07:02
    
    Call Erase_And_Create_Worksheet("Feuilles_du_Classeur")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Feuilles_du_Classeur")
    wsOutput.Range("A1").value = "Feuille"
    wsOutput.Range("B1").value = "CodeName"
    wsOutput.Range("C1").value = "TimeStamp"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"))

    'Specify the full path and name of the closed workbook
    Dim wbPath As String
    wbPath = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Sortie.xlsx"
    
    'Open the workbook in read-only mode
    Dim wb As Workbook: Set wb = Workbooks.Open(fileName:=wbPath, ReadOnly:=True)
    
    Dim wbName As String
    wbName = wb.name
    
    'Loop through each worksheet in the workbook and add its name to the immediate window
    Dim arr() As Variant
    ReDim arr(1 To 100, 1 To 3)
    Dim ws As Worksheet
    Dim TimeStamp As String
    Dim r As Integer
    Dim f As Integer
    For Each ws In wb.Worksheets
        r = r + 1
        arr(r, 1) = ws.name
        arr(r, 2) = ws.CodeName
        TimeStamp = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        arr(r, 3) = TimeStamp
    Next ws
    
    'Close the workbook without saving changes
    wb.Close SaveChanges:=False
    
    Call Array_2D_Resizer(arr, r, UBound(arr, 2))
    
    Call Array_2D_Bubble_Sort(arr)
    
    For r = 1 To UBound(arr, 1)
        wsOutput.Cells(r + 1, 1) = arr(r, 1)
        wsOutput.Cells(r + 1, 2) = arr(r, 2)
        wsOutput.Cells(r + 1, 3) = arr(r, 3)
        f = f + 1
    Next r
    
    wsOutput.columns.AutoFit
    
   'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = r + 2
    wsOutput.Range("A" & lastUsedRow).value = "*** " & Format(f, "###,##0") & _
                                    " feuilles pour le workbook '" & wbName & "' ***"
    
    lastUsedRow = wsOutput.Range("A9999").End(xlUp).row
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Liste des feuilles d'un classeur"
    Dim header2 As String: header2 = wbName
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "P")
    
    ThisWorkbook.Worksheets("Feuilles_du_Classeur").Activate
    
    'Cleanup - 2024-07-14 @ 07:03
    Set rngToPrint = Nothing
    Set wb = Nothing
    Set wsOutput = Nothing
    
End Sub

Sub Code_Search_Everywhere() '2024-07-11 @ 06:27
    
    'Declare lineOfCode() as variant
    Dim allLinesOfCode As Variant
    ReDim allLinesOfCode(1 To 15000, 1 To 4)
    
    Application.ScreenUpdating = False
    'Allows up to 3 search strings
    Dim search1 As String, search2 As String, search3 As String
    search1 = InputBox("Enter the search string ? ", "Search1")
    search2 = InputBox("Enter the search string ? ", "Search2")
    search3 = InputBox("Enter the search string ? ", "Search3")
    Application.ScreenUpdating = True
    
    'Loop through all VBcomponents (modules, class and forms) in the active workbook
    Dim oType As String
    Dim lineNum As Long
    Dim x As Integer
    Dim vbComp As Object
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
            oType = oType & "_?????"
            Stop
        End Select
        
        'Get the code module for the component
        Dim vbCodeMod As Object: Set vbCodeMod = vbComp.CodeModule
        
        'Loop through all lines in the code module to save all the lines in memory
        For lineNum = 1 To vbCodeMod.CountOfLines
            If Trim(vbCodeMod.Lines(lineNum, 1)) <> "" Then
                x = x + 1
                allLinesOfCode(x, 1) = oType
                allLinesOfCode(x, 2) = vbComp.name
                allLinesOfCode(x, 3) = lineNum
                allLinesOfCode(x, 4) = Trim(vbCodeMod.Lines(lineNum, 1))
            End If
        Next lineNum
    Next vbComp
    
    'At this point allLinesOfCode contains all lines of code of the application - 2024-07-10 @ 17:33
    
    Call Array_2D_Resizer(allLinesOfCode, x, UBound(allLinesOfCode, 2))
    
    Call Search_Every_Lines_Of_Code(allLinesOfCode, search1, search2, search3)
    
    'Clean up memory - 2024-07-10 @ 18:08
    Set vbComp = Nothing
    Set vbCodeMod = Nothing
    
End Sub

'@EntryPoint
Sub List_Conditional_Formatting_All() '2024-06-23 @ 18:37

    'Work in memory
    Dim arr() As Variant
    ReDim arr(1 To 100, 1 To 7)
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim area As Range
    Dim ruleIndex As Integer
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
                Debug.Print ws.name & " - " & area.FormatConditions.count
                ' Loop through each conditional formatting rule in the area
                For ruleIndex = 1 To area.FormatConditions.count
                    Set cf = area.FormatConditions(ruleIndex)
                    i = i + 1
                    arr(i, 1) = ws.name & Chr(0) & area.Address
                    arr(i, 2) = ws.name
                    arr(i, 3) = area.Address
                    arr(i, 4) = cf.Type
                    arr(i, 5) = cf.Formula1
                    
                    On Error Resume Next
                    If cf.Type = xlCellValue And (cf.Operator = xlBetween Or cf.Operator = xlNotBetween) Then
                        arr(i, 6) = cf.Formula2
                    End If
                    On Error GoTo 0
                    
                    arr(i, 7) = Format(Now(), "yyyy-mm-dd hh:mm")
                Next ruleIndex
            Next area
        End If
        
        'Reset the range variable for the next worksheet
    Next ws
    
    Call Array_2D_Resizer(arr, i, UBound(arr, 2))
    Call Array_2D_Bubble_Sort(arr)

    'Setup and prepare the output worksheet
    Dim wsOutput As Worksheet: Set wsOutput = wshzDocConditionalFormatting
    Dim lastUsedRow As Long
    lastUsedRow = wsOutput.Range("A9999").End(xlUp).row
    If lastUsedRow > 1 Then
        wsOutput.Range("A2:F" & lastUsedRow).ClearContents
    End If
    
    'Assign array to range
    wsOutput.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
    wsOutput.Range("A:A").EntireColumn.Hidden = True 'Do not show the SortKey
   
    MsgBox "J'ai trouvé " & i & " Conditional Formatting"
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set area = Nothing
    Set cf = Nothing
    Set rng = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing

End Sub

Sub List_Data_Validations_All() '2024-07-15 @ 06:52

    'Prepare the result worksheet (wsOutput)
    Call Erase_And_Create_Worksheet("Doc_Data_Validations")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Doc_Data_Validations")
    wsOutput.Cells(1, 1).value = "SortKey"
    wsOutput.Cells(1, 2).value = "Worksheet"
    wsOutput.Cells(1, 3).value = "CellAddress"
    wsOutput.Cells(1, 4).value = "ValidationType"
    wsOutput.Cells(1, 5).value = "Formula1"
    wsOutput.Cells(1, 6).value = "Formula2"
    wsOutput.Cells(1, 7).value = "Operator"
    wsOutput.Cells(1, 8).value = "TimeStamp"
    
    Call Make_It_As_Header(wsOutput.Range("A1:H1"))
    
    'Create the Array to store results in memory
    Dim arr() As Variant
    ReDim arr(1 To 5000, 1 To 8)
    
    ' Loop through each worksheet in the workbook
    Dim dvType As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim TimeStamp As String
    Dim x As Long: x = 1
    Dim xAnalyzed As Long
    For Each ws In ThisWorkbook.Worksheets
        'Loop through each cell in the worksheet
        For Each cell In ws.usedRange
            'Check if the cell has data validation
            xAnalyzed = xAnalyzed + 1

            On Error Resume Next
            dvType = ""
            dvType = cell.Validation.Type
            On Error GoTo 0
            
            If dvType <> "" And dvType <> "0" Then
                'Write the data validation details to the output sheet
                arr(x, 1) = ws.name & Chr(0) & cell.Address 'Sort Key
                arr(x, 2) = ws.name
                arr(x, 3) = cell.Address
                arr(x, 4) = dvType
                Select Case dvType
                    Case "2"
                        arr(x, 4) = "Min/Max"
                    Case "3"
                        arr(x, 4) = "Liste"
                    Case Else
                        arr(x, 4) = dvType
                End Select
                On Error Resume Next
                arr(x, 5) = "'" & cell.Validation.Formula1
                On Error GoTo 0
                
                On Error Resume Next
                arr(x, 6) = "'" & cell.Validation.Formula2
                On Error GoTo 0
                
                On Error Resume Next
                arr(x, 7) = "'" & cell.Validation.Operator
                On Error GoTo 0
                
                TimeStamp = Format(Now(), "dd/mm/yyyy hh:mm:ss")
                arr(x, 8) = TimeStamp

                'Increment the output row counter
                x = x + 1
            End If
        Next cell
    Next ws

    If x > 1 Then
    
        x = x - 1
        
        Call Array_2D_Resizer(arr, x, UBound(arr, 2))
        
        Call Array_2D_Bubble_Sort(arr)
        
        'Array to Worksheet
        Dim outputRow As Long: outputRow = 2
        wsOutput.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
        wsOutput.Range("A:A").EntireColumn.Hidden = True 'Do not show the sortKey
        wsOutput.columns(4).HorizontalAlignment = xlCenter
        wsOutput.columns(7).HorizontalAlignment = xlCenter
        wsOutput.columns(8).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        Dim lastUsedRow As Long
        lastUsedRow = wsOutput.Range("B99999").End(xlUp).row
        Dim j As Long, oldWorksheet As String
        oldWorksheet = wsOutput.Range("B" & lastUsedRow).value
        For j = lastUsedRow To 2 Step -1
            If wsOutput.Range("B" & j).value <> oldWorksheet Then
                wsOutput.rows(j + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                oldWorksheet = wsOutput.Range("B" & j).value
            End If
        Next j
        
        'Since we might have inserted new row, let's update the lastUsedRow
        lastUsedRow = wsOutput.Range("B99999").End(xlUp).row
        With wsOutput.Range("B2:H" & lastUsedRow)
            On Error Resume Next
            Cells.FormatConditions.delete
            On Error GoTo 0
        
            .FormatConditions.add Type:=xlExpression, Formula1:= _
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
    'Initialize the output row counter

    'AutoFit the columns for better readability
    wsOutput.columns.AutoFit
    
    'Result print setup - 2024-07-15 @ 09:22
    lastUsedRow = lastUsedRow + 2
    wsOutput.Range("B" & lastUsedRow).value = "*** " & Format(xAnalyzed, "###,##0") & _
                                    " cellules analysées dans l'application ***"
    Dim header1 As String: header1 = "Cells Data Validations"
    Dim header2 As String: header2 = "All worksheets"
    Call Simple_Print_Setup(wsOutput, wsOutput.Range("B2:H" & lastUsedRow), _
                           header1, _
                           header2, _
                           "L")
    
    MsgBox "Data validation list were created in worksheet: " & wsOutput.name
    
End Sub

Sub Erase_And_Create_Worksheet(sheetName As String)

    Dim ws As Worksheet
    Dim wsExists As Boolean

    'Check if the worksheet exists
    wsExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = sheetName Then
            wsExists = True
            Exit For
        End If
    Next ws

    'If the worksheet exists, delete it
    If wsExists Then
        Application.DisplayAlerts = False
        ws.delete
        Application.DisplayAlerts = True
    End If

    'Create a new worksheet with the specified name
    Set ws = ThisWorkbook.Worksheets.add
    ws.name = sheetName
    
    'Clean up - 2024-07-11 @ 08:27
    Set ws = Nothing
    
End Sub

Sub List_Formulas_All() '2024-06-22 @ 15:42
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    
    'Prepare existing worksheet to receive data
    Dim lastUsedRow As Long
    lastUsedRow = wshzDocFormulas.Range("A9999").End(xlUp).row 'Last used row
    If lastUsedRow > 1 Then wshzDocFormulas.Range("A2:G" & lastUsedRow).ClearContents
    
    'Create an Array to receive the formulas informations
    Dim outputArray() As Variant
    ReDim outputArray(1 To 7500, 1 To 8)
    
    'Loop through each worksheet
    Dim ws As Worksheet
    Dim name As String, usedRange As String, cellsCount As String
    For Each ws In wb.Sheets
        If ws.CodeName = "wshzDocNamedRange" Or _
            ws.CodeName = "wshzDocFormules" Then
                GoTo nextIteration
        End If
        'Save information for this worksheet
        name = ws.name
        usedRange = ws.usedRange.Address
        cellsCount = ws.usedRange.count
        'Loop through all cells in the used range
        Dim cell As Range
        Dim i As Long
        For Each cell In ws.usedRange
            'Does the cell contain a Formula
            If Left(cell.formula, 1) = "=" Then
                'Write formula information to the destination worksheet
                i = i + 1
                outputArray(i, 1) = ws.CodeName & Chr(0) & cell.Address
                outputArray(i, 2) = ws.CodeName
                outputArray(i, 3) = name
                outputArray(i, 4) = usedRange
                outputArray(i, 5) = cellsCount
                outputArray(i, 6) = cell.Address
                outputArray(i, 7) = "'=" & Mid(cell.formula, 2) 'Add ' to preserve formulas
                outputArray(i, 8) = Format(Now(), "yyyy-mm-dd hh:mm") 'Timestamp
            End If
        Next cell
nextIteration:
    Next ws
    
    Call Array_2D_Resizer(outputArray, i, UBound(outputArray, 2))
    Call Array_2D_Bubble_Sort(outputArray)
    
    'Transfer the array data to the worksheet
    wshzDocFormulas.Range("A2").Resize(UBound(outputArray, 1), UBound(outputArray, 2)).value = outputArray
    wshzDocFormulas.Range("A:A").EntireColumn.Hidden = True 'Do not show the outputArray

    MsgBox "J'ai trouvé " & Format(i, "#,##0") & " formules"
    
    'Cleaning memory - 2024-07-01 @ 09:34
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
        char = Mid(codeLine, i, 1)
        
        'Toggle inString flag if a double quote is encountered
        If char = """" Then
            inString = Not inString
        End If
        
        'If the current character is ' and we are not within a string...
        If char = "'" Then
            If Not inString Then
                commentPart = Mid(codeLine, i)
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
        commentPart = ""
    Else
        commentPart = Trim(UCase(commentPart))
    End If
    
    HandleComments = codePart & commentPart
    
End Function

Sub List_All_Shapes_Properties()

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    
    Dim r As Integer
    r = 2
    ws.Range("D" & r).value = "Type"
    ws.Range("E" & r).value = "Shape Name"
    ws.Range("F" & r).value = "ZOrder"
    ws.Range("G" & r).value = "Top"
    ws.Range("H" & r).value = "Left"
    ws.Range("I" & r).value = "Width"
    ws.Range("J" & r).value = "Height"
    
    'Loop through all shapes on the worksheet
    Dim shp As Shape
    r = 3
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
    
    'Cleaning memory - 2024-07-01 @ 09:34
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
            Debug.Print "Sheet: " & ws.name; Tab(40); "Table: " & lo.name & vbCrLf
        Next lo
    Next ws
    
    'Cleanup - 2024-07-14 @ 12:05
    Set lo = Nothing
    Set ws = Nothing
    
End Sub

Sub List_Named_Ranges_All() '2024-06-23 @ 07:40
    
    'Setup and clear the output worksheet
    Dim ws As Worksheet: Set ws = wshzDocNamedRange
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A9999").End(xlUp).row
    ws.Range("A2:I" & lastUsedRow).ClearContents
    
    'Loop through each named range in the workbook
    Dim arr() As Variant
    ReDim arr(1 To 300, 1 To 9)
    Dim i As Long
    Dim nr As name
    Dim TimeStamp As String
    For Each nr In ThisWorkbook.Names
        i = i + 1
        arr(i, 1) = UCase(nr.name) & Chr(0) & UCase(nr.RefersTo) 'Sort Key
        arr(i, 2) = nr.name
        arr(i, 3) = "'" & nr.RefersTo
        If InStr(nr.RefersTo, "#REF!") Then
            arr(i, 4) = "'#REF!"
        End If
        
        'Check if the name refers to a range
        On Error Resume Next
        Dim rng As Range
        Set rng = nr.RefersToRange
        On Error GoTo 0
        
        If Not rng Is Nothing Then
            arr(i, 5) = rng.Worksheet.name
            arr(i, 6) = rng.Address
        End If
        Set rng = Nothing
        
        arr(i, 7) = nr.Comment
        If nr.Visible = False Then
            arr(i, 8) = nr.Visible
        End If
        TimeStamp = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        arr(i, 9) = TimeStamp
    Next nr
    
    Call Array_2D_Resizer(arr, i, UBound(arr, 2))
    Call Array_2D_Bubble_Sort(arr)
    
    'Transfer the array data to the worksheet
    wshzDocNamedRange.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
    wshzDocNamedRange.Range("A:A").EntireColumn.Hidden = True 'Do not show the outputArray
    
    'Result print setup - 2024-07-14 2 07:10
    If i > 1 Then
        Dim header1 As String: header1 = "List all Named Ranges"
        Dim header2 As String: header2 = ""
        Call Simple_Print_Setup(wshzDocNamedRange, wshzDocNamedRange.Range("B2:I" & i), _
                               header1, _
                               header2, _
                               "L")
    End If
   
    MsgBox "J'ai trouvé " & i & " named ranges"
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set nr = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
End Sub

Sub Reorganize_Tests_And_Todos_Worksheet() '2024-03-02 @ 15:21

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet: Set ws = wshzDocTests_And_Todos
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A999").End(xlUp).row
    Dim rng As Range: Set rng = ws.Range("A1:E" & lastUsedRow)
    
    With ws.ListObjects("tblTests_And_Todo").Sort
        Application.EnableEvents = False
        .SortFields.clear
        .SortFields.Add2 _
            key:=Range("tblTests_And_Todo[Statut]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            key:=Range("tblTests_And_Todo[Module]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            key:=Range("tblTests_And_Todo[Priorité]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SortFields.Add2 _
            key:=Range("tblTests_And_Todo[TimeStamp]"), _
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
    
    While ws.Range("D2").value = "a"
        Set rowToMove = tbl.ListRows(1).Range
        lastRow = tbl.ListRows.count
        rowToMove.Cut Destination:=tbl.DataBodyRange.rows(lastRow + 1)
        tbl.ListRows(1).delete
    Wend

    ws.Calculate
    
    Application.EnableEvents = False
    
    Application.ScreenUpdating = True
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rng = Nothing
    Set rowToMove = Nothing
    Set tbl = Nothing
    Set ws = Nothing
    
End Sub

Sub Search_Every_Lines_Of_Code(arr As Variant, search1 As String, search2 As String, search3 As String)

    'Declare arr() to keep results in memory
    Dim arrResult() As Variant
    ReDim arrResult(1 To 1000, 1 To 7)

    Dim posProcedure As Long, posFunction As Long
    Dim saveLineOfCode As String, trimmedLineOfCode As String, procedureName As String
    Dim TimeStamp As String
    Dim x As Long, xr As Long
    For x = LBound(arr, 1) To UBound(arr, 1)
        trimmedLineOfCode = arr(x, 4)
        saveLineOfCode = trimmedLineOfCode
        
        'Handle comments (second parameter is either Remove or Uppercase)
        If InStr(1, trimmedLineOfCode, "'") <> 0 Then
            trimmedLineOfCode = HandleComments(trimmedLineOfCode, "U")
        End If
        
        If trimmedLineOfCode <> "" Then
            'Is this a procedure (Sub) declaration line ?
            If InStr(trimmedLineOfCode, "Sub ") <> 0 Then
                If InStr(trimmedLineOfCode, "End Sub") = 0 And _
                    InStr(trimmedLineOfCode, "Sub = ") = 0 And _
                    InStr(trimmedLineOfCode, "Sub As ") = 0 And _
                    InStr(trimmedLineOfCode, "Exit Sub") = 0 Then
                        procedureName = Mid(saveLineOfCode, InStr(trimmedLineOfCode, "Sub "))
                End If
            End If
            
            If InStr(trimmedLineOfCode, "End Sub") = 1 Then
                procedureName = ""
            End If

            'Is this a function declaration line ?
            If InStr(trimmedLineOfCode, "Function ") <> 0 Then
                If InStr(trimmedLineOfCode, "End Function") = 0 And _
                    InStr(trimmedLineOfCode, "Function = ") = 0 And _
                    InStr(trimmedLineOfCode, "Function As ") = 0 And _
                    InStr(trimmedLineOfCode, "Exit Function") = 0 Then
                        procedureName = Mid(saveLineOfCode, InStr(trimmedLineOfCode, "Function "))
                End If
            End If
            
            If InStr(trimmedLineOfCode, "End Function") = 1 Then
                procedureName = ""
            End If
            
            'Do we find the search1 or search2 or sreach3 strings in this line of code ?
            If (search1 <> "" And InStr(trimmedLineOfCode, search1) <> 0) Or _
                (search2 <> "" And InStr(trimmedLineOfCode, search2) <> 0) Or _
                (search3 <> "" And InStr(trimmedLineOfCode, search3) <> 0) Then
                'Found an occurence
                xr = xr + 1
                arrResult(xr, 2) = arr(x, 1) 'oType
                arrResult(xr, 3) = arr(x, 2) 'oName
                arrResult(xr, 4) = arr(x, 3) 'LineNum
                arrResult(xr, 5) = procedureName
                arrResult(xr, 6) = "'" & saveLineOfCode
                TimeStamp = Format(Now(), "dd/mm/yyyy hh:mm:ss")
                arrResult(xr, 7) = TimeStamp
                arrResult(xr, 1) = UCase(arr(x, 1)) & Chr(0) & UCase(arr(x, 2)) & Chr(0) & Format(arr(x, 3), "0000") & Chr(0) & procedureName 'Future sort key
            End If
        End If
    Next x

    'Prepare the result worksheet
    Call Erase_And_Create_Worksheet("Doc_Search_Utility_Results")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Doc_Search_Utility_Results")
    wsOutput.Range("A1").value = "SortKey"
    wsOutput.Range("B1").value = "Type"
    wsOutput.Range("C1").value = "ModuleName"
    wsOutput.Range("D1").value = "LineNo"
    wsOutput.Range("E1").value = "ProcedureName"
    wsOutput.Range("F1").value = "Code"
    wsOutput.Range("G1").value = "TimeStamp"
    
    Call Make_It_As_Header(wsOutput.Range("A1:G1"))
    
    'Is there anything to show ?
    If xr > 0 Then
    
        'Data starts at row 2
        Dim r As Long: r = 2

        Call Array_2D_Resizer(arrResult, xr, UBound(arrResult, 2))
        
        'Sort the 2D array based on column 1
        Call Array_2D_Bubble_Sort(arrResult)
    
        'Transfer the array to the worksheet
        wsOutput.Range("A2").Resize(UBound(arrResult, 1), UBound(arrResult, 2)).value = arrResult
        wsOutput.Range("A:A").EntireColumn.Hidden = True 'Do not show the sortKey
        wsOutput.columns(4).HorizontalAlignment = xlCenter
        wsOutput.columns(7).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        Dim lastUsedRow As Long
        lastUsedRow = wsOutput.Range("B9999").End(xlUp).row
        Dim j As Long, oldProcedure As String
        oldProcedure = wsOutput.Range("C" & lastUsedRow).value & wsOutput.Range("E" & lastUsedRow).value
        For j = lastUsedRow To 2 Step -1
            If wsOutput.Range("C" & j).value & wsOutput.Range("E" & j).value <> oldProcedure Then
                wsOutput.rows(j + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                oldProcedure = wsOutput.Range("C" & j).value & wsOutput.Range("E" & j).value
            End If
        Next j
        
        'Since we might have inserted new row, let's update the lastUsedRow
        lastUsedRow = wsOutput.Range("B9999").End(xlUp).row
        With wsOutput.Range("B2:G" & lastUsedRow)
            On Error Resume Next
            Cells.FormatConditions.delete
            On Error GoTo 0
        
            .FormatConditions.add Type:=xlExpression, Formula1:= _
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
    wsOutput.Range("B" & lastUsedRow).value = "*** " & Format(x, "###,##0") & " lignes de code dans l'application ***"
    Dim header1 As String: header1 = "Search Utility Results"
    Dim header2 As String
    header2 = "Searched strings '" & search1 & "'"
    If search2 <> "" Then header2 = header2 & " '" & search2 & "'"
    If search3 <> "" Then header2 = header2 & " '" & search3 & "'"
    Call Simple_Print_Setup(wsOutput, wsOutput.Range("B2:G" & lastUsedRow), _
                           header1, _
                           header2, _
                           "L")
    
    'Display the final message
    If xr Then
        MsgBox "J'ai trouvé " & xr & " lignes avec les chaines '" & search1 & "'" & vbNewLine & _
                vbNewLine & "après avoir analysé un total de " & _
                Format(x, "#,##0") & " lignes de code"
    Else
        MsgBox "Je n'ai trouvé aucune occurences avec les chaines '" & search1 & "'" & vbNewLine & _
                vbNewLine & "après avoir analysé un total de " & _
                Format(x, "#,##0") & " lignes de code"
    End If
    
    'Clean up - 2024-07-13 @ 08:02
    Set wsOutput = Nothing
    
End Sub

Sub List_All_Macros_Used_With_Objects() '2024-07-25 @ 11:17
    
    'Prepare the result worksheet
    Call Erase_And_Create_Worksheet("Doc_All_Macros_Used_With_Object")

    Dim wsOutputSheet As Worksheet
    Set wsOutputSheet = ThisWorkbook.Worksheets("Doc_All_Macros_Used_With_Object")
    
    wsOutputSheet.Cells(1, 1).value = "Worksheet"
    wsOutputSheet.Cells(1, 2).value = "Object Type"
    wsOutputSheet.Cells(1, 3).value = "Object Name"
    wsOutputSheet.Cells(1, 4).value = "Macro Name"
    
    Call Make_It_As_Header(wsOutputSheet.Range("A1:D1"))

    Dim outputRow As Long
    outputRow = 2 'Start writing from the second row

    'Iterate through each worksheet in the workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim macroName As String

    For Each ws In ThisWorkbook.Worksheets
        'Skip the output sheet to avoid listing its own shapes
        If ws.name <> "Doc_All_Macros_Used_With_Object" Then
            'Check for macros assigned to shapes
            For Each shp In ws.Shapes
                On Error Resume Next
                macroName = shp.OnAction
                On Error GoTo 0
                If macroName <> "" Then
                    wsOutputSheet.Cells(outputRow, 1).value = ws.name
                    wsOutputSheet.Cells(outputRow, 2).value = "Shape"
                    wsOutputSheet.Cells(outputRow, 3).value = shp.name
                    wsOutputSheet.Cells(outputRow, 4).value = macroName
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
                If macroName <> "" Then
                    wsOutputSheet.Cells(outputRow, 1).value = ws.name
                    wsOutputSheet.Cells(outputRow, 2).value = "ActiveX Control"
                    wsOutputSheet.Cells(outputRow, 3).value = obj.name
                    wsOutputSheet.Cells(outputRow, 4).value = macroName
                    outputRow = outputRow + 1
                End If
            Next obj
        End If
    Next ws

'    'Part # 2 - Analyze the code itself
'
'    'Iterate through the VBA project to identify all callable subs/functions
'    Dim lineNum As Long
'    Dim procName As String
'    Dim procType As VBIDE.vbext_ProcKind
'    Dim lineText As String
'    Dim oType As String
'
'    Dim vbComp As Object
'    For Each vbComp In ThisWorkbook.VBProject.VBComponents
'        Select Case vbComp.Type
'            Case 1
'                oType = "1_Module"
'            Case 2
'                oType = "2_Class"
'            Case 3
'                oType = "3_userForm"
'            Case Else
'                oType = oType & "_????"
'        End Select
'
'       'Get the code module for the component
'        Dim vbMod As Object: Set vbMod = vbComp.CodeModule
'        lineNum = 1
'
'        'Loop through all lines in the code module to save all the lines in memory
'        For lineNum = 1 To vbMod.CountOfLines
'            lineText = Trim(vbMod.Lines(lineNum, 1))
'            If Left(lineText, 1) <> "'" And Left(lineText, 3) <> "Rem" Then
'                procName = vbMod.ProcOfLine(lineNum, procType)
'                wsOutputSheet.Cells(outputRow, 1).value = vbComp.name
'                wsOutputSheet.Cells(outputRow, 2).value = "Procedure"
'                wsOutputSheet.Cells(outputRow, 3).value = procName
'                wsOutputSheet.Cells(outputRow, 4).value = "Called in code"
'                outputRow = outputRow + 1
'            End If
'        Next lineNum
'    Next vbComp

    'Autofit columns for better readability
    wsOutputSheet.columns("A:D").AutoFit
    outputRow = outputRow - 1 'Did not use the last line
    
    'Sort the results, based on column 1, 2, 3 & 4
    If outputRow > 2 Then
        'Sort the data by columns 1, 2, 3, and 4
        With wsOutputSheet.Sort
            .SortFields.clear
            .SortFields.add key:=wsOutputSheet.Range("A2:A" & outputRow - 1), Order:=xlAscending
            .SortFields.add key:=wsOutputSheet.Range("B2:B" & outputRow - 1), Order:=xlAscending
            .SortFields.add key:=wsOutputSheet.Range("C2:C" & outputRow - 1), Order:=xlAscending
            .SortFields.add key:=wsOutputSheet.Range("D2:D" & outputRow - 1), Order:=xlAscending
            .SetRange wsOutputSheet.Range("A1:D" & outputRow - 1)
            .Header = xlYes
            .Apply
        End With
        
        'Automatically insert a empty line between worksheet
        Dim i As Long
        For i = outputRow To 2 Step -1
            If wsOutputSheet.Cells(i, 1).value <> wsOutputSheet.Cells(i + 1, 1).value Then
                wsOutputSheet.rows(i + 1).Insert Shift:=xlDown
            End If
        Next i
    End If
    
    'Set conditional formatting for the worksheet (alternate colors)
    outputRow = wsOutputSheet.Range("A9999").End(xlUp).row
    Dim rngArea As Range: Set rngArea = wsOutputSheet.Range("A2:D" & outputRow)
    Debug.Print rngArea.Address
    Call Apply_Conditional_Formatting_Alternate(rngArea, 1, True) 'There are blankrows to account for
    
    outputRow = wsOutputSheet.Range("A9999").End(xlUp).row
    Dim rngToPrint As Range: Set rngToPrint = wsOutputSheet.Range("A2:D" & outputRow)
    Dim header1 As String: header1 = "Liste des macros associées à des contrôles"
    Dim header2 As String: header2 = ThisWorkbook.name
    Call Simple_Print_Setup(wsOutputSheet, rngToPrint, header1, header2, "P")
    
    MsgBox "La liste des macros assignées à des contrôles est dans " & _
                vbNewLine & vbNewLine & "la feuille 'Doc_All_Macros_Used_With_Object'.", vbInformation
                
    'Cleanup - 2024-07-25 @ 11:07
    Set obj = Nothing
    Set wsOutputSheet = Nothing
    Set rngArea = Nothing
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub List_Subs_And_Functions_All() '2024-06-22 @ 10:41
    
    Dim posProcedure As Integer, posExitProcedure As Integer
    Dim posFonction As Integer, posExitFonction As Integer
    Dim posSpace As Integer, posREM As Integer, posParam As Integer
    Dim scope As String, sType As String
    
    'Loop through all VBcomponents (modules, class and forms) in the active workbook
    Dim vbComp As Object
    Dim oType As String
    Dim arr() As Variant
    ReDim arr(1 To 500, 1 To 10)
    Dim trimmedLineOfCode As String, savedLineOfCode As String, remarks As String, params As String
    Dim lineNum As Long, lread As Long
    Dim i As Integer

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
            Dim vbCodeMod As Object: Set vbCodeMod = vbComp.CodeModule
            'Loop through all lines in the code module
            For lineNum = 1 To vbCodeMod.CountOfLines
                lread = lread + 1
                'Check if the line contains 'Sub' or 'Function' without beeing a Remark line
                savedLineOfCode = Trim(vbCodeMod.Lines(lineNum, 1))
                trimmedLineOfCode = Trim(vbCodeMod.Lines(lineNum, 1))
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
                    arr(i, 3) = vbComp.name
                    arr(i, 4) = lineNum
                    'Goback to savedLineOfCode
                    trimmedLineOfCode = Trim(vbCodeMod.Lines(lineNum, 1))
                    posREM = InStr(trimmedLineOfCode, ") '")
                    If posREM > 0 Then
                        remarks = Trim(Mid(trimmedLineOfCode, posREM + 2))
                        trimmedLineOfCode = Trim(Left(trimmedLineOfCode, posREM))
                    End If
                    posParam = InStr(trimmedLineOfCode, "(")
                    If posParam > 0 Then
                        params = Trim(Mid(trimmedLineOfCode, posParam))
                        trimmedLineOfCode = Trim(Left(trimmedLineOfCode, posParam - 1))
                    End If
                    
                    If InStr(trimmedLineOfCode, "Sub ") > 1 Or InStr(trimmedLineOfCode, "Function ") > 1 Then
                        posSpace = InStr(trimmedLineOfCode, " ")
                        scope = Left(trimmedLineOfCode, posSpace - 1)
                        trimmedLineOfCode = Trim(Mid(trimmedLineOfCode, posSpace + 1))
                    Else
                        scope = ""
                    End If
                    arr(i, 5) = scope
                    If InStr(trimmedLineOfCode, "Sub ") = 1 Then
                        sType = "Sub"
                        trimmedLineOfCode = Trim(Mid(trimmedLineOfCode, 5))
                    Else
                        If InStr(trimmedLineOfCode, "Function ") = 1 Then
                            sType = "Function"
                            trimmedLineOfCode = Trim(Mid(trimmedLineOfCode, 10))
                        End If
                    End If
                    arr(i, 6) = sType
                    arr(i, 7) = trimmedLineOfCode
                    arr(i, 1) = UCase(oType) & Chr(0) & UCase(vbComp.name) & Chr(0) & UCase(trimmedLineOfCode) 'Future sort key
                    If params <> "" Then arr(i, 8) = params
                    If remarks <> "" Then arr(i, 9) = remarks
                    arr(i, 10) = Format(Now(), "yyyy-mm-dd hh:mm")
                    params = ""
                    remarks = ""
                End If
            Next lineNum
        End If
    Next vbComp
    
    'Prepare the output worksheet
    Dim lastUsedRow As Long
    lastUsedRow = wshzDocSubsAndFunctions.Range("A9999").End(xlUp).row 'Last Used Row
    wshzDocSubsAndFunctions.Range("A2:I" & lastUsedRow).ClearContents

    Call Array_2D_Resizer(arr, i, UBound(arr, 2))
    
    'Sort the 2D array based on column 1
    Call Array_2D_Bubble_Sort(arr)
    
    'Transfer the array to the worksheet
    wshzDocSubsAndFunctions.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
    wshzDocSubsAndFunctions.Range("A:A").EntireColumn.Hidden = True 'Do not show the sortKey
    
    MsgBox "J'ai trouvé " & i & " lignes Sub or Function" & vbNewLine & _
                vbNewLine & "après avoir analysé un total de " & _
                Format(lread, "#,##0") & " Lignes de code"
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set vbComp = Nothing
    Set vbCodeMod = Nothing
    
End Sub

Sub Test_Array_To_Range() '2024-03-18 @ 17:34

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(2)
    
    Dim arr() As Variant
    ReDim arr(1 To 1000, 1 To 20)
    
    Dim i As Integer, j As Integer
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            arr(i, j) = "i = " & i & " and j = " & j & " - *********"
        Next j
    Next i
    
    ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Sub TestArray_2D_Resizer()
    Dim originalArray() As Variant
    Dim i As Long, j As Long
    
    ' Dimension the original array to a fixed size (e.g., 10 rows and 5 columns)
    ReDim originalArray(1 To 10, 1 To 5)
    
    ' Fill the original array with some example data
    For i = 1 To 10
        For j = 1 To 5
            originalArray(i, j) = "R" & i & "C" & j
        Next j
    Next i
    
    ' Output the original array to the immediate window
    Debug.Print "Original Array:"
    For i = 1 To 10
        For j = 1 To 5
            Debug.Print originalArray(i, j);
        Next j
        Debug.Print
    Next i
    
    ' Trim the array to 6 rows and 3 columns
    Call Array_2D_Resizer(originalArray, 6, 3)
    
    ' Output the trimmed array to the immediate window
    Debug.Print "Trimmed Array:"
    For i = 1 To 6
        For j = 1 To 3
            Debug.Print originalArray(i, j);
        Next j
        Debug.Print
    Next i
End Sub

Sub TestGetQuarterDates()

    Dim fiscalYearStartMonth As Integer
    Dim fiscalYear As Integer
    Dim result As String
    
    'Set the fiscal year start month (e.g., April is 4)
    fiscalYearStartMonth = 8
    
    'Set the fiscal year
    fiscalYear = 2024
    
    'Get the quarter dates
    result = GetQuarterDates(fiscalYearStartMonth, fiscalYear)
    
    'Display the result
    MsgBox result
    
End Sub

Sub Toggle_A1_R1C1_Reference()

    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If

End Sub

Sub List_Worksheets_From_Current_Workbook_All() '2024-07-24 @ 10:14
    
    Call Erase_And_Create_Worksheet("Feuilles_du_Classeur")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Feuilles_du_Classeur")
    wsOutput.Range("A1").value = "Feuille"
    wsOutput.Range("B1").value = "CodeName"
    wsOutput.Range("C1").value = "TimeStamp"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"))

    'Loop through all worksheets in the active workbook
    Dim arr() As Variant
    ReDim arr(1 To 100, 1 To 3)
    
    Dim ws As Worksheet
    Dim TimeStamp As String
    Dim i As Long
    For Each ws In ThisWorkbook.Sheets
        i = i + 1
        arr(i, 1) = ws.name
        arr(i, 2) = ws.CodeName
        TimeStamp = Format(Now(), "dd-mm-yyyy hh:mm:ss")
        arr(i, 3) = TimeStamp
    Next ws
    
    Call Array_2D_Resizer(arr, i, UBound(arr, 2))
    
    Call Array_2D_Bubble_Sort(arr)
    
    Dim f As Integer
    For i = 1 To UBound(arr, 1)
        wsOutput.Cells(i + 1, 1) = arr(i, 1)
        wsOutput.Cells(i + 1, 2) = arr(i, 2)
        wsOutput.Cells(i + 1, 3) = arr(i, 3)
        f = f + 1
    Next i
    
    wsOutput.columns.AutoFit
    
   'Result print setup - 2024-07-20 @ 14:31
    Dim lastUsedRow As Long
    lastUsedRow = i + 2
    wsOutput.Range("A" & lastUsedRow).value = "*** " & Format(f, "###,##0") & _
                                    " feuilles pour le workbook '" & ThisWorkbook.name & "' ***"
    
    lastUsedRow = wsOutput.Range("A9999").End(xlUp).row
    Dim rngToPrint As Range: Set rngToPrint = wsOutput.Range("A2:C" & lastUsedRow)
    Dim header1 As String: header1 = "Liste des feuilles d'un classeur"
    Dim header2 As String: header2 = ThisWorkbook.name
    Call Simple_Print_Setup(wsOutput, rngToPrint, header1, header2, "P")
    
    ThisWorkbook.Worksheets("Feuilles_du_Classeur").Activate
    
    'Cleanup - 2024-07-14 @ 07:03
    Set rngToPrint = Nothing
    Set wsOutput = Nothing

End Sub


