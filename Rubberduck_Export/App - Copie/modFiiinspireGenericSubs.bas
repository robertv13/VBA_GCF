Attribute VB_Name = "modFiiinspireGenericSubs"
Option Explicit

'YouTube Channel: fiiinspires
'https://www.youtube.com/c/fiiinspires
'Please SUBSCRIBE
'contains reusable/utility functions and sub-routines for this and other projects

Enum DateComponent
    dayComponent = 1
    monthComponent = 2
    yearComponent = 3
End Enum

Enum DatesCompare
    Lesser = -1
    Equal = 0
    Greater = 1
    LessOrEqual = 10
    GreaterOrEqual = 11
End Enum

Function compareDates(date1 As Date, _
                      Optional date2 As Date = 0, _
                      Optional compareType As DatesCompare = Greater) As Boolean
    If date2 = 0 Then date2 = Date
    Select Case compareType
    Case -1: If date1 < date2 Then compareDates = True
    Case 0: If date1 = date2 Then compareDates = True
    Case 1: If date1 > date2 Then compareDates = True
    Case 10: If date1 <= date2 Then compareDates = True
    Case 11: If date1 >= date2 Then compareDates = True
    End Select
End Function

Function criticalDateMsg(Optional msg As String = vbNullString, _
                         Optional title As String = vbNullString) As Boolean
    If msg = vbNullString Then msg = "Error in dating"
    If title = vbNullString Then title = "Check date"
    criticalDateMsg = MsgBox(msg, vbOKOnly + vbCritical, title)
End Function

Function asValidDateCbo(cboDay As MSForms.ComboBox, _
                        cboMonth As MSForms.ComboBox, _
                        cboYear As MSForms.ComboBox, _
                        Optional strict As Boolean = True) As Date
    'make a date if it is valid
    Dim textDateStr As String
    Dim monthInt As Integer
    If isValidDateCbo(cboDay, cboMonth, cboYear, strict) Then
        If strict Then
            monthInt = IIf(IsNumeric(cboMonth.Value), cboMonth.Value, _
                           monthToNum(cboMonth.Value))
            asValidDateCbo = DateSerial(cboYear, monthInt, cboDay)
        Else
            textDateStr = cboYear & " " & cboMonth & " " & cboDay
            asValidDateCbo = DateValue(textDateStr)
        End If
    End If
    'by default it returns 12:00:00 AM which is a date. _
    'x = DateValue("12:00:00 AM")
    'year(x), month(x), day(x), Format(x, "yyyy-mm-dd")
End Function

Function isValidDateCbo(cboDay As MSForms.ComboBox, _
                        cboMonth As MSForms.ComboBox, _
                        cboYear As MSForms.ComboBox, _
                        Optional strict As Boolean = True) As Boolean
    'It attempts to make a date and if successful it returns true else false
    Dim textDateStr As String
    Dim monthInt As Integer
    If strict Then
        monthInt = IIf(IsNumeric(cboMonth.Value), cboMonth.Value, _
                       monthToNum(cboMonth.Value))
        If cboDay = vbNullString Or cboMonth = vbNullString Or cboYear = _
           vbNullString Then isValidDateCbo = False: Exit Function
        If Not isOffDateserial(cboYear.Value, monthInt, cboDay.Value) Then _
           isValidDateCbo = True
    ElseIf Not strict Then
        textDateStr = cboYear & " " & cboMonth & " " & cboYear
        If IsDate(textDateStr) Then isValidDateCbo = True
    End If
End Function

Function isOffDateserial(iYear As Integer, iMonth As Integer, _
                         iDay As Integer) As Boolean
    'Determine if specified date is off in the case of Feb, Apr, Jun, Sep, Nov
    Dim entered As Date
    Dim compared As Date
    entered = DateSerial(iYear, iMonth, iDay)
    compared = DateSerial(iYear, iMonth, 1)
    If Month(entered) <> Month(compared) Then isOffDateserial = True
End Function

Function monthToNum(sMonthName As String) As Integer
    'note that the date string should make sense to be able to compute the _
    'datevalue e.g. "Jan"
    monthToNum = Month(DateValue("01 " & sMonthName & " 2023"))
End Function

Sub makeDateCbo(datePart As DateComponent, _
                cbo As MSForms.ComboBox, _
                Optional initValue As Boolean = False, _
                Optional initPos As Boolean = False, _
                Optional listStyle As Boolean = True, _
                Optional startYear As Integer = 0, _
                Optional numYears As Integer = 5)
    
    'Adds days or months or years to a combo-box (cbo)
    Dim i As Integer
    Dim endYear As Integer
    
    If startYear <= 0 Then startYear = Year(Date)
    endYear = startYear + numYears - 1
    
    Select Case datePart
    Case 1
        For i = 1 To 31: cbo.AddItem i: Next i
        If initValue Then cbo.Value = Day(Date)
        If initPos Then cbo.ListIndex = 0
    Case 2
        For i = 1 To 12: cbo.AddItem MonthName(i, True): Next i
        If initValue Then cbo.Value = MonthName(Month(Date), True)
        If initPos Then cbo.ListIndex = 0
    Case 3
        For i = startYear To endYear: cbo.AddItem i: Next i
        If initValue Then cbo.Value = Year(Date)
        If initPos Then cbo.ListIndex = 0
    Case Else
        Exit Sub
    End Select
    cbo.Style = IIf(listStyle, 2, 0)
End Sub

Sub makeDateCboAll(cboDay As MSForms.ComboBox, _
                   cboMonth As MSForms.ComboBox, _
                   cboYear As MSForms.ComboBox, _
                   Optional initValue As Boolean = False, _
                   Optional initPos As Boolean = False, _
                   Optional listStyle As Boolean = True, _
                   Optional startYear As Integer = 0, _
                   Optional numYears As Integer = 5)
    
    'Uses makeDateCbo to get all date components
    Call makeDateCbo(1, cboDay, initValue, initPos, listStyle, startYear, numYears)
    Call makeDateCbo(2, cboMonth, initValue, initPos, listStyle, startYear, numYears)
    Call makeDateCbo(3, cboYear, initValue, initPos, listStyle, startYear, numYears)
End Sub

Sub writeToSheet(wkSheet As Worksheet, ParamArray Args() As Variant)
    
    'Writes a row of a worksheet
    Dim nextRow As Long
    With wkSheet
        nextRow = .Range("A1").CurrentRegion.Rows.Count + 1
        .Range(.Cells(nextRow, 1), .Cells(nextRow, UBound(Args) + 1)) = Args
    End With
End Sub

Function turnRangeIntoColtn(rng As Range, _
                            Optional onlyUnique As Boolean = True, _
                            Optional removeEmpty As Boolean = True) As Collection

    Dim Cell As Range
    Dim coltn As Collection
    Set coltn = New Collection
    
    On Error Resume Next
    Select Case removeEmpty
    Case True
        For Each Cell In rng
            If Not IsEmpty(Cell) And onlyUnique Then coltn.Add Cell, CStr(Cell)
            If Not IsEmpty(Cell) And Not onlyUnique Then coltn.Add Cell
        Next
    Case False
        For Each Cell In rng
            If onlyUnique Then
                coltn.Add Cell, CStr(Cell)
            Else
                coltn.Add Cell
            End If
        Next
    End Select
    On Error GoTo 0
    
    Set turnRangeIntoColtn = coltn
End Function

Sub addRangeToCbo(rng As Range, _
                  cbo As MSForms.ComboBox, _
                  Optional clearCbo As Boolean = True, _
                  Optional initPos As Integer = -1, _
                  Optional listStyle As Boolean = True, _
                  Optional useUnique As Boolean = True, _
                  Optional removeEmpty As Boolean = True)
    
    'Adds values in a range as combo-box item
    
    Dim i As Long
    Dim coltn As Collection
    Set coltn = turnRangeIntoColtn(rng, useUnique, removeEmpty)
    
    If clearCbo Then cbo.Clear
    For i = 1 To coltn.Count
        cbo.AddItem coltn(i)
    Next i
    
    cbo.Enabled = True                           'Index and style might not work for disabled cbo
    If initPos < -1 Or initPos >= coltn.Count Then initPos = -1
    cbo.ListIndex = initPos
    cbo.Style = IIf(listStyle, 2, 0)
End Sub

Sub makeAddRangeToCbo(wkSheet As Worksheet, _
                      colNum As Integer, _
                      cbo As MSForms.ComboBox, _
                      Optional clearCbo As Boolean = True, _
                      Optional initPos As Integer = -1, _
                      Optional listStyle As Boolean = True, _
                      Optional useUnique As Boolean = True, _
                      Optional removeEmpty As Boolean = True, _
                      Optional header As Boolean = True)
    
    'Uses addRangeToCbo but first creates the range that has the data
    
    Dim lastRow As Long
    Dim firstRow As Long
    Dim rng As Range
    
    firstRow = IIf(header, 2, 1)
    
    With wkSheet
        lastRow = .Cells(Rows.Count, colNum).End(xlUp).Row
        Set rng = .Range(.Cells(firstRow, colNum), .Cells(lastRow, colNum))
    End With

    If header And lastRow = 1 Then Exit Sub
    Call addRangeToCbo(rng, cbo, clearCbo, initPos, listStyle, useUnique, removeEmpty)
End Sub

Sub CopyWorksheetDataToListBox(wkSheet As Worksheet, _
                               lstBox As MSForms.ListBox, _
                               header As Boolean)
    
    Dim dataArray() As Variant
    Dim lastRow As Long, headerOffset As Long

    'Clear existing data in the ListBox
    lstBox.Clear
    
    'Assuming data starts in column A and ends in column C (adjust as needed)
    lastRow = wkSheet.Cells(wkSheet.Rows.Count, "A").End(xlUp).Row
    
    If header = True Then
        headerOffset = 1                         'With header in the worksheet
    Else
        headerOffset = 0                         'No header in the worksheet
    End If
    
    'Resize the array to match the size of the data range
    ReDim dataArray(1 To lastRow - headerOffset, 1 To wkSheet.UsedRange.Columns.Count)
    
    'Populate the array with data from the worksheet
    Dim i As Long
    For i = 1 To lastRow - headerOffset
        Dim j As Long
        For j = 1 To wkSheet.UsedRange.Columns.Count
            dataArray(i, j) = wkSheet.Cells(i + headerOffset, j).Value
        Next j
    Next i
    
    'Transfer data from the array to the ListBox
    With lstBox
        .ColumnCount = wkSheet.UsedRange.Columns.Count
        .List = dataArray
    End With

End Sub

Sub clearCtrs(ctrSet As Object, tagged As String)
    
    Dim ctl As Control

    For Each ctl In ctrSet.Controls
        If ctl.Tag = tagged Then
            If TypeOf ctl Is TextBox Then
                ctl.Value = ""
            ElseIf TypeOf ctl Is ComboBox Then
                ctl.Value = ""
            End If
        End If
    Next ctl

End Sub

Function makeRangeContains(wkSheet As Worksheet, _
                           colNum As Integer, _
                           Optional rowNumStart As Long = 1, _
                           Optional rngContains As String = vbNullString, _
                           Optional offsetColBy As Integer = 0, _
                           Optional offsetContains As String = vbNullString) As Range
    'Make range with data of selected column, colNum
    'default rngContains (vbnullstring) means don't use any criteria
    'offsetColBy- another column to test offsetContains against
    Dim rngUnion As Range
    Dim rng As Range
    Dim Cell As Range
    Dim rowLen As Long
    Dim i As Long
    With wkSheet
        'rowLen = .Cells(Rows.Count, 1).End(xlUp).Row
        rowLen = .Cells(rowNumStart, colNum).CurrentRegion.Rows.Count
        Set rng = .Range(.Cells(rowNumStart, colNum), .Cells(rowLen, colNum))
    End With
    If rngContains = vbNullString Then
        Set rngUnion = rng
    Else
        offsetColBy = IIf(offsetContains = vbNullString, 0, offsetColBy) 'if you offset, you must also set the criteria else offset is ignored
        offsetContains = IIf(offsetContains = vbNullString Or offsetColBy = 0, _
                             rngContains, offsetContains)
        
        For Each Cell In rng
            If CStr(Cell.Value) = rngContains And CStr(Cell.Offset(0, offsetColBy).Value) = _
               offsetContains Then
                i = i + 1
                If i = 1 Then Set rngUnion = Cell
                Set rngUnion = Union(rngUnion, Cell)
            End If
        Next
    End If
    
    If Not rngUnion Is Nothing Then Set makeRangeContains = rngUnion
End Function

Sub setCtrToNullString(ParamArray ctrSet() As Variant)
    Dim i As Integer
    For i = LBound(ctrSet) To UBound(ctrSet)
        ctrSet(i).Value = vbNullString
    Next i
End Sub

Function requireFieldMsg() As Boolean
    'message to show if info is missing
    Dim msg As String
    Dim title As String
    msg = "Fields marked with * are required."
    title = "Missing Information"
    If MsgBox(msg, Buttons:=vbOKOnly + vbCritical, title:=title) = vbOK Then requireFieldMsg = True
End Function

Function requireFieldExists(ctrSet As Object, _
                            Optional tagAs As String = "Required") As Boolean
    'Checks if required field has no data
    Dim ctr As Control
    For Each ctr In ctrSet.Controls
        If InStr(ctr.Tag, tagAs) Then
            If Len(Trim(ctr)) = 0 Then
                requireFieldExists = True
                ctr.SetFocus
                Exit For
            End If
        End If
    Next
End Function


