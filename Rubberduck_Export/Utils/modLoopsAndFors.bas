Attribute VB_Name = "modLoopsAndFors"
Option Explicit

'For Each WS In Worksheets
'    'Update or do something here
'Next

Public Sub RemoveFormatting()
    Dim c As Range

    For Each c In Range("A:A")
        c.ClearFormats
    Next c

End Sub

Sub ListSheets()
    Dim WS As Worksheet
    Dim x As Integer

    x = 1

    Sheets("Sheet1").Range("A:A").Clear

    For Each WS In Worksheets
        Sheets("Sheet1").Cells(x, 1).Select
        ActiveSheet.Hyperlinks.Add _
        Anchor:=Selection, Address:="", SubAddress:= _
        WS.Name & "!A1", TextToDisplay:=WS.Name
        x = x + 1
    Next WS

End Sub

Function Sheet_Exists(WorkSheet_Name As String) As Boolean

    Dim Work_sheet As Worksheet
    Sheet_Exists = False
    For Each Work_sheet In ThisWorkbook.Worksheets
        If Work_sheet.Name = WorkSheet_Name Then
            Sheet_Exists = True
        End If
    Next

End Function

Sub DeleteAllShapes()
    'Activate sheet to delete autoshapes
    Sheet1.Activate

    Dim GetShape As Shape
    For Each GetShape In ActiveSheet.Shapes
        GetShape.Delete
    Next

End Sub

Sub DeleteNamedRanges()
    
    Dim MyName As Name
    For Each MyName In Names
        ActiveWorkbook.Names(MyName.Name).Delete
    Next

End Sub

Sub Highlight_Duplicates(Values As Range)
    
    Dim Cell
    For Each Cell In Values
        If WorksheetFunction.CountIf(Values, Cell.Value) > 1 Then
            Cell.Interior.ColorIndex = 6
        End If
    Next Cell

End Sub

'Sub ListSheets()
'    Dim WS As Worksheet
'    Dim x As Integer
'
'    x = 1
'
'    Sheets("Sheet1").Range("A:A").Clear
'
'    For Each WS In Worksheets
'        Sheets("Sheet1").Cells(x, 1) = WS.Name
'        x = x + 1
'    Next WS
'
'End Sub

Public Sub LoopColumn()
    Dim c As Range
    For Each c In Range("A:A")
        If c.Value = "FindMe" Then
            MsgBox "FindMe found at " & c.Address
        End If
    Next

End Sub

Public Sub LoopRow()
    Dim c As Range
    For Each c In Range("1:1")
        If c.Value = "FindMe" Then
            MsgBox "FindMe found at " & c.Address
        End If
    Next c

End Sub

Sub LoopThroughString()
    Dim Counter As Integer
    Dim MyString As String
    MyString = "AutomateExcel"                   'define string

    For Counter = 1 To Len(MyString)
        'Do something to each character in string here we'll msgbox each character
        MsgBox Mid(MyString, Counter, 1)
    Next

End Sub

Sub ZoomAll()
    Dim WS As Worksheet
    For Each WS In Worksheets
        WS.Activate
        ActiveWindow.Zoom = 50
    Next
End Sub

Sub Sort_Sheets()
    Dim Sort_Mode_Descending As Boolean
    Dim No_of_Sheets As Integer
    Dim Outer_Loop As Integer
    Dim Inner_Loop As Integer

    No_of_Sheets = Sheets.Count

    'Change Flag As appropriate
    Sort_Mode_Descending = False

    For Outer_Loop = 1 To No_of_Sheets
        For Inner_Loop = 1 To Outer_Loop
            If Sort_Mode_Descending = True Then
                If UCase(Sheets(Outer_Loop).Name) > UCase(Sheets(Inner_Loop).Name) Then
                    Sheets(Outer_Loop).Move Before:=Sheets(Inner_Loop)
                End If
            End If
            If Sort_Mode_Descending = False Then
                If UCase(Sheets(Outer_Loop).Name) < UCase(Sheets(Inner_Loop).Name) Then
                    Sheets(Outer_Loop).Move Before:=Sheets(Inner_Loop)
                End If
            End If
        Next Inner_Loop
    Next Outer_Loop

End Sub

Function Color_By_Numbers(cr As Range, Color_Index As Integer) As Double
    Dim Cell

    'Will look at cells that are in the range and if the color interior property
    'matches the cell color required then it will sum

    'Loop Through range
    For Each Cell In cr
        If (Cell.Interior.ColorIndex = Color_Index) Then
            Color_By_Numbers = Color_By_Numbers + Cell.Value
        End If
    Next Cell

End Function

Sub TestByWorkbookName()
    Dim wb As Workbook

    For Each wb In Workbooks
        'Replace between the quotes with workbook to test for
        If wb.Name = "New Microsoft Excel Worksheet.xls" Then
            MsgBox "Found it"
            Exit Sub                             'call code here, we'll just exit for now
        End If
    Next

End Sub

Sub UnhideAll()
    Dim WS As Worksheet

    For Each WS In Worksheets
        WS.Visible = True
    Next

End Sub


