Attribute VB_Name = "modMisc"
Option Explicit

'Feuil1.Range("A1").Comment.Delete

'Feuil1.Range("A1").AddComment ("Hello World")

Function Create_Vector(Matrix_Range As Range) As Variant

    Dim No_of_Cols As Integer, No_Of_Rows As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Cell

    No_of_Cols = Matrix_Range.Columns.Count
    No_Of_Rows = Matrix_Range.Rows.Count
    ReDim Temp_Array(No_of_Cols * No_Of_Rows)

    'Eliminate NULL Conditions
    If Matrix_Range Is Nothing Then Exit Function
    If No_of_Cols = 0 Then Exit Function
    If No_Of_Rows = 0 Then Exit Function

    For j = 1 To No_Of_Rows
        For i = 0 To No_of_Cols - 1
            Temp_Array((i * No_Of_Rows) + j) = Matrix_Range.Cells(j, i + 1)
        Next i
    Next j

    Create_Vector = Temp_Array

End Function

Function Create_Matrix(Vector_Range As Range, No_Of_Cols_in_output As Integer, No_of_Rows_in_output As Integer) As Variant

    ReDim Temp_Array(No_Of_Cols_in_output, No_of_Rows_in_output)
    Dim No_Of_Elements_In_Vector As Integer
    Dim Col_Count As Integer, Row_Count As Integer
    Dim Cell

    No_Of_Elements_In_Vector = Vector_Range.Rows.Count

    'Eliminate NULL Conditions
    If Vector_Range Is Nothing Then Exit Function
    If No_Of_Cols_in_output = 0 Then Exit Function
    If No_of_Rows_in_output = 0 Then Exit Function
    If No_Of_Elements_In_Vector = 0 Then Exit Function

    For Col_Count = 1 To No_Of_Cols_in_output
        For Row_Count = 1 To No_of_Rows_in_output
            Temp_Array(Col_Count, Row_Count) = Vector_Range.Cells(((No_of_Rows_in_output) * (Col_Count - 1) + Row_Count), 1)
        Next Row_Count
    Next Col_Count

    Create_Matrix = Temp_Array

End Function

Sub MakeChart()
    Dim MyRange As Range
    Set MyRange = Application.InputBox(Prompt:="Select chart inputs", Type:=8)
    Charts.Add
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=MyRange, _
                              PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsNewSheet

End Sub

Sub ProgramValidate()

    Dim Choices As String
    Choices = "1. Choice1, 2. Choice2, 3. Choice3"

    Range("A1").Select

    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:=Choices
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub ListSheets()
    Dim ws As Worksheet
    Dim x As Integer

    x = 1

    Sheets("Feuil1").Range("A:A").Clear

    For Each ws In Worksheets
        Sheets("Feuil1").Cells(x, 1).Select
        ActiveSheet.Hyperlinks.Add _
        Anchor:=Selection, Address:="", SubAddress:= _
        ws.Name & "!A1", TextToDisplay:=ws.Name
        x = x + 1
    Next ws

End Sub

Sub DeleteAllShapes()
    'Activate sheet to delete autoshapes
    Feuil1.Activate

    Dim GetShape As Shape

    For Each GetShape In ActiveSheet.Shapes
        GetShape.Delete
    Next

End Sub

Sub DeleteHyperLinks()

    Feuil1.Hyperlinks.Delete

End Sub

Sub DynamicBoxes()

    Dim x As Double

    'This makes horizontal boxes
    For x = 0 To 240 Step 48

        'reference to the 4 numbers left,top,width,height
        ActiveSheet.Shapes.AddShape(msoShapeFlowchartProcess, x, 0, 48, 12.75).Select
        Selection.ShapeRange.Fill.ForeColor.SchemeColor = 11
        Selection.ShapeRange.Fill.Solid
        Selection.ShapeRange.Fill.Visible = msoTrue
    Next x

    'This makes vertical boxes
    For x = 0 To 127.5 Step 12.75

        ActiveSheet.Shapes.AddShape(msoShapeFlowchartProcess, 0, x, 48, 12.75).Select
        Selection.ShapeRange.Fill.ForeColor.SchemeColor = 11
        Selection.ShapeRange.Fill.Solid
        Selection.ShapeRange.Fill.Visible = msoTrue
    Next x

End Sub

Sub ExitWithoutPrompt()

    Application.DisplayAlerts = False
    Application.Quit

End Sub

Function ExportRange(WhatRange As Range, Where As String, Delimiter As String) As String

    Dim HoldRow As Long                          'test for new row variable
    HoldRow = WhatRange.Row
    Dim c As Range                               'loop through range variable

    For Each c In WhatRange
        If HoldRow <> c.Row Then
            'add linebreak and remove extra delimeter
            ExportRange = Left(ExportRange, Len(ExportRange) - 1) & vbCrLf & c.Text & Delimiter
            HoldRow = c.Row
        Else
            ExportRange = ExportRange & c.Text & Delimiter
        End If
    Next c

    'Trim extra delimiter
    ExportRange = Left(ExportRange, Len(ExportRange) - 1)

    'Kill the file if it already exists
    If Len(Dir(Where)) > 0 Then
        Kill Where
    End If

    Open Where For Append As #1                  'write the new file
    Print #1, ExportRange
    Close #1

End Function

Public Sub HideMyExcel()

    Application.Visible = False
    Application.Wait Now + TimeValue("00:00:05")
    Application.Visible = True

End Sub

Sub Macro1()

    Call Macro2

End Sub

Private Sub Macro2()

    MsgBox "You can only see Macro1"

End Sub

Sub Randomise_Range(Cell_Range As Range)

    Application.ScreenUpdating = False

    ' Will randomise each cell in Range
    Dim Cell

    For Each Cell In Cell_Range
        Cell.Value = Rnd * 1000
    Next Cell

    Application.ScreenUpdating = True
End Sub

Sub SizeChart2Range()

    Dim MyChart As Chart
    Dim MyRange As Range

    Set MyChart = ActiveSheet.ChartObjects(1).Chart
    Set MyRange = Feuil1.Range("B2:D6")

    With MyChart.Parent
        .Left = MyRange.Left
        .Top = MyRange.Top
        .Width = MyRange.Width
        .Height = MyRange.Height
    End With

End Sub

Sub Mail_Workbook()

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim OutApp As Object
    Dim OutMail As Object
    Dim FilePath As String
    Dim Project_Name As String
    Dim Template_Name As String
    Dim ReviewDate As String
    Dim SaveLocation As String
    Dim Path As String
    Dim Name As String

    'Create Initial variables
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    Project_Name = Sheets("Feuil1").Range("ProjectName").Value
    Template_Name = ActiveSheet.Name

    'Ask for Input used in Email
    ReviewDate = InputBox(Prompt:="Provide date by when you'd like the submission reviewed.", title:="Enter Date", Default:="MM/DD/YYYY")

    If ReviewDate = "Enter Date" Or ReviewDate = vbNullString Then GoTo endmacro

    'Save Worksheet as own workbook
    Path = ActiveWorkbook.Path
    Name = Trim(Mid(ActiveSheet.Name, 4, 99))


    Set ws = ActiveSheet
    Set oldWB = ThisWorkbook

    SaveLocation = InputBox(Prompt:="Choose File Name and Location", title:="Save As", Default:=CreateObject("WScript.Shell").SpecialFolders("Desktop") & "/" & Name & ".xlsx")

    If Dir(SaveLocation) <> "" Then
        MsgBox ("A file with that name already exists. Please choose a new name or delete existing file.")
        SaveLocation = InputBox(Prompt:="Choose File Name and Location", title:="Save As", Default:=CreateObject("WScript.Shell").SpecialFolders("Desktop") & "/" & Name & ".xlsx")
    End If
    
    If SaveLocation = vbNullString Then GoTo endmacro

    'unprotect sheet if needed
    ActiveSheet.Unprotect Password:="password"

    Set newWB = Workbooks.Add

    'Adjust Display
    ActiveWindow.Zoom = 80
    ActiveWindow.DisplayGridlines = False

    'Copy + Paste Values
    oldWB.Activate
    oldWB.ActiveSheet.Cells.Select
    Selection.Copy
    newWB.Activate
    newWB.ActiveSheet.Cells.Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                           SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
                           SkipBlanks:=False, Transpose:=False

    'Select new WB and turn off cutcopy mode
    newWB.ActiveSheet.Range("A10").Select
    Application.CutCopyMode = False

    'Save File
    newWB.SaveAs Filename:=SaveLocation, _
                 FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    FilePath = Application.ActiveWorkbook.FullName
    
    'Reprotect oldWB
    oldWB.ActiveSheet.Protect Password:="password", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                              AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                              AllowFormattingRows:=True

    'Email
    On Error Resume Next
    With OutMail
        .To = "email@email.com"
        .CC = ""
        .BCC = ""
        .Subject = Project_Name & ": " & Template_Name & " for review"
        .Body = "Project Name: " & Project_Name & ", " & Name & " For review by " & ReviewDate
        .Attachments.Add (FilePath)
        .Display
        ' .Send      'Optional to automate sending of email.
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing

    'End Macro, Restore Screenupdating, Calcs, etc...
endmacro:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub LaunchSite()

    Dim newsite As Object

    Set newsite = CreateObject("InternetExplorer.application")
    newsite.Visible = True
    newsite.Navigate "http://www.google.com/search?hl=en&ie=UTF-8&q=" & ActiveCell.Text

End Sub

Function Show_Cell_Formulae(Cell As Range) As String
    Show_Cell_Formulae = "Cell " & Cell.Address & " has the formulae: " & Cell.Formula & " '"
End Function

Sub SayThisCell()
    Cells(1, 1).Speak

End Sub

Sub SayThisString()

    Dim SayThis As String

    SayThis = "I love Microsoft Excel"
    Application.Speech.Speak (SayThis)

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Feuil1.Cells.CheckSpelling

End Sub

Sub WhatIsTheType()

    Dim Mark As Variant

    Mark = 1
    MsgBox Mark & " makes the variable " & TypeName(Mark)

    Mark = 111111
    MsgBox Mark & " makes the variable " & TypeName(Mark)

    Mark = 111111.11
    MsgBox Mark & " makes the variable " & TypeName(Mark)

End Sub

Sub Benchark()
    Dim Count As Long
    Dim BenchMark As Double
    BenchMark = Timer

    'Start of Code to Test
    For Count = 1 To 250000
        Feuil1.Cells(1, 1) = "test"
    Next Count

    'End of Code to Test
    MsgBox Format(Timer - BenchMark, "###.000") & " seconds"

End Sub

Sub MyTimer()
    Application.Wait Now + TimeValue("00:00:05")
    MsgBox ("5 seconds")

End Sub


