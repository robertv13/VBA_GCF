Attribute VB_Name = "modEFF_MACROS"
Option Explicit

Sub ChangeBackgroundColor()
    ActiveSheet.Range("A1").Interior.Color = RGB(255, 255, 0)
End Sub

Sub ChangeCellBorders()
    With ThisWorkbook.Sheets("Sheet1").Range("A1")
        On Error Resume Next
        .Borders(xlEdgeTop).LineStyle = xlDouble
        .Borders(xlEdgeBottom).LineStyle = xlDouble
        .Borders(xlEdgeLeft).LineStyle = xlDouble
        .Borders(xlEdgeRight).LineStyle = xlDouble
        On Error GoTo 0
    End With
End Sub

Sub MergeCells()
    ActiveSheet.Range("A1:B1").Merge             ' Change to your desired sheet and range to Merge
    'ActiveSheet.Range("A1:B1").UnMerge 'Use UnMerge to unmerge cells
End Sub

Sub WrapTextInCell()
    ActiveSheet.Range("A1").WrapText = True      ' Change this to your desired sheet and cell (Set to False to not wrap in cell)
End Sub

Sub AddCellComment()
    With Worksheets("Sheet1").Range("A1")        ' Modify the sheet and cell reference as needed
        .ClearComments
        .AddComment "This is a cell comment."    ' Add a comment to the specified cell
        .Comment.Visible = True                  ' Make the comment visible
    End With
End Sub

Sub SetListValidation()
    'Set List Type Validation
    Dim targetCell As Range
    Dim ValidationFormula As String
    
    ' Set the target cell where you want to change the data validation.
    Set targetCell = ActiveSheet.Range("A1")     ' Change A1 to your desired cell
    
    ' Define the list of values you want to use for data validation.
    ' You can customize this list as needed, separating values with commas. (or Use an existing named range)
    ValidationFormula = "Option1,Option2,Option3"
    targetCell.Validation.Delete                 'Remove any existing data validation from the cell.
    
    ' Apply data validation as a list to the cell.
    With targetCell.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:=ValidationFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Sub ChangeFontStyleAndSize()
    Dim targetCell As Range
    ' Set the target cell on the active sheet
    Set targetCell = ActiveSheet.Range("A1")     ' Change this to your desired cell
    targetCell.Font.Name = "Arial"               'Set Font Style
    targetCell.Font.Size = 12                    'Set Font Size
End Sub

Sub LockAndProtectCells()
    ' Lock and protect specific cells or a range in the active worksheet
    ActiveSheet.Unprotect                        ' Remove protection (if applied)
    ' Define your cell or range selection
    ActiveSheet.Range("A1:B10").Locked = True    ' Lock the specified cells
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
End Sub

Sub PasswordProtectCells()
    ' Password protect specific cells or a range in the active worksheet
    Dim Password As String
    Password = "YourPassword"                    ' Set your desired password
    ActiveSheet.Protect Password:=Password, DrawingObjects:=False, Contents:=True, Scenarios:=False
End Sub

Sub FindAndReplace()
    Dim WS As Worksheet
    Dim SearchValue As Variant, ReplaceValue As Variant
    
    Set WS = ThisWorkbook.Worksheets("Sheet1")   ' Set the target worksheet
    
    ' Define the value you want to find and replace
    SearchValue = "Old Value"
    ReplaceValue = "New Value"
    
    ' Perform the find and replace operation
    WS.Cells.Replace What:=SearchValue, Replacement:=ReplaceValue, LookAt:=xlWhole, _
                     SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    ' Optionally, display a message to confirm the replacement
    MsgBox "Values replaced: " & SearchValue & " with " & ReplaceValue
End Sub

Sub FormatCellAsGeneral()
    ' Reference the cell you want to format (change A1 to the desired cell)
    Dim targetCell As Range
    Set targetCell = ThisWorkbook.Sheets("Sheet1").Range("A1")

    ' Apply the "General" format to the cell
    targetCell.NumberFormat = "General"
End Sub

Sub FormatCellAsCurrency()
    ActiveSheet.Range("A1").NumberFormat = "$#,##0.00"
End Sub

Sub FormatCellAsText()
    ActiveSheet.Range("A1").NumberFormat = "@"
End Sub

Sub ConvertDateToString()
    Dim UserInput As String
    Dim ConvertedDate As Date
    
    ' Prompt the user for a date input
    UserInput = InputBox("Enter a date (YYYY-MM-DD):")
    
    ' Check if the user input is a valid date
    If IsDate(UserInput) Then
        ' If it's a valid date, convert it using CDate
        ConvertedDate = CDate(UserInput)
        MsgBox "You entered a valid date: " & ConvertedDate
    Else
        MsgBox "Invalid date format. Please enter a date in the format YYYY-MM-DD."
    End If
End Sub

Sub CreateDateWithDateSerial()
    Dim YearValue As Integer
    Dim MonthValue As Integer
    Dim DayValue As Integer
    Dim NewDate As Date
    
    ' Specify year, month, and day components
    YearValue = 2023
    MonthValue = 11
    DayValue = 8
    
    ' Create a date using DateSerial
    NewDate = DateSerial(YearValue, MonthValue, DayValue)
    
    ' Display the created date
    MsgBox "Created Date: " & Format(NewDate, "Long Date")
End Sub

Sub GetCurrentDate()
    Dim CurDateTime As Date
    CurDateTime = Now                            'Set Current Date & time
End Sub

Sub ReturnWeekdayNumber()
    Dim DayNumber As Long
    DayNumber = Weekday(Date)
    MsgBox "The current day number is " & DayNumber
End Sub

Sub FormatDate()
    Dim FormattedDate As String
    FormattedDate = Format(Date, "dd/mm/yyyy")
End Sub

Sub BrowseForAnyFile()
    Dim SelectedFilePath As String
    Dim fileDialog As fileDialog
    
    ' Create a FileDialog object
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    
    ' Set properties of the file dialog
    With fileDialog
        .title = "Select an Excel Workbook"
        .Filters.Clear
        .Filters.Add "Select Any File", "*.*"    'Wildcards used for any file type
        
        ' Show the file dialog and store the selected file path
        If .Show = -1 Then
            SelectedFilePath = .SelectedItems(1)
            ' Display the selected file path
            MsgBox "Selected file path: " & SelectedFilePath, vbInformation
        Else
            ' User canceled or closed the dialog
            MsgBox "No file selected.", vbInformation
        End If
    End With
End Sub

Sub BrowseForFolder()
    Dim YourFolder As fileDialog
    Dim FolderPath As String
    Set YourFolder = Application.fileDialog(msoFileDialogFolderPicker)
    With YourFolder
        .title = "Please browse for your folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NotSelected
        FolderPath = .SelectedItems(1)           'Full Folder Path
        MsgBox FolderPath
NotSelected:
    End With
End Sub

Sub ExportToCSV()
    Dim WS As Worksheet                          ' Define variables
    Dim ExportRange As Range
    Dim filePath As String
    Dim lastRow As Long
    ' Set the worksheet to export
    Set WS = ThisWorkbook.Sheets("Sheet1")       ' Change to the desired worksheet
    lastRow = WS.Range("A9999").End(xlUp).Row    'Set Last Row of data
    ' Set the range of data to export (e.g., A1 to C10)
    Set ExportRange = WS.Range("A1:J" & lastRow) ' Change to the desired range

    ' Specify the path and filename for the CSV file
    filePath = "C:\Users\neilr\Desktop\ExportedDatas.csv" ' Change to the desired file path and name

    ' Export data to CSV file
    ExportRange.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs filePath, xlCSV
    ActiveWorkbook.Close False

    ' Inform the user about the export
    MsgBox "Data exported to CSV file.", vbInformation
End Sub

Sub CreateZipFile()
    Dim FolderToZipPath As Variant, ZippedFileFullName As Variant
    Dim ShellApp As Object
    FolderToZipPath = "C:\Path\To\YourFolderToZip"
    ZippedFileFullName = "C:\Path\To\YourFolder\YourZippedFiles.zip"
    'Create an empty zip file
    Open ZippedFileFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1

    'Copy the files & folders into the zip file
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(ZippedFileFullName).CopyHere ShellApp.Namespace(FolderToZipPath).Items

    'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
    On Error Resume Next
    Do Until ShellApp.Namespace(ZippedFileFullName).Items.Count = ShellApp.Namespace(FolderToZipPath).Items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0
End Sub

Sub Send_Email()
    Dim OutApp As Object, OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody = "This is the body of the email"

    On Error Resume Next
    With OutMail
        .To = "recipient@example.com"
        .CC = ""
        .BCC = ""
        .Subject = "This is the subject of the email"
        .Body = strbody
        .Attachments.Add ActiveWorkbook.FullName 'Add Attachments If any
        .Display                                 'Use .Send to send without displaying
    End With
    On Error GoTo 0

    Set OutMail = Nothing                        'Clear Out Mail Variable
    Set OutApp = Nothing                         'Clear Out Outlook Application Variable
End Sub

Sub RunAdvancedFilterWithCriteria()
    Dim lastRow As Long
    lastRow = ActiveSheet.Range("A99999").End(xlUp).Row
    ActiveSheet.Range("A3:Z" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=ActiveSheet.Range("AA2:AB3"), CopyToRange:=ActiveSheet.Range("AE2:AF2"), Unique:=True
End Sub

Sub AdvancedSort()
    ' Define the worksheet and range to be sorted and find details below
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Sheets("Sheet1")
    Dim rng As Range
    Set rng = WS.Range("A1:E10")                 ' Adjust the range as needed

    ' Sort the range with advanced options
    With rng
        .Sort Key1:=.Range("A2"), Order1:=xlAscending, header:=xlYes, _
              OrderCustom:=1, MatchCase:=False, _
              Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
              DataOption2:=xlSortNormal, DataOption3:=xlSortNormal
    End With
End Sub

Sub FindSpecialCells()
    Dim rng As Range
    Dim SpecialCells As Range

    ' Set the range where you want to search for special cells
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("A1:A10")

    ' Use the SpecialCells method to find cells with specific characteristics
    ' For example, find all blank cells in the range
    On Error Resume Next
    Set SpecialCells = rng.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    ' Check if special cells were found
    If Not SpecialCells Is Nothing Then
        ' Special cells found, select or manipulate them as needed
        SpecialCells.Select                      ' You can perform actions on the selected cells
    Else
        ' No special cells found
        MsgBox "No special cells with the specified characteristics found in the range."
    End If
End Sub

Sub SetPrintOptions()
    Dim PrintRange As Range
    Set PrintRange = ThisWorkbook.Sheets("Sheet1").Range("A1:E10") ' Change to your desired sheet and range
    Dim WS As Worksheet
    Set WS = PrintRange.Worksheet
   
    With WS.PageSetup                            ' Set page layout options for the worksheet
        .Orientation = xlLandscape               ' Change to xlPortrait for portrait orientation
        .PaperSize = xlPaperLetter               ' Change to your desired paper size
        .TopMargin = Application.InchesToPoints(0.5) ' Adjust margin values as needed
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True               ' Center on page horizontally
        .CenterVertically = True                 ' Center on page vertically
        .HeaderMargin = Application.InchesToPoints(0.25) ' Header margin size
        .FooterMargin = Application.InchesToPoints(0.25) ' Footer margin size
        .Zoom = 100                              ' Set zoom level (in percent)
        .PrintGridlines = True                   ' Print gridlines
        .PrintHeadings = False                   ' Print row and column headings
    End With
    
    ' Set print options
    WS.PageSetup.PrintArea = PrintRange.Address
    WS.PageSetup.PrintTitleRows = "$1:$1"        ' Set print title rows if needed
    WS.PageSetup.PrintTitleColumns = "$A:$A"     ' Set print title columns if needed
    
    PrintRange.PrintOut                          ' Print the specified range
End Sub

Sub FormatShape()
    Dim WS As Worksheet
    Dim shp As Shape
    Set WS = ThisWorkbook.Sheets("Sheet1")       ' Set the worksheet where the shape exists

    ' Check if the shape "MyShape" exists on the worksheet
    On Error Resume Next
    Set shp = WS.Shapes("MyShape")
    On Error GoTo 0

    If Not shp Is Nothing Then
        ' Format the shape properties
        With shp
            .Fill.ForeColor.RGB = RGB(0, 255, 0) ' Set fill color to green
            .Line.ForeColor.RGB = RGB(0, 0, 255) ' Set line color to blue
            .TextFrame.Characters.Text = "Formatted Shape" ' Add or modify text within the shape
            .TextFrame.Characters.Font.Size = 12 ' Set font size
        End With
    Else
        MsgBox "Shape not found on the worksheet."
    End If
End Sub

Sub AddAndNameUserForm()
    ' Declare a variable to hold the new UserForm
    Dim uf As Object
    
    ' Create a new UserForm and assign it to the variable
    Set uf = ThisWorkbook.VBProject.VBComponents.Add(3)
    
    ' Define the name for the UserForm (change "MyUserForm" to your desired name)
    uf.Name = "MyUserForm"
    
    ' Optional: Add controls, code, and design elements to the UserForm here
    
    ' Show the UserForm (optional)
    ' uf.Show
End Sub

Sub CopyAndRenameWorksheet()
    ' Define the source worksheet to be copied
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Sheets("SourceSheet") ' Adjust the source sheet name as needed

    ' Copy the source worksheet to a new worksheet
    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    ' Rename the new worksheet
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "NewSheetName" ' Adjust the new sheet name as needed
End Sub

Sub CopyWorkbook()
    Dim SourcePath As String
    Dim DestinationPath As String
    
    ' Define the source and destination paths
    SourcePath = "C:\SourceFolder\SourceWorkbook.xlsx" ' Replace with the source workbook path (workbook should be closed)
    DestinationPath = "C:\DestinationFolder\DestinationWorkbook.xlsx" ' Replace with the destination workbook path
    
    ' Copy the source workbook to the destination
    FileCopy Source:=SourcePath, Destination:=DestinationPath
    
    ' Optionally, you can open the copied workbook
    Workbooks.Open DestinationPath
End Sub

Sub ListMacrosAndModulesFromWorkbook()
    Dim xlApp As Object, xlWorkbook As Object, vbProj As Object, vbComp As Object
    Dim ModuleCount As Long, CodeLine As Integer, CurRow As Long
    Dim LineText As String

    Set xlApp = CreateObject("Excel.Application") 'Create an instance of Excel application
    xlApp.Visible = False                        'You can set this to True if you want to see the Excel application
   
    Set xlWorkbook = xlApp.Workbooks.Open("C:\VBA\Reference\Excel for Freelancers\" & _
                                          "Personal_Macro_Workbook.xlsm") 'Open the workbook
    
    Set vbProj = xlWorkbook.VBProject            'Create an instance of the VBIDE.VBProject object
    
    With Sheets("Feuil1")
        .Range("A1:B999").ClearContents          'Clear the existing content in a worksheet _
                                        '(change "Sheet1" to your sheet's name)
        ' Set up the header row
        .Range("A1").Value = "Module Name"
        .Range("B1").Value = "Macro Name"
    End With
    
    CurRow = 2                                   'Set the intial row number
    
    'Loop through all VBComponents in the workbook
    For Each vbComp In vbProj.VBComponents
      
        If vbComp.Type = 1 Then                  'Check if the component is a module
            Sheets("Feuil1").Range("A" & CurRow).Value = vbComp.Name ' List module name _
                                                         'in the first column
            
            'Loop through all macros in the module
            For CodeLine = 1 To vbComp.CodeModule.CountOfLines
                'Check if the line contains a Sub or Function definition
                If Left(Trim(vbComp.CodeModule.Lines(CodeLine, 1)), 3) = "Sub" Or Left(Trim(vbComp.CodeModule.Lines(CodeLine, 1)), 8) = "Function" Then
                    LineText = vbComp.CodeModule.Lines(CodeLine, 1)
                    'Extract the macro name and list it in the second column
                    LineText = Replace(Replace(LineText, "Sub ", ""), "Function ", "") 'Remove Sub or Function name
                    LineText = Left(LineText, InStr(LineText, "(") - 1)
                    Sheets("Feuil1").Cells(CurRow, 2).Value = LineText
                    
                    'Move to the next row
                    CurRow = CurRow + 1
                End If
            Next CodeLine
        End If
    Next vbComp
    
    'Close the workbook without saving changes
    'xlWorkbook.Close False
    
    'Quit Excel application
    xlApp.Quit
    
    'Release objects
    Set vbProj = Nothing
    Set vbComp = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
End Sub


