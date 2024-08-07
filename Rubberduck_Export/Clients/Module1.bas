Attribute VB_Name = "Module1"
Option Explicit

Public iWidth As Integer
Public iHeight As Integer
Public iLeft As Integer
Public iTop As Integer
Public bState As Boolean

Sub Reset()

    Dim iRow As Long
    iRow = [Counta(Database!A:A)] 'Identifying the last row
    
    With frmForm
        .txtID.Value = ""
        .txtName.Value = ""
        .optMale.Value = False
        .optFemale.Value = False
        
        'Default Color
        .txtID.BackColor = vbWhite
        .txtName.BackColor = vbWhite
        .txtCity.BackColor = vbWhite
        .txtCountry.BackColor = vbWhite
        .cmbDepartment.BackColor = vbWhite
        '--------------------------------
        
        .cmbDepartment.Clear
        
        'Creating a dynamic name for department
        
        wshSupport.Range("A2", wshSupport.Range("A" & Application.Rows.Count).End(xlUp)).Name = "Dynamic"
        .cmbDepartment.RowSource = "Dynamic"
        .cmbDepartment.Value = ""
        .txtRowNumber.Value = ""
        
        .txtCity.Value = ""
        .txtCountry.Value = ""
        
        'Below code are associated with Search Feature - Part 3
        Call Add_SearchColumn
        ThisWorkbook.Sheets("Database").AutoFilterMode = False
        ThisWorkbook.Sheets("SearchData").AutoFilterMode = False
        ThisWorkbook.Sheets("SearchData").Cells.Clear
        '-----------------------------------------------
        
        .lstDatabase.ColumnCount = 9
        .lstDatabase.ColumnHeads = True
        
        .lstDatabase.ColumnWidths = "30;60;75;40;60;45;55;70;70"
        
        If iRow > 1 Then
            .lstDatabase.RowSource = "Database!A2:I" & iRow
        Else
            .lstDatabase.RowSource = "Database!A2:I2"
        End If
    End With

End Sub

Sub Submit()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Database")
    
    Dim iRow As Long
    If frmForm.txtRowNumber.Value = "" Then
        iRow = [Counta(Database!A:A)] + 1
    Else
        iRow = frmForm.txtRowNumber.Value
    End If
    
    With sh
        .Cells(iRow, 1) = "=Row()-1" 'Dynamic Serial Number
        .Cells(iRow, 2) = frmForm.txtID.Value
        .Cells(iRow, 3) = frmForm.txtName.Value
        .Cells(iRow, 4) = IIf(frmForm.optFemale.Value = True, "Female", "Male")
        .Cells(iRow, 5) = frmForm.cmbDepartment.Value
        .Cells(iRow, 6) = frmForm.txtCity.Value
        .Cells(iRow, 7) = frmForm.txtCountry.Value
        .Cells(iRow, 8) = Application.UserName
        .Cells(iRow, 9) = [Text(Now(), "DD-MM-YYYY HH:MM:SS")]
    End With

End Sub

Sub Show_Form()
    
    frmForm.Show

End Sub

Function Selected_List() As Long

    Selected_List = 0
    
    Dim i As Long
    For i = 0 To frmForm.lstDatabase.ListCount - 1
        If frmForm.lstDatabase.Selected(i) = True Then
            Selected_List = i + 1
            Exit For
        End If
    Next i

End Function

Sub Add_SearchColumn()

    frmForm.EnableEvents = False

    With frmForm.cmbSearchColumn
        .Clear
        .AddItem "All"
        
        .AddItem "Employee Id"
        .AddItem "Employee Name"
        .AddItem "Gender"
        .AddItem "Department"
        .AddItem "City"
        .AddItem "Country"
        .AddItem "Submitted By"
        .AddItem "Submitted On"
        
        .Value = "All"
    End With
    
    frmForm.EnableEvents = True
    
    frmForm.txtSearch.Value = ""
    frmForm.txtSearch.Enabled = False
    frmForm.cmdSearch.Enabled = False

End Sub

Sub SearchData()

    Application.ScreenUpdating = False
    
    Dim iColumn As Integer 'To hold the selected column number in Database sheet
    Dim iDatabaseRow As Long 'To store the last non-blank row number available in Database sheet
    Dim iSearchRow As Long 'To hold the last non-blank row number available in SearachData sheet
    
    Dim sColumn As String 'To store the column selection
    Dim sValue As String 'To hold the search text value
    
    Dim shDatabase As Worksheet ' Database sheet
    Set shDatabase = ThisWorkbook.Sheets("Database")
    Dim shSearchData As Worksheet 'SearchData sheet
    Set shSearchData = ThisWorkbook.Sheets("SearchData")
    
    iDatabaseRow = ThisWorkbook.Sheets("Database").Range("A" & Application.Rows.Count).End(xlUp).Row
    sColumn = frmForm.cmbSearchColumn.Value
    sValue = frmForm.txtSearch.Value
    iColumn = Application.WorksheetFunction.Match(sColumn, shDatabase.Range("A1:I1"), 0)
    
    'Remove filter from Database worksheet
    If shDatabase.FilterMode = True Then
        shDatabase.AutoFilterMode = False
    End If

    'Apply filter on Database worksheet
    If frmForm.cmbSearchColumn.Value = "Employee Id" Then
        shDatabase.Range("A1:I" & iDatabaseRow).AutoFilter Field:=iColumn, Criteria1:=sValue
    Else
        shDatabase.Range("A1:I" & iDatabaseRow).AutoFilter Field:=iColumn, Criteria1:="*" & sValue & "*"
    End If
    
    If Application.WorksheetFunction.Subtotal(3, shDatabase.Range("C:C")) >= 2 Then
        'Code to remove the previous data from SearchData worksheet
        shSearchData.Cells.Clear
        shDatabase.AutoFilter.Range.Copy shSearchData.Range("A1")
        Application.CutCopyMode = False
        iSearchRow = shSearchData.Range("A" & Application.Rows.Count).End(xlUp).Row
        frmForm.lstDatabase.ColumnCount = 9
        frmForm.lstDatabase.ColumnWidths = "30, 60, 75, 40, 60, 45, 55, 70, 70"
        If iSearchRow > 1 Then
            frmForm.lstDatabase.RowSource = "SearchData!A2:I" & iSearchRow
            MsgBox "Records found."
        End If
    Else
       MsgBox "No record found."
    End If

    shDatabase.AutoFilterMode = False
    Application.ScreenUpdating = True

End Sub

Function ValidateEntries() As Boolean

    ValidateEntries = True
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Print")
    
    Dim iEmployeeID As Variant
    iEmployeeID = frmForm.txtID.Value
    
    With frmForm
        'Default Color
        .txtID.BackColor = vbWhite
        .txtName.BackColor = vbWhite
        .txtCity.BackColor = vbWhite
        .txtCountry.BackColor = vbWhite
        .cmbDepartment.BackColor = vbWhite
        '--------------------------------
        
        If Trim(.txtID.Value) = "" Then
            MsgBox "Please enter Employee ID.", vbOKOnly + vbInformation, "Emp ID"
            ValidateEntries = False
            .txtID.BackColor = vbRed
            .txtID.SetFocus
            Exit Function
        End If
    
        'Validating Duplicate Entries
        
        If Not sh.Range("B:B").Find(what:=iEmployeeID, lookat:=xlWhole) Is Nothing Then
            MsgBox "Duplicate Employee ID found.", vbOKOnly + vbInformation, "Emp ID"
            ValidateEntries = False
            .txtID.BackColor = vbRed
            .txtID.SetFocus
            Exit Function
        End If
        
        If Trim(.txtName.Value) = "" Then
            MsgBox "Please enter Employee Name.", vbOKOnly + vbInformation, "Emp Name"
            ValidateEntries = False
            .txtName.BackColor = vbRed
            .txtName.SetFocus
            Exit Function
        End If
        
        'Validating Gender
        If .optFemale.Value = False And .optMale.Value = False Then
            MsgBox "Please select gender.", vbOKOnly + vbInformation, "Gender"
            ValidateEntries = False
            Exit Function
        End If
        
        If Trim(.cmbDepartment.Value) = "" Then
            MsgBox "Please select department name from drop-down.", vbOKOnly + vbInformation, "Dpartment"
            ValidateEntries = False
            .cmbDepartment.BackColor = vbRed
            .cmbDepartment.SetFocus
            Exit Function
        End If
        
        If Trim(.txtCity.Value) = "" Then
            MsgBox "Please enter City Name.", vbOKOnly + vbInformation, "City Name"
            ValidateEntries = False
            .txtCity.BackColor = vbRed
            .txtCity.SetFocus
            Exit Function
        End If
        
        If Trim(.txtCountry.Value) = "" Then
            MsgBox "Please enter Country Name.", vbOKOnly + vbInformation, "Country Name"
            ValidateEntries = False
            .txtCountry.BackColor = vbRed
            .txtCountry.SetFocus
            Exit Function
        End If
    End With

End Function

Function ValidatePrintDetails() As Boolean

    ValidatePrintDetails = True
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Print")
    
    Dim iEmployeeID As Variant
    iEmployeeID = frmForm.txtID.Value
    
    With frmForm
        'Default Color
        .txtID.BackColor = vbWhite
        .txtName.BackColor = vbWhite
        .txtCity.BackColor = vbWhite
        .txtCountry.BackColor = vbWhite
        .cmbDepartment.BackColor = vbWhite
        '--------------------------------
        
        If Trim(.txtID.Value) = "" Then
            MsgBox "Please enter Employee ID.", vbOKOnly + vbInformation, "Emp ID"
            ValidatePrintDetails = False
            .txtID.BackColor = vbRed
            .txtID.SetFocus
            Exit Function
        End If
        
        If Trim(.txtName.Value) = "" Then
            MsgBox "Please enter Employee Name.", vbOKOnly + vbInformation, "Emp Name"
            ValidatePrintDetails = False
            .txtName.BackColor = vbRed
            .txtName.SetFocus
            Exit Function
        End If
        
        'Validating Gender
        If .optFemale.Value = False And .optMale.Value = False Then
            MsgBox "Please select gender.", vbOKOnly + vbInformation, "Gender"
            ValidatePrintDetails = False
            Exit Function
        End If
        
        If Trim(.cmbDepartment.Value) = "" Then
            MsgBox "Please select department name from drop-down.", vbOKOnly + vbInformation, "Dpartment"
            ValidatePrintDetails = False
            .cmbDepartment.BackColor = vbRed
            .cmbDepartment.SetFocus
            Exit Function
        End If
        
        If Trim(.txtCity.Value) = "" Then
            MsgBox "Please enter City Name.", vbOKOnly + vbInformation, "City Name"
            ValidatePrintDetails = False
            .txtCity.BackColor = vbRed
            .txtCity.SetFocus
            Exit Function
        End If
        
        If Trim(.txtCountry.Value) = "" Then
            MsgBox "Please enter Country Name.", vbOKOnly + vbInformation, "Country Name"
            ValidatePrintDetails = False
            .txtCountry.BackColor = vbRed
            .txtCountry.SetFocus
            Exit Function
        End If
    End With

End Function

Sub Print_Form()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Print")
    
    With frmForm
        sh.Range("E5").Value = .txtID.Value
        sh.Range("E7").Value = .txtName.Value
        sh.Range("E9").Value = IIf(.optFemale.Value = True, "Female", "Male")
        sh.Range("E11").Value = .cmbDepartment.Value
        sh.Range("E13").Value = .txtCity.Value
        sh.Range("E15").Value = .txtCountry.Value
    End With
    
    'Code to Print the form or Export to PDF
    sh.PageSetup.PrintArea = "$B$2:$I$17"
    'sh.PrintOut copies:=1, IgnorePrintAreas:=False
    sh.ExportAsFixedFormat xlTypePDF, ThisWorkbook.Path & Application.PathSeparator & frmForm.txtName.Value & ".pdf"
    
    MsgBox "Employee details have been printed.", vbOKOnly + vbInformation, "Print"
   
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

