Attribute VB_Name = "modFunctions"
Option Explicit

Function Fn_GetID_From_Initials(i As String)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("dnrProf_All")
        If cell.Value2 = i Then
            Fn_GetID_From_Initials = cell.Offset(0, 1).value
            Exit Function
        End If
    Next cell

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    
End Function

Function Fn_GetID_From_Client_Name(nomCLient As String) '2024-02-14 @ 06:07

    Dim ws As Worksheet: Set ws = wshBD_Clients
    
    On Error Resume Next
    Dim dynamicRange As Range: Set dynamicRange = ws.Range("dnrClients_All")
    On Error GoTo 0

    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'Clients' ou le DynamicRange 'dnrClients_All' n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(nomCLient, _
        dynamicRange.columns(1), dynamicRange.columns(2), _
        "Not Found", 0, 1)
    
    If result <> "Not Found" Then
        Fn_GetID_From_Client_Name = result
    Else
        MsgBox "Impossible de retrouver la valeur dans la première colonne du client", vbExclamation
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set dynamicRange = Nothing
    Set ws = Nothing

End Function

Function Fn_FournID_From_Fourn_Name(nomFournisseur As String) '2024-07-03 @ 16:13

    Dim ws As Worksheet: Set ws = wshBD_Fournisseurs
    
    On Error Resume Next
    Dim dynamicRange As Range: Set dynamicRange = ws.Range("dnrSuppliers_All")
    On Error GoTo 0

    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'BD_Fournisseurs' ou le DynamicRange 'dnrSuppliers_All' n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(nomFournisseur, _
        dynamicRange.columns(1), dynamicRange.columns(2), _
        "Not Found", 0, 1)
    
    If result <> "Not Found" Then
        Fn_FournID_From_Fourn_Name = result
    Else
        Fn_FournID_From_Fourn_Name = 0
    End If
    
    'Cleaning memory - 2024-07-03 @ 16:13
    Set dynamicRange = Nothing
    Set ws = Nothing

End Function

Function Fn_Find_Data_In_A_Range(r As Range, cs As Long, ss As String, cr As Long) As Variant() '2024-03-29 @ 05:39
    
    'This function is used to retrieve information from a range
    'If found, it returns Variant, with the cell address, the row and the value
    '2024-03-09 - First version
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("Functions:Fn_Find_Data_In_A_Range()")
    
    Dim foundInfo(1 To 3) As Variant 'Address, Row, Value
    Dim dataValue As Variant
    
    'Search for the string in a given range (r) at the column specified (cs)
    Dim foundCell As Range: Set foundCell = r.columns(cs).Find(What:=ss, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the string was found
    If Not foundCell Is Nothing Then
        'With the foundCell get the the address, the row number and the value
        foundInfo(1) = foundCell.Address
        foundInfo(2) = foundCell.row
        foundInfo(3) = foundCell.Offset(0, cr - cs).value 'Return Column - Searching column
        Fn_Find_Data_In_A_Range = foundInfo 'foundInfo is an array
    Else
        Fn_Find_Data_In_A_Range = foundInfo 'foundInfo is an array
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set foundCell = Nothing

    Call Output_Timer_Results("Functions:Fn_Find_Data_In_A_Range()", timerStart)

End Function

Function GetCheckBoxPosition(chkBox As OLEObject) As String

    'Get the cell that contains the top-left corner of the CheckBox
    GetCheckBoxPosition = chkBox.TopLeftCell.Address
    
End Function

Public Function Fn_Get_GL_Code_From_GL_Description(glDescr As String) 'XLOOKUP - 2024-01-09 @ 09:19

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("Functions:Fn_Get_GL_Code_From_GL_Description()")
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Admin")
    
    On Error Resume Next
    Dim dynamicRange As Range: Set dynamicRange = ws.Range("dnrPlanComptable_All")
    On Error GoTo 0
    
    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'Admin' ou le DynamicRange n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(glDescr, _
        dynamicRange.columns(1), dynamicRange.columns(2), _
        "Not Found", 0, 1)
    
    If result <> "Not Found" Then
        Fn_Get_GL_Code_From_GL_Description = result
    Else
        MsgBox "Impossible de retrouver la valeur dans la première colonne", vbExclamation
    End If

    'Cleaning memory - 2024-07-01 @ 09:34
    Set dynamicRange = Nothing
    Set ws = Nothing

    Call Output_Timer_Results("Functions:Fn_Get_GL_Code_From_GL_Description()", timerStart)

End Function

Public Function Fn_Get_TEC_Row_Number_By_TEC_ID(ByVal uniqueID As Variant, ByVal lookupRange As Range) As Long
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("Functions:Fn_Get_TEC_Row_Number_By_TEC_ID()")
    
    Dim matchResult As Variant

    'Use the Match function to find the row number of the unique TEC_ID
    matchResult = Application.Match(uniqueID, lookupRange.columns(1), 0)
    matchResult = matchResult + 2 'Two header lines...

    'Check if Match found a result
    If Not IsError(matchResult) Then
        Fn_Get_TEC_Row_Number_By_TEC_ID = matchResult
    Else
        'If Match did not find a result, return 0
        Fn_Get_TEC_Row_Number_By_TEC_ID = 0
    End If
    
    Call Output_Timer_Results("Functions:Fn_Get_TEC_Row_Number_By_TEC_ID()", timerStart)
    
End Function

Function Fn_Get_AR_Balance_For_Invoice(ws As Worksheet, invNo As String)

    'Define the source data
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).row
    If lastUsedRow < 4 Then Exit Function
    
    'Define the range for the source data
    Dim sourceRng As Range: Set sourceRng = ws.Range("A3:F" & lastUsedRow)
    
    'Define the criteria range
    Dim criteriaRng As Range: Set criteriaRng = ws.Range("V2:V3")
    ws.Range("V3").value = invNo
    
    'Define the destination range & clear the old data
    Dim destinationRng As Range: Set destinationRng = ws.Range("X3:AC3")
    lastUsedRow = ws.Range("X9999").End(xlUp).row
    If lastUsedRow > 3 Then
        ws.Range("X4:AB" & lastUsedRow).ClearContents
    End If
    
    'Execute the AdvancedFilter
    sourceRng.AdvancedFilter xlFilterCopy, criteriaRng, destinationRng, False
    
    lastUsedRow = ws.Range("X9999").End(xlUp).row
    If lastUsedRow < 3 Then
        Fn_Get_AR_Balance_For_Invoice = 0
    Else
        Dim i As Long, balanceFacture As Currency
        For i = 4 To lastUsedRow
            balanceFacture = balanceFacture + CCur(ws.Range("AB" & i).value)
        Next i
        Fn_Get_AR_Balance_For_Invoice = balanceFacture
    End If

    'Cleaning memory - 2024-07-01 @ 09:34
    Set criteriaRng = Nothing
    Set destinationRng = Nothing
    Set sourceRng = Nothing
    
End Function

Function Fn_ValidateDaySpecificMonth(d As Integer, m As Integer, y As Integer) As Boolean
    'Returns TRUE or FALSE if d, m and y combined are VALID values
    
    Fn_ValidateDaySpecificMonth = False
    
    Dim isLeapYear As Boolean
    If y Mod 4 = 0 And (y Mod 100 <> 0 Or y Mod 400 = 0) Then
        isLeapYear = True
    Else
        isLeapYear = False
    End If
    
    'Last day of each month (0 to 11)
    Dim mdpm As Variant
    mdpm = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    If isLeapYear Then mdpm(1) = 29 'Adjust February for Leap Year
    
    If m < 1 Or m > 12 Or _
       d > mdpm(m - 1) Or _
       Abs(year(Now()) - y) > 75 Then
            Exit Function
    Else
        Fn_ValidateDaySpecificMonth = True
    End If

End Function

Function CompleteDate(dateInput As String) As Variant
    Dim defaultDay As Integer
    Dim defaultMonth As Integer
    Dim defaultYear As Integer
    Dim dayPart As Integer
    Dim monthPart As Integer
    Dim yearPart As Integer
    Dim parsedDate As Date
    Dim parts() As String
    
    'Catch all errors
    On Error GoTo InvalidDate
    
    'Get the current date components
    defaultDay = day(Date)
    defaultMonth = month(Date)
    defaultYear = year(Date)
    
    ' Split the input date into parts, considering different delimiters
    dateInput = Replace(Replace(Replace(dateInput, "/", "-"), ".", "-"), " ", "")
    parts = Split(Replace(dateInput, "-01-1900", ""), "-")
    
    Select Case UBound(parts)
        Case -1
            'Nothing provided
            dayPart = defaultDay       'Use current day
            monthPart = defaultMonth   'Use current month
            yearPart = defaultYear     'Use current year
        Case 0
            'Only day provided
            dayPart = CInt(parts(0))   'Use entered day
            monthPart = defaultMonth   'Use current month
            yearPart = defaultYear     'Use current year
        Case 1
            'Day and month provided
            dayPart = CInt(parts(0))   'Use entered day
            monthPart = CInt(parts(1)) 'Use entered month
            yearPart = defaultYear     'Use current year
        Case 2
            'Day, month, and year provided
            dayPart = CInt(parts(0))   'Use entered day
            monthPart = CInt(parts(1)) 'Use entered month
            yearPart = CInt(parts(2))  'Use entered year
        Case Else
            GoTo InvalidDate
    End Select
    
    'Fine validation taking into consideration leap year AND 75 years (past or future)
    If Fn_ValidateDaySpecificMonth(dayPart, monthPart, yearPart) = False Then
        GoTo InvalidDate
    End If
    
    'Construct the full date
    parsedDate = DateSerial(yearPart, monthPart, dayPart)
    
    'Return a VALID date
    CompleteDate = CDate(parsedDate)
    Exit Function

InvalidDate:

    CompleteDate = "Invalid Date"
    
End Function

Function Fn_Sort_Dictionary_By_Value(dict As Object, Optional descending As Boolean = False) As Variant '2024-07-11 @ 15:16
    
    'Sort a dictionary by its values and return keys in an array
    Dim keys() As Variant
    Dim values() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    
    ReDim keys(0 To dict.count - 1)
    ReDim values(0 To dict.count - 1)
    
    Dim key As Variant
    i = 0
    For Each key In dict.keys
        keys(i) = key
        values(i) = dict(key)
        i = i + 1
    Next key
    
    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If (values(i) < values(j) And descending) Or (values(i) > values(j) And Not descending) Then
                'Swap values
                temp = values(i)
                values(i) = values(j)
                values(j) = temp
                
                'Swap keys accordingly
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    Fn_Sort_Dictionary_By_Value = keys
    
End Function

Public Function Fn_TEC_Is_Data_Valid() As Boolean

    Fn_TEC_Is_Data_Valid = False
    
    'Validations first (one field at a time)
    If ufSaisieHeures.cmbProfessionnel.value = "" Then
        MsgBox Prompt:="Le professionnel est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.cmbProfessionnel.SetFocus
        Exit Function
    End If

    If ufSaisieHeures.txtDate.value = "" Or IsDate(ufSaisieHeures.txtDate.value) = False Then
        MsgBox Prompt:="La date est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.txtDate.SetFocus
        Exit Function
    End If

    If ufSaisieHeures.txtClient.value = "" Then
        MsgBox Prompt:="Le client est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.txtClient.SetFocus
        Exit Function
    End If
    
    If ufSaisieHeures.txtHeures.value = "" Or IsNumeric(ufSaisieHeures.txtHeures.value) = False Then
        MsgBox Prompt:="Le nombre d'heures est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        ufSaisieHeures.txtHeures.SetFocus
        Exit Function
    End If

    Fn_TEC_Is_Data_Valid = True

End Function

Public Function Fn_Get_Hourly_Rate(ProfID As Integer, dte As Date)

        'Use the Dynamic Named Range
        Dim rng As Range
        On Error Resume Next
        Set rng = ThisWorkbook.Names("dnrTauxHoraire").RefersToRange
        On Error GoTo 0

        'Check if the range is set correctly
        If Not rng Is Nothing Then
            Dim rowRange As Range
            Dim i As Long
            'Loop through each row in the range
            For i = rng.rows.count To 1 Step -1
                'Set the row range
                Set rowRange = rng.rows(i)
                If rowRange.Cells(1, 1).value = ProfID Then
                    If CDate(dte) >= CDate(rowRange.Cells(1, 2).value) Then
                        Fn_Get_Hourly_Rate = rowRange.Cells(1, 3).value
                        Exit Function
                    End If
                End If
                'Loop through each cell in the row
            Next i
        Else
            MsgBox "La plage nommée 'dnrTauxHoraire' n'a pas été trouvée!", vbExclamation
        End If

End Function

Public Function Fn_Get_Tax_Rate(d As Date, taxType As String) As Double

    Dim row As Integer
    Dim rate As Double
    With wshAdmin
        For row = 18 To 11 Step -1
            If .Range("L" & row).value = taxType Then
                If d >= .Range("M" & row).value Then
                    rate = .Range("N" & row).value
                    Exit For
                End If
            End If
        Next row
    End With
    
    Fn_Get_Tax_Rate = rate
    
End Function

Function Fn_Is_Date_Valide(d As String) As Boolean

    Fn_Is_Date_Valide = False
    If d = "" Or IsDate(d) = False Then
        MsgBox "Une date d'écriture est obligatoire." & vbNewLine & vbNewLine & _
            "Veuillez saisir une date valide!", vbCritical, "Date Invalide"
    Else
        Fn_Is_Date_Valide = True
    End If

End Function

Function Fn_Is_Ecriture_Balance() As Boolean

    Fn_Is_Ecriture_Balance = False
    If wshGL_EJ.Range("H26").value <> wshGL_EJ.Range("I26").value Then
        MsgBox "Votre écriture ne balance pas." & vbNewLine & vbNewLine & _
            "Débits = " & wshGL_EJ.Range("H26").value & " et Crédits = " & wshGL_EJ.Range("I26").value & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée.", vbCritical, "Veuillez vérifier votre écriture!"
    Else
        Fn_Is_Ecriture_Balance = True
    End If

End Function

Function Fn_Is_Debours_Balance() As Boolean

    Fn_Is_Debours_Balance = False
    If wshDEB_Saisie.Range("O6").value <> wshDEB_Saisie.Range("I26").value Then
        MsgBox "Votre transaction ne balance pas." & vbNewLine & vbNewLine & _
            "Total saisi = " & Format(wshDEB_Saisie.Range("O6").value, "#,##0.00 $") _
            & " vs. Ventilation = " & Format(wshDEB_Saisie.Range("I26").value, "#,##0.00 $") _
            & vbNewLine & vbNewLine & "Elle n'est donc pas reportée.", _
            vbCritical, "Veuillez vérifier votre écriture!"
    Else
        Fn_Is_Debours_Balance = True
    End If

End Function

Function Fn_Is_JE_Valid(rmax As Long) As Boolean

    Fn_Is_JE_Valid = True 'Optimist
    If rmax <= 9 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        Fn_Is_JE_Valid = False
    End If
    
    Dim i As Long
    For i = 9 To rmax
        If wshGL_EJ.Range("E" & i).value <> "" Then
            If wshGL_EJ.Range("H" & i).value = "" And wshGL_EJ.Range("I" & i).value = "" Then
                MsgBox "Il existe une ligne avec un compte, sans montant !"
                Fn_Is_JE_Valid = False
            End If
        End If
    Next i

End Function

Function Fn_Is_Deb_Saisie_Valid(rmax As Long) As Boolean

    Fn_Is_Deb_Saisie_Valid = True 'Optimist
    If rmax < 9 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        Fn_Is_Deb_Saisie_Valid = False
    End If
    
    Dim i As Long
    For i = 9 To rmax
        If wshDEB_Saisie.Range("E" & i).value <> "" Then
            If wshDEB_Saisie.Range("N" & i).value = "" Then
                MsgBox "Il existe une ligne avec un compte, sans montant !"
                Fn_Is_Deb_Saisie_Valid = False
            End If
        End If
    Next i

End Function

Public Function Fn_Clear_Range_Borders(r As Range)

    'MsgBox "Range to clear = " & r.Address
    With r
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

End Function

Public Function Fn_Pad_A_String(s As String, fillCaracter As String, length As Integer, leftOrRight As String) As String

    Dim paddedString As String
    Dim charactersNeeded As Integer
    
    charactersNeeded = length - Len(s)
    
    If charactersNeeded > 0 Then
        If leftOrRight = "R" Then
            paddedString = s & String(charactersNeeded, fillCaracter)
        Else
            paddedString = String(charactersNeeded, fillCaracter) & s
        End If
    Else
        paddedString = s
    End If

    Fn_Pad_A_String = paddedString
        
End Function

Function Fn_Get_Chart_Of_Accounts(nbCol As Integer) As Variant '2024-06-07 @ 07:31

    'Reference the named range
    Dim planComptable As Range: Set planComptable = wshAdmin.Range("dnrPlanComptable_All")
    
    'Iterate through each row of the named range
    Dim rowNum As Long, row As Range, rowRange As Range
    Dim arr() As String
    If nbCol = 1 Then
        ReDim arr(1 To planComptable.rows.count) As String '1D array
    Else
        ReDim arr(1 To planComptable.rows.count, 1 To 2) As String '2D array
    End If
    For rowNum = 1 To planComptable.rows.count
        'Get the entire row as a range
        Set rowRange = planComptable.rows(rowNum)
        'Process each cell in the row
        For Each row In rowRange.rows
            If nbCol = 1 Then
                arr(rowNum) = row.Cells(1, 2)
            Else
                arr(rowNum, 1) = row.Cells(1, 2)
                arr(rowNum, 2) = row.Cells(1, 1)
            End If
        Next row
    Next rowNum
    
    Fn_Get_Chart_Of_Accounts = arr
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set planComptable = Nothing
    Set row = Nothing
    Set rowRange = Nothing
    
End Function

Public Function GetCurrentRegion(ByVal dataRange As Range, Optional headerSize As Long = 1) As Range

    Set GetCurrentRegion = dataRange.CurrentRegion
    If headerSize > 0 Then
        With GetCurrentRegion
            'Remove the header
            Set GetCurrentRegion = .Offset(headerSize).Resize(.rows.count - headerSize)
        End With
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set GetCurrentRegion = Nothing
    
End Function

Public Function ConvertValueBooleanToText(val As Boolean) As String

    Select Case val
        Case 0, "False" 'False
            ConvertValueBooleanToText = "FAUX"
        Case -1, "True" 'True
            ConvertValueBooleanToText = "VRAI"
        Case "VRAI", "FAUX"
            
        Case Else
            MsgBox val & " est une valeur INVALIDE !"
    End Select

End Function

Sub Fn_Get_Tax_RateZ(r As Range, d As Date, tx As String)
    
'    'Set the range to search
'    Dim dataRange As Range: Set dataRange = r
'
'    'Setup return value (rate)
'    Dim rate As Double
'    rate = 0
'
'    'Loop through the data range
'    Dim cell As Range
'    For Each cell In dataRange.columns(1).Cells
'        If cell.value = tx And cell.Offset(0, 1).value < d Then
'            'If the code matches and the date is smaller, store the result
'            rate = cell.value
'            rate = cell.Offset(0, 1).value
'        End If
'    Next cell
'
'    MsgBox "Search complete. Results are in columns D and E."
    
End Sub

Public Function GetOneDrivePath(ByVal fullWorkbookName As String) As String '2024-05-27 @ 10:10
    
    'Try the 3 key types in the registry to find the file
    Dim oneDrive As Variant
    oneDrive = Array("OneDriveCommercial", "OneDriveConsumer", "OneDrive")
    
    Dim ShellScript As Object
    Set ShellScript = CreateObject("WScript.Shell")
    Dim oneDriveRegLocalPath As String
    
    Dim key As Variant
    For Each key In oneDrive
    
        'Get the Get OneDrive path from the registry - If doesn't exist go to the next key
        On Error Resume Next
        oneDriveRegLocalPath = ShellScript.RegRead("HKEY_CURRENT_USER\Environment\" & key)
        If oneDriveRegLocalPath = vbNullString Then GoTo continue
        On Error GoTo 0
                    
        'Get the end part of the path from the URL name
        Dim fileEndPart As String
        fileEndPart = GetEndPath(fullWorkbookName)
        If Len(fileEndPart) = 0 Then GoTo continue
        
        'Build the final filename by combining registry drive and the end part of url
        GetOneDrivePath = Replace(oneDriveRegLocalPath & fileEndPart, "/", "\")
        
        'Check if the file exists
        If Dir(GetOneDrivePath) = "" Then
            GetOneDrivePath = ""
        Else
            Exit For
        End If
continue:
    Next key
    
    If GetOneDrivePath = "" Then Err.Raise 53, "GetOneDrivePath" _
                , "Could not find the file [" & fullWorkbookName & "] on OneDrive."
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set key = Nothing
    
End Function

Public Function GetEndPath(ByVal fullWorkbookName As String) As String

    'Remove the url part of the name which is preceded by the text "/Documents"
    If InStr(1, fullWorkbookName, "my.sharepoint.com") <> 0 Then
        'Get the part of the string after "/Documents"
        Dim arr() As String
        arr = Split(fullWorkbookName, "/Documents")
        GetEndPath = arr(UBound(arr))
    ElseIf InStr(1, fullWorkbookName, "d.docs.live.net") <> 0 Then
        'Get the part of the filename without the URL
        Dim firstPart As String
        firstPart = Split(fullWorkbookName, "/")(4)
        GetEndPath = Mid(fullWorkbookName, InStr(fullWorkbookName, firstPart) - 1)
    Else
        GetEndPath = ""
    End If
    
End Function

Function GetQuarterDates(fiscalYearStartMonth As Integer, fiscalYear As Integer) As String
    Dim startDate As Date
    Dim endDate As Date
    Dim quarterDates As String
    Dim i As Integer
    
    'Initialize the quarterDates variable
    quarterDates = ""

    'Loop through the 4 quarters
    For i = 0 To 3
        'Calculate the start date of the quarter
        startDate = DateSerial(fiscalYear, fiscalYearStartMonth + (i * 3), 1)
        
        'Calculate the end date of the quarter
        endDate = DateAdd("m", 3, startDate) - 1
        
        'Add the quarter dates to the string
        quarterDates = quarterDates & "Q" & (i + 1) & ": " & Format(startDate, "dd/mm/yyyy") & " to " & Format(endDate, "dd-mmm-yyyy") & vbCrLf
    Next i
    
    'Return the quarter dates
    GetQuarterDates = quarterDates
    
End Function




