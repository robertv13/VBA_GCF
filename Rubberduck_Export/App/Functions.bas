Attribute VB_Name = "Functions"
Option Explicit

Function GetID_FromInitials(i As String)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("dnrProf")
        If cell.Value2 = i Then
            GetID_FromInitials = cell.Offset(0, 1).value
            Exit Function
        End If
    Next cell

End Function

Function GetID_From_Client_Name(nomCLient As String) '2024-02-14 @ 06:07

    Dim ws As Worksheet, dynamicRange As Range
    On Error Resume Next
    Set ws = wshBD_Clients
    Set dynamicRange = ws.Range("dnrClients_All")
    On Error GoTo 0

    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'Clients' ou le DynamicRange 'dnrClients_All' n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly
    Dim result As Variant
    result = Application.WorksheetFunction.XLookup(nomCLient, _
        dynamicRange.Columns(1), dynamicRange.Columns(2), _
        "Not Found", 0, 1)
    
    If result <> "Not Found" Then
        GetID_From_Client_Name = result
    Else
        MsgBox "Impossible de retrouver la valeur dans la première colonne du client", vbExclamation
    End If
    
    'Free up memory - 2024-02-23
    Set ws = Nothing

End Function

Function Lookup_Data_In_A_Range(r As Range, cs As Long, ss As String, cr As Long) As Variant()
    
    'This function is used to retrieve information from a range
    'If found, it returns Address, Row number and the value for a specific column
    '2024-03-09 - First version
    
    Dim timerStart As Double: timerStart = Timer
    
    Dim foundInfo(1 To 3) As Variant 'Address, Row, Value
    Dim foundCell As Range
    Dim dataValue As Variant
    
    'Search for the string in a given range (r) at the column specified (cs)
    Set foundCell = r.Columns(cs).Find(What:=ss, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the string was found
    If Not foundCell Is Nothing Then
        'With the foundCell get the the address, the row number and the value
        foundInfo(1) = foundCell.Address
        foundInfo(2) = foundCell.row
        foundInfo(3) = foundCell.Offset(0, cr - cs).value 'Return Column - Searching column
        Lookup_Data_In_A_Range = foundInfo 'foundInfo is an array
    Else
        Lookup_Data_In_A_Range = foundInfo 'foundInfo is an array
    End If
    
    Set foundCell = Nothing
    
    Call Output_Timer_Results("Lookup_Data_In_A_Range()", timerStart)

End Function

Public Function Get_GL_Code_From_GL_Description(GLDescr As String) 'XLOOKUP - 2024-01-09 @ 09:19

    Dim dynamicRange As Range
    Dim result As Variant
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Admin")
    Set dynamicRange = ws.Range("dnrPlanComptableDescription")
    On Error GoTo 0
    
    If ws Is Nothing Or dynamicRange Is Nothing Then
        MsgBox "La feuille 'Admin' ou le DynamicRange n'a pas été trouvé!", _
            vbExclamation
        Exit Function
    End If
    
    'Using XLOOKUP to find the result directly
    result = Application.WorksheetFunction.XLookup(GLDescr, _
        dynamicRange.Columns(1), dynamicRange.Columns(2), _
        "Not Found", 0, 1)
    
    If result <> "Not Found" Then
        Get_GL_Code_From_GL_Description = result
    Else
        MsgBox "Impossible de retrouver la valeur dans la première colonne", vbExclamation
    End If

    'Free up memory - 2024-02-23
    Set ws = Nothing
    Set dynamicRange = Nothing

End Function

Public Function Get_TEC_Row_Number_By_TEC_ID(ByVal uniqueID As Variant, ByVal lookupRange As Range) As Long
    
    Dim matchResult As Variant

    'Use the Match function to find the row number of the unique TEC_ID
    matchResult = Application.Match(uniqueID, lookupRange.Columns(1), 0)
    matchResult = matchResult + 2 'Two header lines...

    'Check if Match found a result
    If Not IsError(matchResult) Then
        Get_TEC_Row_Number_By_TEC_ID = matchResult
    Else
        'If Match did not find a result, return 0
        Get_TEC_Row_Number_By_TEC_ID = 0
    End If
    
End Function

Public Function Validate_A_Date(paramDate As String) '2024-03-02 @ 08:04

    paramDate = Trim(paramDate)
    'User can enter / or - for separators
    Dim sDateDelimiter As String
    sDateDelimiter = "-"
    'Make sure that paramDate uses proper delimiter
    paramDate = Replace(paramDate, "/", sDateDelimiter)
    
    Dim sDate As String, isValidDate As Boolean
    sDate = ""
    isValidDate = False

    'Uses today's date as default
    Dim d, m, y As Integer
    d = Day(Now())
    m = Month(Now())
    y = Year(Now())

    Select Case Len(paramDate)
        Case 0                                       'Today's date
            sDate = Format(d, "00") & sDateDelimiter & Format(m, "00") & sDateDelimiter & Format(y, "0000")
        Case 1, 2                                    'Day only
            sDate = Format(paramDate, "00") & sDateDelimiter & Format(m, "00") & sDateDelimiter & Format(y, "0000")
        Case 3                                       'd/m only
            sDate = Format(Left(paramDate, 1), "00") & sDateDelimiter & Format(Mid(paramDate, 3, 1), "00") & sDateDelimiter & Format(y, "0000")
        Case 4                                       'd/mm or dd/m
            If InStr(paramDate, sDateDelimiter) = 2 Then
                sDate = Format(Left(paramDate, 1), "00") & sDateDelimiter & Format(Mid(paramDate, 3, 2), "00") & sDateDelimiter & Format(y, "0000")
            ElseIf InStr(paramDate, sDateDelimiter) = 3 Then
                sDate = Format(Left(paramDate, 2), "00") & sDateDelimiter & Format(Mid(paramDate, 4, 1), "00") & sDateDelimiter & Format(y, "0000")
            End If
        Case 5                                       'dd/mm
            If InStr(paramDate, sDateDelimiter) = 3 Then
                sDate = Format(Left(paramDate, 2), "00") & sDateDelimiter & Format(Mid(paramDate, 4, 2), "00") & sDateDelimiter & Format(y, "0000")
            End If
        Case 8                                       'd/m/yyyy or yy/mm/dd
            If Mid(paramDate, 2, 1) = sDateDelimiter And Mid(paramDate, 4, 1) = sDateDelimiter Then
                sDate = Format(Left(paramDate, 1), "00") & sDateDelimiter & Format(Mid(paramDate, 3, 1), "00") & sDateDelimiter & Format(Mid(paramDate, 5, 4), "0000")
            End If
            If Mid(paramDate, 3, 1) = sDateDelimiter And Mid(paramDate, 6, 1) = sDateDelimiter Then
                sDate = Format(Mid(paramDate, 7, 2), "00") & sDateDelimiter & Format(Mid(paramDate, 4, 2), "00") & sDateDelimiter & IIf(Left(paramDate, 2) >= 50, "19", "20") & Format(Left(paramDate, 2), "00")
            End If
            
        Case 9                                       'dd/m/yyyy or d/mm/yyyy
            If Mid(paramDate, 2, 1) = sDateDelimiter And Mid(paramDate, 5, 1) = sDateDelimiter Then
                sDate = Format(Left(paramDate, 1), "00") & sDateDelimiter & Format(Mid(paramDate, 3, 2), "00") & sDateDelimiter & Format(Mid(paramDate, 6, 4), "0000")
            End If
            If Mid(paramDate, 3, 1) = sDateDelimiter And Mid(paramDate, 5, 1) = sDateDelimiter Then
                sDate = Format(Left(paramDate, 2), "00") & sDateDelimiter & Format(Mid(paramDate, 4, 1), "00") & sDateDelimiter & Format(Mid(paramDate, 6, 4), "0000")
            End If
        Case 10                                      'dd/mm/yyyy or yyyy/mm/dd
            If Mid(paramDate, 3, 1) = sDateDelimiter And Mid(paramDate, 6, 1) = sDateDelimiter Then
                sDate = paramDate
            ElseIf Mid(paramDate, 5, 1) = sDateDelimiter And Mid(paramDate, 8, 1) = sDateDelimiter Then
                sDate = Mid(paramDate, 9, 2) & sDateDelimiter & Mid(paramDate, 6, 2) & sDateDelimiter & Left(paramDate, 4)
            End If
        Case Else
            sDate = ""
    End Select
        
    'Is the 'built' date valid ?
    isValidDate = IsDate(sDate)
    If isValidDate Then Validate_A_Date = sDate
        
End Function

Public Function IsDataValid() As Boolean

    IsDataValid = False
    
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

    IsDataValid = True

End Function

Public Function GetTaxRate(d As Date, taxType As String) As Double

    Dim row As Integer
    With wshAdmin
        For row = 18 To 11 Step -1
            If .Range("L" & row).value = taxType Then
                If d >= .Range("M" & row).value Then
                    GetTaxRate = .Range("N" & row).value
                    Exit For
                End If
            End If
        Next row
    End With
    
End Function

Function IsDateValide() As Boolean

    IsDateValide = False
    If wshGL_EJ.Range("K4").value = "" Or IsDate(wshGL_EJ.Range("K4").value) = False Then
        MsgBox "Une date d'écriture est obligatoire." & vbNewLine & vbNewLine & _
            "Veuillez saisir une date valide!", vbCritical, "Date Invalide"
    Else
        IsDateValide = True
    End If

End Function

Function IsEcritureBalance() As Boolean

    IsEcritureBalance = False
    If wshGL_EJ.Range("H26").value <> wshGL_EJ.Range("I26").value Then
        MsgBox "Votre écriture ne balance pas." & vbNewLine & vbNewLine & _
            "Débits = " & wshGL_EJ.Range("H26").value & " et Crédits = " & wshGL_EJ.Range("I26").value & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée.", vbCritical, "Veuillez vérifier votre écriture!"
    Else
        IsEcritureBalance = True
    End If

End Function

Function IsEcritureValide(rmax As Long) As Boolean

    IsEcritureValide = True 'Optimist
    If rmax <= 9 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        IsEcritureValide = False
    End If
    
    Dim i As Long
    For i = 9 To rmax
        If wshGL_EJ.Range("E" & i).value <> "" Then
            If wshGL_EJ.Range("H" & i).value = "" And wshGL_EJ.Range("I" & i).value = "" Then
                MsgBox "Il existe une ligne avec un compte, sans montant !"
                IsEcritureValide = False
            End If
        End If
    Next i

End Function

Public Function ClearRangeBorders(r As Range)

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

Public Function Pad_A_String(s As String, f As String, l As Integer, lr As String) As String

    Dim paddedString As String
    Dim charactersNeeded As Integer
    
    charactersNeeded = l - Len(s)
    
    If charactersNeeded > 0 Then
        If lr = "R" Then
            paddedString = s & String(charactersNeeded, f)
        Else
            paddedString = String(charactersNeeded, f) & s
        End If
    Else
        paddedString = s
    End If

    Pad_A_String = paddedString
        
End Function

