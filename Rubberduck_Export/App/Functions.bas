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

'    Dim lastRow As Long
'    lastRow = wshClientDB.Range("A99999").End(xlUp).row
    
    Dim ws As Worksheet, dynamicRange As Range
    On Error Resume Next
    Set ws = wshClientDB
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

Public Function Get_GL_Code_From_GL_Description(GLDescr As String) 'XLOOKUP - 2024-01-09 @ 09:19

    Dim dynamicRange As Range
    Dim result As Variant
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Admin")
    Set dynamicRange = ws.Range("dnrPlanComptable")
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

    'Check if Match found a result
    If Not IsError(matchResult) Then
        Get_TEC_Row_Number_By_TEC_ID = matchResult
    Else
        'If Match did not find a result, return 0
        Get_TEC_Row_Number_By_TEC_ID = 0
    End If
    
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


