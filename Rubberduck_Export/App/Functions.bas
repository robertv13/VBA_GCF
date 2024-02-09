Attribute VB_Name = "Functions"
Option Explicit

Function GetID_FromInitials(i As String)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("dnrProf")
        If cell.Value2 = i Then
            GetID_FromInitials = cell.Offset(0, 1).value
        End If
    Next cell

End Function

Function GetID_FromClientName(ClientNom As String)

    Dim LastRow As Long
    LastRow = wshClientDB.Range("A99999").End(xlUp).row
    
    Dim i As Long
    For i = 1 To LastRow
        If wshClientDB.Cells(i, 2) = ClientNom Then
            'Debug.Print "ID du client - '" & wshClientDB.Cells(i, 1).value & "'"
            GetID_FromClientName = wshClientDB.Cells(i, 1).value
            Exit Function
        End If
    Next i

End Function

Public Function GetAccountNoFromDescription(GLDescr As String) 'XLOOKUP - 2024-01-09 @ 09:19

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
        GetAccountNoFromDescription = result
    Else
        MsgBox "Impossible de retrouver la valeur dans la première colonne", vbExclamation
    End If

End Function

Public Function IsDataValid() As Boolean

    IsDataValid = False
    
    'Validations first (one field at a time)
    If frmSaisieHeures.cmbProfessionnel.value = "" Then
        MsgBox Prompt:="Le professionnel est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.cmbProfessionnel.SetFocus
        Exit Function
    End If

    If frmSaisieHeures.txtDate.value = "" Or IsDate(frmSaisieHeures.txtDate.value) = False Then
        MsgBox Prompt:="La date est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.txtDate.SetFocus
        Exit Function
    End If

    If frmSaisieHeures.txtClient.value = "" Then
        MsgBox Prompt:="Le client est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.txtClient.SetFocus
        Exit Function
    End If
    
    If frmSaisieHeures.txtHeures.value = "" Or IsNumeric(frmSaisieHeures.txtHeures.value) = False Then
        MsgBox Prompt:="Le nombre d'heures est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.txtHeures.SetFocus
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


