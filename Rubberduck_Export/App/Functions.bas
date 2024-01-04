Attribute VB_Name = "Functions"
Option Explicit

Function GetID_FromInitials(i As String)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("Prof_Initiales")
        If cell.Value2 = i Then
            GetID_FromInitials = cell.Offset(0, 1).value
        End If
    Next cell

End Function

Function GetID_FromClientName(ClientNom As String)

    Dim lastRow As Long
    lastRow = wshClientDB.Range("A99999").End(xlUp).row
    
    Dim i As Long
    For i = 1 To lastRow
        If wshClientDB.Cells(i, 2) = ClientNom Then
            'Debug.Print "ID du client - '" & wshClientDB.Cells(i, 1).value & "'"
            GetID_FromClientName = wshClientDB.Cells(i, 1).value
            Exit Function
        End If
    Next i

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


