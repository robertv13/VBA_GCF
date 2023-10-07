Attribute VB_Name = "modJE"
Option Explicit

Sub Post_JE()

    Dim SumDt, SumCt As Currency
    SumDt = Range("G24").Value
    SumCt = Range("H24").Value
    If SumDt <> SumCt Then
        MsgBox "Votre écriture ne balance pas." & vbNewLine & vbNewLine & _
            "Débits = " & SumDt & " et Crédits = " & SumCt & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée.", vbCritical, "Veuillez vérifier votre écriture!"
        Exit Sub
    End If

    Dim EJRow As Long
    'Détermine la dernière ligne utilisée
    EJRow = wshJE.Range("D99999").End(xlUp).Row  'Last Used Row in wshJE
    If EJRow < 10 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        Exit Sub
    End If
    
    Dim TransRow, TransFirstRow As Long
    'Détermine la prochaine ligne disponible
    TransRow = wshGL.Range("C99999").End(xlUp).Row + 1  'First Empty Row in wshGL
    TransFirstRow = TransRow
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    Dim Ligne As Long
    With wshGL
        For Ligne = 9 To EJRow + 2
            .Range("C" & TransRow).Value = wshJE.Range("B1").Value
            .Range("D" & TransRow).Value = wshJE.Range("J4").Value
            .Range("E" & TransRow).Value = wshJE.Range("B1").Value
            .Range("F" & TransRow).Value = wshJE.Range("E4").Value
            If Ligne <= EJRow Then
                .Range("G" & TransRow).Value = wshJE.Range("D" & Ligne).Value
                .Range("H" & TransRow).Value = wshJE.Range("G" & Ligne).Value
                .Range("I" & TransRow).Value = wshJE.Range("H" & Ligne).Value
                .Range("J" & TransRow).Value = wshJE.Range("I" & Ligne).Value
            Else
                If Ligne = EJRow + 1 Then
                    .Range("G" & TransRow).Value = wshJE.Range("E6").Value
                End If
            End If
            .Range("K" & TransRow).Value = "=ROW()"
            TransRow = TransRow + 1
        Next Ligne
    End With
    'Les lignes subséquentes sont en police blanche...
    For Ligne = TransFirstRow + 1 To TransRow
        With wshGL.Range("D" & (TransFirstRow + 1) & ":F" & TransRow).Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
    Next Ligne
    
'    Dim r1 As Range
'    Set r1 = wshJE.Range("D118:J120")
'    r1.BorderAround LineStyle:=xlContinuous, Weight:=xlThick, Color:=vbBlack
    
    With wshJE
        'Increment Next JE number
        .Range("B1").Value = .Range("B1").Value + 1
        .Range("E4,J4,E6").ClearContents
        .Range("D9:D22,G9:G22,H9:H22,I9:I22").ClearContents
        .Range("E4").Activate
    End With
    
End Sub

Sub SetUpFrame(R As Range)
    
    With R
        .Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = vbBlack
            .TintAndShade = -1
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = vbBlack
            .TintAndShade = 0.5
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlHairline
        End With
    End With
End Sub
