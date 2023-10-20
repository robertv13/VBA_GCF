Attribute VB_Name = "modJE"
Option Explicit

Sub JE_Post()

    If IsDateValide = False Then Exit Sub
    
    If IsEcritureBalance = False Then Exit Sub
    
    Dim RowEJLast As Long
    RowEJLast = wshJE.Range("D99").End(xlUp).Row  'Last Used Row in wshJE
    If IsEcritureValide(RowEJLast) = False Then Exit Sub
    
    Dim rowGLTrans, rowGLTransFirst As Long
    'Détermine la prochaine ligne disponible
    rowGLTrans = wshGL.Range("C99999").End(xlUp).Row + 1  'First Empty Row in wshGL
    rowGLTransFirst = rowGLTrans
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    FromJE2GL RowEJLast, rowGLTrans

    'Les lignes subséquentes sont en police blanche...
    With wshGL.Range("D" & (rowGLTransFirst + 1) & ":F" & (rowGLTrans - 1)).Font
        .Color = vbWhite
    End With
    
    'Ajoute des bordures à l'entrée de journal (extérieur)
    Dim r1 As Range
    Set r1 = wshGL.Range("D" & rowGLTransFirst & ":K" & (rowGLTrans - 2))
    r1.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=vbBlack
    
    With wshGL.Range("H" & (rowGLTrans - 2) & ":K" & (rowGLTrans - 2))
        .Font.Italic = True
        .Font.Bold = True
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
    End With
    
    If wshJE.ckbRecurrente = True Then
        SaveEJRecurrente RowEJLast
    End If
    
    With wshJE
        'Increment Next JE number
        .Range("B1").Value = .Range("B1").Value + 1
        Call wshJEClearAllCells
        .Range("E4").Activate
    End With
    
End Sub

Sub FromJE2GL(rEJLast As Long, ByRef rGLTrans)

    Dim l As Long
    With wshGL
        For l = 9 To rEJLast + 2
            .Range("C" & rGLTrans).Value = wshJE.Range("B1").Value
            .Range("D" & rGLTrans).Value = Format(CDate(wshJE.Range("J4").Value), "dd/mm/yyyy")
            .Range("E" & rGLTrans).Value = wshJE.Range("B1").Value
            .Range("F" & rGLTrans).Value = wshJE.Range("E4").Value
            If l <= rEJLast Then
                .Range("G" & rGLTrans).Value = wshJE.Range("K" & l).Value
                .Range("H" & rGLTrans).Value = wshJE.Range("D" & l).Value
                .Range("I" & rGLTrans).Value = wshJE.Range("G" & l).Value
                .Range("J" & rGLTrans).Value = wshJE.Range("H" & l).Value
                .Range("K" & rGLTrans).Value = wshJE.Range("I" & l).Value
            Else
                If l = rEJLast + 1 Then
                    .Range("H" & rGLTrans).Value = wshJE.Range("E6").Value
                End If
            End If
            .Range("L" & rGLTrans).Value = "=ROW()"
            rGLTrans = rGLTrans + 1
        Next l
    End With

End Sub

Sub SaveEJRecurrente(ll As Long)

    Dim EJAutoNo As Long
    EJAutoNo = wshJERecurrente.Range("B1").Value
    wshJERecurrente.Range("B1").Value = wshJERecurrente.Range("B1").Value + 1
    
    Dim rowEJAuto, rowEJAutoSave As Long
    rowEJAuto = wshJERecurrente.Range("D99999").End(xlUp).Row + 3 'First available Row in wshJERecurrente
    rowEJAutoSave = rowEJAuto
    
    Dim r As Integer
    For r = 9 To ll
        wshJERecurrente.Range("C" & rowEJAuto).Value = EJAutoNo
        wshJERecurrente.Range("D" & rowEJAuto).Value = wshJE.Range("K" & r).Value
        wshJERecurrente.Range("E" & rowEJAuto).Value = wshJE.Range("D" & r).Value
        wshJERecurrente.Range("F" & rowEJAuto).Value = wshJE.Range("G" & r).Value
        wshJERecurrente.Range("G" & rowEJAuto).Value = wshJE.Range("H" & r).Value
        wshJERecurrente.Range("H" & rowEJAuto).Value = wshJE.Range("I" & r).Value
        wshJERecurrente.Range("I" & rowEJAuto).Value = "=ROW()"
        rowEJAuto = rowEJAuto + 1
    Next
    'Ligne de description
    wshJERecurrente.Range("C" & rowEJAuto).Value = EJAutoNo
    wshJERecurrente.Range("E" & rowEJAuto).Value = wshJE.Range("E6").Value
    wshJERecurrente.Range("I" & rowEJAuto).Value = "=ROW()"
    rowEJAuto = rowEJAuto + 1
    'Ligne vide
    wshJERecurrente.Range("C" & rowEJAuto).Value = EJAutoNo
    wshJERecurrente.Range("I" & rowEJAuto).Value = "=ROW()"
    rowEJAuto = rowEJAuto + 1
    
    'Ajoute la description dans la liste des E/J automatiques (K1:L99999)
    Dim rowEJAutoDesc As Long
    rowEJAutoDesc = wshJERecurrente.Range("K99999").End(xlUp).Row + 1 'First available Row in wshJERecurrente
    wshJERecurrente.Range("K" & rowEJAutoDesc).Value = wshJE.Range("E6").Value
    wshJERecurrente.Range("L" & rowEJAutoDesc).Value = EJAutoNo

    'Ajoute des bordures à l'entrée de journal récurrente
    Dim r1 As Range
    Set r1 = wshJERecurrente.Range("D" & rowEJAutoSave & ":H" & (rowEJAuto - 2))
    r1.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=vbBlack

End Sub

Sub LoadJEAutoIntoJE(EJAutoDesc As String, NoEJAuto As Long)

    'On copie l'E/J automatique vers wshEJ
    Dim rowJEAuto, rowJE As Long
    rowJEAuto = wshJERecurrente.Range("C99999").End(xlUp).Row  'Last Row used in wshJERecuurente
    
    Call wshJEClearAllCells
    rowJE = 9
    
    Dim r As Long
    For r = 2 To rowJEAuto
        If wshJERecurrente.Range("C" & r).Value = NoEJAuto And wshJERecurrente.Range("D" & r).Value <> "" Then
            wshJE.Range("D" & rowJE).Value = wshJERecurrente.Range("E" & r).Value
            wshJE.Range("G" & rowJE).Value = wshJERecurrente.Range("F" & r).Value
            wshJE.Range("H" & rowJE).Value = wshJERecurrente.Range("G" & r).Value
            wshJE.Range("I" & rowJE).Value = wshJERecurrente.Range("H" & r).Value
            wshJE.Range("K" & rowJE).Value = wshJERecurrente.Range("D" & r).Value
            rowJE = rowJE + 1
        End If
    Next r
    wshJE.Range("E6").Value = "Auto - " & EJAutoDesc
    wshJE.Range("J4").Activate

End Sub

Sub wshJEClearAllCells()

    'Efface toutes les cellules de la feuille
    With wshJE
        .Range("E4,J4,E6:J6").ClearContents
        .Range("D9:F22,G9:G22,H9:H22,I9:J22,K9:K22").ClearContents
        .ckbRecurrente = False
    End With

End Sub

Sub BuildDate(cell As String, r As Range)
        Dim d, m, y As Integer
        Dim strDateJour, strDateConsruite As String
        Dim dateValide As Boolean
        dateValide = True

        cell = Replace(cell, "/", "")
        cell = Replace(cell, "-", "")

        'Utilisation de la date du jour
        strDateJour = Now
        d = Day(strDateJour)
        m = Month(strDateJour)
        y = Year(strDateJour)

        Select Case Len(cell)
            Case 0
                strDateConsruite = Format(d, "00") & "/" & Format(m, "00") & "/" & Format(y, "0000")
            Case 1, 2
                strDateConsruite = Format(cell, "00") & "/" & Format(m, "00") & "/" & Format(y, "0000")
            Case 3
                strDateConsruite = Format(Left(cell, 1), "00") & "/" & Format(Mid(cell, 2, 2), "00") & "/" & Format(y, "0000")
            Case 4
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(y, "0000")
            Case 6
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & "20" & Format(Mid(cell, 5, 2), "00")
            Case 8
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(Mid(cell, 5, 4), "0000")
            Case Else
                dateValide = False
        End Select
        dateValide = IsDate(strDateConsruite)

    If dateValide Then
        r.Value = Format(strDateConsruite, "dd/mm/yyyy")
    Else
        MsgBox "La saisie est invalide...", vbInformation, "Il est impossible de construire une date"
    End If

End Sub

Function IsDateValide() As Boolean

    IsDateValide = False
    If wshJE.Range("J4").Value = "" Or IsDate(wshJE.Range("J4").Value) = False Then
        MsgBox "Une date d'écriture est obligatoire." & vbNewLine & vbNewLine & _
            "Veuillez saisir une date valide!", vbCritical, "Date Invalide"
    Else
        IsDateValide = True
    End If

End Function

Function IsEcritureBalance() As Boolean

    IsEcritureBalance = False
    If wshJE.Range("G25").Value <> wshJE.Range("H25").Value Then
        MsgBox "Votre écriture ne balance pas." & vbNewLine & vbNewLine & _
            "Débits = " & wshJE.Range("G25").Value & " et Crédits = " & wshJE.Range("H25").Value & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée.", vbCritical, "Veuillez vérifier votre écriture!"
    Else
        IsEcritureBalance = True
    End If

End Function

Function IsEcritureValide(rmax As Long) As Boolean

    IsEcritureValide = False
    If rmax < 10 Or rmax > 23 Then
        MsgBox "L'écriture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas reportée!", vbCritical, "Vous devez vérifier l'écriture"
        IsEcritureValide = False
    Else
        IsEcritureValide = True
    End If

End Function
