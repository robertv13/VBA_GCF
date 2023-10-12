Attribute VB_Name = "modJE"
Option Explicit

Sub Post_JE()

    Dim SumDt, SumCt As Currency
    SumDt = Range("G25").Value
    SumCt = Range("H25").Value
    If SumDt <> SumCt Then
        MsgBox "Votre �criture ne balance pas." & vbNewLine & vbNewLine & _
            "D�bits = " & SumDt & " et Cr�dits = " & SumCt & vbNewLine & vbNewLine & _
            "Elle n'est donc pas report�e.", vbCritical, "Veuillez v�rifier votre �criture!"
        Exit Sub
    End If

    Dim LastEJRow As Long
    'D�termine la derni�re ligne utilis�e dans l'entr�e de journal
    LastEJRow = wshJE.Range("D99999").End(xlUp).row  'Last Used Row in wshJE
    If LastEJRow < 10 Or LastEJRow > 23 Then
        MsgBox "L'�criture est invalide !" & vbNewLine & vbNewLine & _
            "Elle n'est donc pas report�e!", vbCritical, "Vous devez v�rifier l'�criture"
        Exit Sub
    End If
    
    Dim TransRow, FirstTransRow As Long
    'D�termine la prochaine ligne disponible
    TransRow = wshGL.Range("C99999").End(xlUp).row + 1  'First Empty Row in wshGL
    FirstTransRow = TransRow
    
    'Transfert des donn�es vers wshGL, ent�te d'abord puis une ligne � la fois
    Dim Ligne As Long
    With wshGL
        For Ligne = 9 To LastEJRow + 2
            .Range("C" & TransRow).Value = wshJE.Range("B1").Value
            .Range("D" & TransRow).Value = wshJE.Range("J4").Value
            .Range("E" & TransRow).Value = wshJE.Range("B1").Value
            .Range("F" & TransRow).Value = wshJE.Range("E4").Value
            If Ligne <= LastEJRow Then
                .Range("G" & TransRow).Value = "1000"
                .Range("H" & TransRow).Value = wshJE.Range("D" & Ligne).Value
                .Range("I" & TransRow).Value = wshJE.Range("G" & Ligne).Value
                .Range("J" & TransRow).Value = wshJE.Range("H" & Ligne).Value
                .Range("K" & TransRow).Value = wshJE.Range("I" & Ligne).Value
            Else
                If Ligne = LastEJRow + 1 Then
                    .Range("H" & TransRow).Value = wshJE.Range("E6").Value
                End If
            End If
            .Range("L" & TransRow).Value = "=ROW()"
            TransRow = TransRow + 1
        Next Ligne
    End With
    'Les lignes subs�quentes sont en police blanche...
    With wshGL.Range("D" & (FirstTransRow + 1) & ":F" & (TransRow - 1)).Font
        .Color = vbWhite
    End With
    
    'Ajoute des bordures � l'entr�e de journal (ext�rieur)
    Dim r1 As Range
    Set r1 = wshGL.Range("D" & FirstTransRow & ":K" & (TransRow - 2))
    r1.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=vbBlack
    
    With wshJE
        'Increment Next JE number
        .Range("B1").Value = .Range("B1").Value + 1
        .Range("E4,J4,E6").ClearContents
        .Range("D9:D22,G9:G22,H9:H22,I9:I22").ClearContents
        .Range("E4").Activate
    End With
    
End Sub
