Attribute VB_Name = "Module1"
'@Folder(Constat_Mensuel.Duplicate)

Option Explicit

Sub Main() 'Création de la feuille suivante (Mois Suivant)
    
    Application.ScreenUpdating = False
    
    Dim NbSheets As Long
    NbSheets = ThisWorkbook.Worksheets.Count
    
    'Noter la période des formules ET établir la nouvelle période des formules
    Dim nouvellePeriode As String, anciennePeriode As String
    nouvellePeriode = ActiveSheet.Range("C4").Formula
    nouvellePeriode = Mid$(nouvellePeriode, 3, 8)
    anciennePeriode = "'" & Mid$(nouvellePeriode, 1, 7) & "'!"
    
    Dim Mois As Long, Annee As Long
    Annee = Year(ActiveSheet.Range("B4").Value)
    Mois = Month(ActiveSheet.Range("B4").Value)
    Mois = Mois + 1
    If Mois > 12 Then
        Annee = Annee + 1
        Mois = Mois - 12
    End If

    nouvellePeriode = "'" & Format$(Annee, "0000") & "-" & Format$(Mois, "00") & "'!"
    
    'On prend l'onglet de base (à copier)
    Dim NomOnglet As String
    NomOnglet = ActiveSheet.Name
    
    'Jour = 1, Détermine le nom du nouvel onglet
    Annee = CInt(Mid$(NomOnglet, 1, 4))
    Mois = CInt(Mid$(NomOnglet, 6, 2))
    Mois = Mois + 1
    If Mois > 12 Then
        Annee = Annee + 1
        Mois = Mois - 12
    End If
    
    Dim NouveauMois As String
    NouveauMois = DateSerial(Annee, Mois, 1)
        
    'Clone de l'onglet courant
    Sheets(NbSheets).Copy After:=Sheets(NbSheets)
    Sheets(NbSheets + 1).Name = Year(NouveauMois) & "-" & Format$(Month(NouveauMois), "00")
    
    'Dernier jour du mois précédent
    Dim DDJMoisPrecedent As Date
    DDJMoisPrecedent = DateSerial(Annee, Mois, 1) - 1
    ActiveSheet.Range("B4").Value = DDJMoisPrecedent
    
    'Détermine la date du dernier jour du nouveau mois
    Dim DernierJour As Double
    DernierJour = Application.WorksheetFunction.EoMonth(ActiveSheet.Range("B4").Value, "1")
    
    'Efface le contenu de certaines cellules du prochain mois
    
    Call InitialiseProchainMois(anciennePeriode, nouvellePeriode, Mois)
    
    Call RemplirTableau(VBA.Format$(DernierJour, "dd/mm/yyyy"))
    
    Dim rng As Range
    Set rng = ActiveSheet.Range("E5:E35,I5:I35,N5:N35,P5:P35")
    Call Set_Formats_Totals_Columns(rng)
    
    Application.ScreenUpdating = True
    
    'Clear the empty lines
    Dim i As Long
    For i = 5 To 35
        If ActiveSheet.Range("A" & i).Value = "" Then
            ActiveSheet.Rows(i).Clear
        End If
    Next i
    
    MsgBox "Le nouvel onglet (" & Sheets(NbSheets + 1).Name & ") a été créé " & _
                                                            "avec succès"

End Sub

Private Sub InitialiseProchainMois(base As String, nouveau As String, m As Long)

    'Remplace les formules pour le nouvel onglet (mois)
    ActiveSheet.Range("C4:P4, X4").Replace What:=base, Replacement:=nouveau, _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    'Enlève les entrées/sorties des $$$ pour le mois
    ActiveSheet.Range("C38:D38, G38:H38, K38:L38").ClearContents
    
    If m <> 1 Then 'Tous les mois, SAUF le premier de l'année
        ActiveSheet.Range("C41").Formula = "=" & nouveau & "C41+C37"
        ActiveSheet.Range("D41").Formula = "=" & nouveau & "D41+D37"
        ActiveSheet.Range("G41").Formula = "=" & nouveau & "G41+G37"
        ActiveSheet.Range("H41").Formula = "=" & nouveau & "H41+H37"
        ActiveSheet.Range("K41").Formula = "=" & nouveau & "K41+K37"
        ActiveSheet.Range("L41").Formula = "=" & nouveau & "L41+L37"
        
        ActiveSheet.Range("C42").Formula = "=" & nouveau & "C42+C38"
        ActiveSheet.Range("D42").Formula = "=" & nouveau & "D42+D38"
        ActiveSheet.Range("G42").Formula = "=" & nouveau & "G42+G38"
        ActiveSheet.Range("H42").Formula = "=" & nouveau & "H42+H38"
        ActiveSheet.Range("K42").Formula = "=" & nouveau & "K42+K38"
        ActiveSheet.Range("L42").Formula = "=" & nouveau & "L42+L38"
    Else
        'Changement de valeur pour l'année
        ActiveSheet.Range("C41").Formula = "=C37"
        ActiveSheet.Range("D41").Formula = "=D37"
        ActiveSheet.Range("G41").Formula = "=G37"
        ActiveSheet.Range("H41").Formula = "=H37"
        ActiveSheet.Range("K41").Formula = "=K37"
        ActiveSheet.Range("L41").Formula = "=L37"
        'Retraits/Dépôts pour l'année
        ActiveSheet.Range("C42").Formula = "=C38"
        ActiveSheet.Range("D42").Formula = "=D38"
        ActiveSheet.Range("G42").Formula = "=G38"
        ActiveSheet.Range("H42").Formula = "=H38"
        ActiveSheet.Range("K42").Formula = "=K38"
        ActiveSheet.Range("L42").Formula = "=L38"
    End If
    
    ActiveSheet.Range("C5:C35").ClearContents
    ActiveSheet.Range("G5:H35").ClearContents
    ActiveSheet.Range("K5:L35").ClearContents
    ActiveSheet.Range("X5:X35").ClearContents
    
    ActiveSheet.Range("C5:AA35").ClearNotes

    'Toutes les cellules sont en mode normal
    ActiveSheet.Range("A5:AA35").Font.Italic = False
    ActiveSheet.Range("A5:AA35").Font.Bold = False
    
    'Affiche toutes les lignes masquées
    ActiveSheet.Range("5:15").EntireRow.Hidden = False

End Sub

Private Sub RemplirTableau(dt As Date)

    Dim MoisCourant As String, PremiereLigne As Long
                
    MoisCourant = Year(dt) & "-" & Format$(Month(dt), "00")
    
    Dim i As Long, row As Long
    row = 35
    For i = 35 To 5 Step -1
        If Year(dt) & "-" & Format$(Month(dt), "00") <> MoisCourant Then
            Exit For
        End If
        If Weekday(dt) <> 1 And Weekday(dt) <> 7 Then
            'On a une ligne normale (lundi au vendredi)
            Call ConstruireLigneNormale(row, dt)
            PremiereLigne = row
            row = row - 1
        End If
        If Weekday(dt) = 2 Then 'On a une ligne 'lundi', on insère une ligne vide avant
            Call ConstruireLigneVide(row)
            row = row - 1
        End If
        dt = DateAdd("d", -1, dt)
    Next i
    
    'Masquer les lignes flottantes au début du tableau
    For i = 6 To row
        ActiveSheet.Rows(i).EntireRow.Hidden = True
    Next i
    
    'Ajustement de la première journée du mois pour les colonnes _
                                        'D', 'M', 'R', 'S' et 'U'
    Dim FormuleAAjuster As String
    
    'Colonne D - REER CSN
    ActiveSheet.Range("D" & PremiereLigne).Formula = "=D4"
    
    'Colonne E - Ajustement de la mise en forme conditionnelle
    With ActiveSheet.Range("E" & PremiereLigne)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=E4"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Colonne I - Ajustement de la mise en forme conditionnelle
    With ActiveSheet.Range("I" & PremiereLigne)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=I4"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Colonne M - 9299-2585 - Cash
    ActiveSheet.Range("M" & PremiereLigne).Formula = "=M4"

    'Colonne N - Ajustement de la mise en forme conditionnelle
    With ActiveSheet.Range("N" & PremiereLigne)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=N4"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Colonne P - Ajustement de la mise en forme conditionnelle
    With ActiveSheet.Range("P" & PremiereLigne)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=P4"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    Dim s1 As String, s2 As String, jourSem As String
    jourSem = Weekday(ActiveSheet.Range("A" & PremiereLigne).Value)
    
    'Colonne R
    FormuleAAjuster = ActiveSheet.Range("R" & PremiereLigne).Formula
    If jourSem <> "2" Then
        s1 = "P" & (PremiereLigne - 1)
    Else
        s1 = "P" & (PremiereLigne - 2)
    End If
    s2 = "P4"
    FormuleAAjuster = Replace(FormuleAAjuster, s1, s2)
    ActiveSheet.Range("R" & PremiereLigne).Formula = FormuleAAjuster
    
    'Colonne S
    FormuleAAjuster = ActiveSheet.Range("S" & PremiereLigne).Formula
    If jourSem <> "2" Then
        s1 = "P" & (PremiereLigne - 1)
    Else
        s1 = "P" & (PremiereLigne - 2)
    End If
    s2 = "P4"
    FormuleAAjuster = Replace(FormuleAAjuster, s1, s2)
    ActiveSheet.Range("S" & PremiereLigne).Formula = FormuleAAjuster
   
    'Colonne Y
    FormuleAAjuster = ActiveSheet.Range("Y" & PremiereLigne).Formula
    If jourSem <> "2" Then
        s1 = "X" & (PremiereLigne - 1)
    Else
        s1 = "X" & (PremiereLigne - 2)
    End If
    s2 = "X4"
    FormuleAAjuster = Replace(FormuleAAjuster, s1, s2)
    ActiveSheet.Range("Y" & PremiereLigne).Formula = FormuleAAjuster

End Sub

Private Sub ConstruireLigneNormale(r As Long, d As Date)

    Dim j As Long
    j = Weekday(d)
    ActiveSheet.Range("B" & r).Value = d
    ActiveSheet.Range("A" & r).Formula = "=B" & r
    ActiveSheet.Range("A" & r).NumberFormat = "ddd"

    'Détermine la ligne de référence
    Dim RowOffset As Long
    If j = vbMonday Then
        RowOffset = 2
    Else
        RowOffset = 1
    End If
    
    '2 cellules copient toujours la valeur de la veille...
    ActiveSheet.Range("D" & r).Formula = "=D" & (r - RowOffset)
    ActiveSheet.Range("M" & r).Formula = "=M" & (r - RowOffset)
    
    'Cellule E - Sous-Total REER = Colonne C + D (Formule & Conditional Formatting)
    With ActiveSheet.Range("E" & r)
        .Formula = "=SUM(RC[-2]:RC[-1])"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=E" & (r - RowOffset)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Cellule I - Sous-Total HORS REER & CELI = Colonne G + H (Somme & CondFormatting)
    With ActiveSheet.Range("I" & r)
        .Formula = "=SUM(RC[-2]:RC[-1])"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=I" & (r - RowOffset)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Cellule N - Sous-Total 9299-2585 Québec inc. - Somme & Cond. Formatting
    With ActiveSheet.Range("N" & r)
        .Formula = "=SUM(RC[-3]:RC[-1])"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=N" & (r - RowOffset)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Cellule P - Grand Total de tous les placements (Somme & Format)
    With ActiveSheet.Range("P" & r)
        .Formula = "=SUM(RC[-11]+RC[-7]+RC[-2])"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=P" & (r - RowOffset)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Colonne 'R' - Variation du jour en $ (calcul)
    If j <> vbMonday Then
        ActiveSheet.Range("R" & r).Formula = "=IF(RC[-15]<>"""",RC[-2]-R[-1]C[-2],"""")"
    Else
        ActiveSheet.Range("R" & r).Formula = "=IF(RC[-15]<>"""",RC[-2]-R[-2]C[-2],"""")"
    End If
    
    'Cellule 'S' - Variation du jour en % (calcul)
    With ActiveSheet.Range("S" & r)
        If j <> vbMonday Then
            .Formula = "=IF(RC[-16]<>"""",RC[-1]/R[-1]C[-3],"""")"
        Else
            .Formula = "=IF(RC[-16]<>"""",RC[-1]/R[-2]C[-3],"""")"
        End If
        
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=S" & (r - RowOffset)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Cellule 'U' - Variation du mois en $ (calcul)
    With ActiveSheet.Range("U" & r)
        .Formula = "=IF(RC[-18]<>"""",RC[-5]-R4C[-5],"""")"
    
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=U" & (r - RowOffset)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
    'Cellule 'V' - Variation du mois en % (calcul)
    With ActiveSheet.Range("V" & r)
        .Formula = "=IF(RC[-19]<>"""",RC[-1]/R4C[-6],"""")"
        
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                              Formula1:="=V" & (r - RowOffset)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.Color = vbRed
        .FormatConditions(1).Font.TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
 
    'Cellule 'Y' - Variation de la journée (TSX) en pourcentage
    With ActiveSheet.Range("Y" & r)
        If j <> vbMonday Then
            .Formula = "=IF(RC[-1]<>"""",(RC[-1]-r[-1]c[-1])/R[-1]C[-1],"""")"
        Else
            .Formula = "=IF(RC[-1]<>"""",(RC[-1]-R[-2]c[-1])/R[-2]C[-1],"""")"
        End If

        .FormatConditions.Delete
    End With

    'Colonne 'Z' - Variation de la semaine (TSX) en pourcentage
    If j = vbFriday Then
        If r <= 11 Then
            ActiveSheet.Range("Z" & r).Formula = "=IF(RC[-2]<>"""",(RC[-2]-R4C[-2])/R4C[-2],"""")"
        Else
            ActiveSheet.Range("Z" & r).Formula = "=IF(RC[-2]<>"""",(RC[-2]-R[-6]C[-2])/R[-6]C[-2],"""")"
        End If
        ActiveSheet.Range("AA" & r).Formula = "=IF(RC[-3]<>"""",(RC[-3]-R4C[-3])/R4C[-3],"""")"
    Else
        ActiveSheet.Range("Z" & r).Formula = ""
    End If

    'Format numérique, sans cents avec séparateur de milliers, noir
    ActiveSheet.Range("C" & r & ":R" & r).NumberFormat = "#,##0"
    ActiveSheet.Range("P" & r).Font.Color = vbBlack
    ActiveSheet.Range("R" & r).Font.Color = vbBlack
    
End Sub

Private Sub ConstruireLigneVide(r As Long)

    ActiveSheet.Range("A" & r & ":AA" & r).Clear

End Sub

Private Sub Set_Formats_Totals_Columns(r As Range)

    With r
        .Font.Bold = True
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.15
            .PatternTintAndShade = 0
        End With
    End With

End Sub


