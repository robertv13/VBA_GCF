Attribute VB_Name = "modGL_Stats_CA"
Option Explicit

Sub ActualiserStatsChiffreAffaires()

    Dim ws As Worksheet
    Set ws = wshGL_Stats_CA
    
    'Enlever toute note/commentaire existant
    Dim cell As Range
    For Each cell In ws.Range("D9:O9")
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete
        End If
    Next cell
    
    'Postes de revenus à considérer dans les REVENUS
    Dim glREV(1 To 2) As String
    Dim GL_Revenus_Consultation As String
    glREV(1) = Fn_NoCompteAPartirIndicateurCompte("Revenus de consultation")
    Dim GL_Revenus_TEC As String
    glREV(2) = Fn_NoCompteAPartirIndicateurCompte("Revenus - Travaux en cours")
    
    'Déterminer le dernier mois complété
    Dim moisPrécédent As Integer
    moisPrécédent = month(DateSerial(year(Date), month(Date), 0))
    Dim dateFinMoisPrécédent As Date
    dateFinMoisPrécédent = DateSerial(year(Date), month(Date), 0)
    Dim dateFinAnnée As Date
    dateFinAnnée = wsdADMIN.Range("AnneeA").Value
    
    'Doit-on insérer une nouvelle ligne (nouvelle année) ? 2025-08-03 @ 09:13
    Dim annéeCelluleC9 As Long
    annéeCelluleC9 = CLng(ws.Range("C9").Value)
    If year(dateFinAnnée) <> annéeCelluleC9 Then
        Call AjusterTableauNouvelleAnnee(ws, dateFinAnnée)
    End If
    
    Dim moisFinAnnéeFinancière As Integer
    moisFinAnnéeFinancière = wsdADMIN.Range("MoisFinAnnéeFinancière").Value
    
    'Mémoriser les colonnes de la feuille pour chacun des 12 mois de l'année financière
    Dim m As Integer
    Dim colMois(1 To 12, 1 To 2) As String
    Dim annee As Integer
    Dim anneeMoisDebutAF As Integer
    Dim anneeMoisFinAF As Integer
    
    Dim noMois As Integer
    Dim saveCol As Integer
    anneeMoisDebutAF = ws.Range("C9").Value - 1
    anneeMoisFinAF = ws.Range("C9").Value
    
    'Le premier mois de l'année financière est en colonne 4 du tableau
    Dim col As Integer
    col = 4
    For m = 1 To 12
        noMois = m + moisFinAnnéeFinancière
        If noMois <= 12 Then
            annee = anneeMoisDebutAF
        Else
            annee = anneeMoisFinAF
            noMois = noMois - 12
        End If
        colMois(m, 1) = col
        colMois(m, 2) = Format$(annee, "0000") & "-" & Format$(noMois, "00") & "-" & _
                Format$(day(DateSerial(annee, noMois + 1, 0)), "00")
        If noMois = month(Date) Then
            saveCol = col
        End If
        col = col + 1
    Next m
    
    Dim dateFinMois As Date
    Dim revenus_mois As Currency
    Dim revenus As Currency
    Dim revenus_TEC As Currency
    Dim r As Integer
    For m = 1 To 12
        col = colMois(m, 1)
        dateFinMois = colMois(m, 2)
        revenus = 0
        revenus_mois = 0
        For r = 1 To UBound(glREV, 1)
            revenus_mois = -Fn_TotalTransGLMois(glREV(r), dateFinMois)
            revenus = revenus + revenus_mois
        Next r
        ws.Cells(9, col).Value = revenus
    Next m
    
    'Variation des TEC - Quelle est la valeur des TEC et le solde au G/L des TEC ? '2025-02-21 @ 14:10
    Dim maxDate As Date
    Call modTEC_Evaluation.CalculerValeurTEC(Date, maxDate) 'Génère gDictHours
    
    Dim prof As Variant
    Dim profID As Long
    Dim strProf As String
    Dim tauxHoraire As Currency
    Dim tecValeur As Currency

    'Parcourir chacun des professionnels à partir de gDictHours
    For Each prof In gDictHours
        strProf = Mid$(prof, 4)
        profID = Fn_ProfIDAPartirDesInitiales(strProf)
        'Heures pour chacun des professionnels
        If gDictHours(prof)(0) <> 0 Then
            tauxHoraire = Fn_Get_Hourly_Rate(profID, Date)
'            Debug.Print prof, gDictHours(prof)(0), tauxHoraire, gDictHours(prof)(0) * tauxHoraire
            tecValeur = tecValeur + gDictHours(prof)(0) * tauxHoraire
        End If
    Next prof

    'Solde au G/L du compte Travaux en Cours en utilisant Fn_SoldesParCompteAvecADO - 2025-08-03 @ 09:23
    Dim glTEC As String
    glTEC = Fn_NoCompteAPartirIndicateurCompte("Travaux en cours")
    Dim dictSoldes As Object
    Set dictSoldes = CreateObject("Scripting.Dictionary")
    Set dictSoldes = modGL_Stuff.Fn_SoldesParCompteAvecADO(glTEC, "", Format$(Date, "yyyy-mm-dd"), True)
    Dim glTECSolde As Currency
    glTECSolde = dictSoldes(glTEC)

    'Ajoute un note à la cellule
    Dim rng As Range
    Set rng = ws.Cells(9, saveCol)
    rng.AddComment
    rng.Comment.text "Inclut un montant de " & Format$(tecValeur - glTECSolde, "###,##0.00 $") & vbLf & "pour la variation des TEC"
    With rng.Comment.Shape
        .Width = 135 'Largeur en points
        .Height = 27 'Hauteur en points
        .Top = rng.Top - 58 'Ajuste la position verticale
        .Left = rng.Left + 85 'Ajuste la position horizontale
    End With
    With rng.Comment.Shape.TextFrame.Characters.Font
        .Name = "Calibri"
        .size = 9
        .Bold = True
        .Italic = True
    End With
    rng.Value = ws.Cells(9, saveCol).Value + (tecValeur - glTECSolde)
    
    'Libérer les objets
    Set prof = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
End Sub

'@Description "Le tableau nécessite beaucoup d'ajustement lors de l'insertion d'une nouvelle année"
Sub AjusterTableauNouvelleAnnee(ws As Worksheet, dateFinAnnée As Date) '2025-08-03 @ 09:13

    With ws
        'Insérer une ligne à la position 9
        .Rows(9).Insert Shift:=xlDown
        .Rows(10).Copy
        .Rows(9).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        .Rows(25).Delete 'Maintenir 15 années d'historique
    End With
    
    ws.Range("C9").Value = year(dateFinAnnée)
    ws.Range("P9").formula = "=sum(D09:O09)"
    
    'Ajustement de la surbrillance pour les 3 dernières années
    With Range("C9:C24").Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Range("C10:C12").Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Ajustement du libellé de la cellule C26
    ws.Range("C26").Value = "Moyennes mensuelles pour " & year(dateFinAnnée) - 3 & " @ " & year(dateFinAnnée) - 1
    
    'Ajustement des formules de la ligne 27
    Dim col As Integer
    For col = 4 To 16 'Colonnes D à P
        ws.Cells(27, col).formula = "=AVERAGE(" & _
            ws.Cells(10, col).Address(False, False) & ":" & _
            ws.Cells(12, col).Address(False, False) & ")"
    Next col
    
    'Ajout de l'année précédente, comme complète, à la ligne 31
    With ws
        .Rows(31).Insert Shift:=xlDown
        .Rows(32).Copy
        .Rows(31).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        .Rows(36).Delete 'Garder que 5 années pour la pondération
    End With
    ws.Range("C31").Value = year(dateFinAnnée) - 1
    
    'Remplir les cellules (D à P) de la nouvelle ligne (31) avec formule pour pondération
    For col = 4 To 16
        ws.Cells(31, col).formula = "=" & ws.Cells(10, col).Address(False, False) & "/$P10"
    Next col
    
    'Ajuster les formules des cellules (D à P) de la ligne (37)
    For col = 4 To 16
        ws.Cells(37, col).formula = "=AVERAGE(" & _
            ws.Cells(31, col).Address(False, False) & ":" & _
            ws.Cells(35, col).Address(False, False) & ")"
    Next col
    
    'Ajustement de la MÉGA formule dans la cellule d'estimation ANNUELLE
    Dim cell As Range
    Set cell = Range("P4")
    Dim formuleActuelle As String
    formuleActuelle = cell.formula
    formuleActuelle = Replace(formuleActuelle, "10", "9")
    cell.formula = formuleActuelle

    'Libérer la mémoire
    Set cell = Nothing
    
End Sub

Sub shpRetournerAuMenu_Click()

    Call RetournerAuMenu

End Sub

Sub RetournerAuMenu()

    Call modAppli.QuitterFeuillePourMenu(wshMenuGL, True) '2025-08-19 @ 07:22

End Sub
