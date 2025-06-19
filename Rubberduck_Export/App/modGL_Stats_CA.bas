Attribute VB_Name = "modGL_Stats_CA"
Option Explicit

Sub shp_GL_PrepEF_Actualiser_Click()

    Dim ws As Worksheet
    Set ws = wshGL_PrepEF
    
    Call Actualiser_Stats_CA
    
End Sub

Sub Actualiser_Stats_CA()

    Dim ws As Worksheet
    Set ws = wshGL_Stats_CA
    
    Dim cell As Range
    For Each cell In ws.Range("D9:O9")
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete
        End If
    Next cell
    
    'Postes de revenus à considérer dans les REVENUS
    Dim glREV(1 To 2) As String
    Dim GL_Revenus_Consultation As String
    glREV(1) = ObtenirNoGlIndicateur("Revenus de consultation")
    Dim GL_Revenus_TEC As String
    glREV(2) = ObtenirNoGlIndicateur("Revenus - Travaux en cours")
    
    'Déterminer le dernier mois complété
    Dim moisPrécédent As Integer
    moisPrécédent = month(DateSerial(year(Date), month(Date), 0))
    Dim dateFinMoisPrécédent As Date
    dateFinMoisPrécédent = DateSerial(year(Date), month(Date), 0)
    
    Dim moisFinAnnéeFinancière As Integer
    moisFinAnnéeFinancière = wsdADMIN.Range("MoisFinAnnéeFinancière").Value
    
    'Mémoriser les colonnes de la feuille pour chacun des 12 mois de l'année financière
    Dim colMois(1 To 12, 1 To 2) As String
    Dim annee As Integer, anneeMoisDebutAF As Integer, anneeMoisFinAF As Integer
    
    Dim m As Integer, noMois As Integer, col As Integer, saveCol As Integer
    anneeMoisDebutAF = ws.Range("C9").Value - 1
    anneeMoisFinAF = ws.Range("C9").Value
    'Le premier mois de l'année financière est en colonne 4 du tableau
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
'        Debug.Print m, noMois, col, annee, colMois(m, 2)
        If noMois = month(Date) Then
            saveCol = col
        End If
        col = col + 1
    Next m
    
    Dim dateFinMois As Date
    Dim revenus_mois As Currency, revenus As Currency, revenus_TEC As Currency
    Dim r As Integer
    For m = 1 To 12
        col = colMois(m, 1)
        dateFinMois = colMois(m, 2)
        revenus = 0
        revenus_mois = 0
        For r = 1 To UBound(glREV, 1)
            revenus_mois = -Fn_Get_GL_Month_Trans_Total(glREV(r), dateFinMois)
            revenus = revenus + revenus_mois
        Next r
        ws.Cells(9, col).Value = revenus
    Next m
    
    'Variation des TEC - Quelle est la valeur des TEC et le solde au G/L des TEC ? '2025-02-21 @ 14:10
    Dim maxDate As Date
    Call TEC_Evaluation_Calcul(Date, maxDate) 'Génère gDictHours
    
    Dim prof As Variant
    Dim profID As Long
    Dim strProf As String
    Dim tauxHoraire As Currency
    Dim tecValeur As Currency

    'Parcourir chacun des professionnels à partir de gDictHours
    For Each prof In gDictHours
        strProf = Mid$(prof, 4)
        profID = Fn_GetID_From_Initials(strProf)
        'Heures pour chacun des professionnels
        If gDictHours(prof)(0) <> 0 Then
            tauxHoraire = Fn_Get_Hourly_Rate(profID, Date)
'            Debug.Print prof, gDictHours(prof)(0), tauxHoraire, gDictHours(prof)(0) * tauxHoraire
            tecValeur = tecValeur + gDictHours(prof)(0) * tauxHoraire
        End If
    Next prof

    'Solde au G/L du compte Travaux en Cours
    Dim glTEC As String
    Dim glTECSolde As Currency
    glTEC = ObtenirNoGlIndicateur("Travaux en cours")
    glTECSolde = Fn_Get_GL_Account_Balance(glTEC, maxDate)

    'Ajoute un note à la cellule
    Dim rng As Range
    Set rng = ws.Cells(9, saveCol)
    rng.AddComment
    rng.Comment.Text "Inclut un montant de " & Format$(tecValeur - glTECSolde, "###,##0.00 $") & vbLf & "pour la variation des TEC"
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
    
'    Debug.Print "777 - " & tecValeur, glTECSolde, tecValeur - glTECSolde
    
    'Libérer les objets
    Set prof = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
End Sub

Sub shp_GL_Stats_CA_Exit_Click()

    Call GL_Stats_CA_Back_To_Menu

End Sub

Sub GL_Stats_CA_Back_To_Menu()
    
    wshGL_Stats_CA.Visible = xlSheetHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
End Sub

