Attribute VB_Name = "modGL_PrepEF"
Option Explicit

Public gDictSoldeCodeEF As Object
Public gSoldeCodeEF() As Variant
Private gSavePremiereLigne As Integer
Private gLigneTotalPassif As Integer
Private gLigneTotalADA As Integer
Public gLigneTotalRevenus As Integer, gLigneTotalDépenses As Integer
Public gLigneAutresRevenus As Integer
Public gLigneRevenuNetAvantImpôts As Integer
Public gTotalRevenuNet_AC As Currency, gTotalRevenuNet_AP As Currency
Public gBNR_Début_Année_AC As Currency, gBNR_Début_Année_AP As Currency
Public gDividendes_Année_AC As Currency, gDividendes_Année_AP As Currency

Sub CalculerSoldesPourEF(ws As Worksheet, dateCutOff As Date) '2025-08-14 @ 06:50
    
    Application.EnableEvents = True
    
    Dim reponse As String
    Dim comparatif As String
    Dim choixEstValide As Boolean

    Do
        reponse = InputBox( _
            "Pour la colonne comparatif (année précédente), voulez-vous :" & vbNewLine & vbNewLine & _
            "1) La même période que l'année courante pour l'année dernière" & vbNewLine & vbNewLine & _
            "2) L'année dernière au complet (12 mois)" & vbNewLine, _
            "Choix du comparatif")

        Select Case Trim(reponse)
            Case "1"
                comparatif = "periode"
                choixEstValide = True
            Case "2"
                comparatif = "annee"
                choixEstValide = True
            Case Else
                MsgBox "Votre réponse est invalide..." & vbCrLf & vbCrLf & _
                        "Veuillez choisir entre 1 et 2 ?", vbExclamation
        End Select
    Loop Until choixEstValide
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:CalculerSoldesPourEF", ws.Name & ", " & dateCutOff, 0)
    
    Dim i As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Qui exécute ce programme ?
    Dim isDeveloppeur As Boolean
    If modFunctions.Fn_UtilisateurWindows() = "RobertMV" Or modFunctions.Fn_UtilisateurWindows() = "robertmv" Then
        isDeveloppeur = True
    End If
    
    'Déterminer la date de cutoff pour le comparatif (même période -OU- année complète)
    Dim cutOffAnPasse As Date
    If comparatif = "periode" Then
        cutOffAnPasse = Fn_DateMoinsUnAn(dateCutOff)
    Else
        cutOffAnPasse = Fn_DateMoinsUnAn(Fn_DernierJourAnneeFinanciere(dateCutOff))
    End If
    
    ws.Range("F5").Value = Format$(dateCutOff, wsdADMIN.Range("B1").Value)
    ws.Range("H5").Value = Format$(cutOffAnPasse, wsdADMIN.Range("B1").Value)
    'Effacer les cellules en place (contenu & format)
    
    ws.Unprotect
    ws.Range("C6:K199").ClearContents

    'Préparer l'appel à Fn_Tableau24MoisSommeTransGL
    Dim periodes As String
    periodes = Fn_Construire24PeriodesGL(dateCutOff)

    'Le tableau contiendra la somme des transactions par mois pour les 25 derniers mois
    Dim tableau() As Variant
    Dim inclureEcritureCloture As Boolean: inclureEcritureCloture = False
    tableau = Fn_Tableau24MoisSommeTransGL(dateCutOff, inclureEcritureCloture)
    ReDim gSoldeCodeEF(1 To UBound(tableau, 1), 1 To 3)
    
    'Déterminer le mois dans l'année financière
    Dim moisAnneeFinanciere As Long
    If month((dateCutOff)) > wsdADMIN.Range("MoisFinAnnéeFinancière").Value Then
        moisAnneeFinanciere = month(dateCutOff) - wsdADMIN.Range("MoisFinAnnéeFinancière").Value
    Else
        moisAnneeFinanciere = month(dateCutOff) + 12 - wsdADMIN.Range("MoisFinAnnéeFinancière").Value
    End If
    
    'Le plan comptable établit l'ordre de traitement
    Dim arr As Variant
    arr = Fn_PlanComptableTableau2D(4) 'Retourne un tableau avec 4 colonnes
    'Construire un Dictionary à partir du plan comptable
    Dim dictPlanComptable As Dictionary: Set dictPlanComptable = New Dictionary
    For i = LBound(arr, 1) To UBound(arr, 1)
        dictPlanComptable.Add arr(i, 1), arr(i, 2) & "|" & arr(i, 4) 'Description + | + Code E/F
    Next i
    
    'Mise en forme des cellules qui contiendront les montants
    ws.Range("C6:C" & UBound(arr, 1) + 7).HorizontalAlignment = xlCenter
    ws.Range("D6:D" & UBound(arr, 1) + 7).HorizontalAlignment = xlLeft
    ws.Range("E6:E" & UBound(arr, 1) + 7).HorizontalAlignment = xlCenter
    ws.Range("F6:H" & UBound(arr, 1) + 7).HorizontalAlignment = xlRight
    ws.Range("C6:H" & UBound(arr, 1) + 7).Font.Name = "Aptos Narrow"
    ws.Range("C6:H" & UBound(arr, 1) + 7).Font.size = 10
    
    'Instantiation du Dictionary
    If Not gDictSoldeCodeEF Is Nothing Then
        gDictSoldeCodeEF.RemoveAll
    End If
    If gDictSoldeCodeEF Is Nothing Then
        Set gDictSoldeCodeEF = CreateObject("Scripting.Dictionary")
    End If

    'Lire chacune des lignes du tableau à 26 colonnes pour calculer les 2 soldes
    Dim currRow As Long: currRow = 6
    Dim noCompteGL As String
    Dim descGL As String
    Dim codeEF As String
    Dim itemDict As String
    Dim soldeCourant As Currency
    Dim TotalAC As Currency
    Dim soldeComparatif As Currency
    Dim TotalAP As Currency
    Dim rowSommCodeEF As Long
    Dim ii As Long
    For i = LBound(tableau, 1) To UBound(tableau, 1)
        noCompteGL = CStr(tableau(i, 0))
        ws.Range("C" & currRow).Value = noCompteGL
        itemDict = dictPlanComptable(noCompteGL)
        descGL = Left(itemDict, InStr(itemDict, "|") - 1)
        codeEF = Mid$(itemDict, InStr(itemDict, "|") + 1)
        ws.Range("D" & currRow).Value = descGL
        ws.Range("E" & currRow).Value = codeEF
        If isDeveloppeur = True Then
            ws.Range("M" & currRow).Value = codeEF
            ws.Range("N" & currRow).Value = noCompteGL
        End If
        Dim ligne(1 To 25) As Variant
        For ii = 1 To 25
            ligne(ii) = tableau(i, ii)
        Next ii
        Call CalculerSoldesCourantEtComparatif(noCompteGL, moisAnneeFinanciere, ligne(), soldeCourant, soldeComparatif)
        ws.Range("F" & currRow).Value = soldeCourant
        TotalAC = TotalAC + soldeCourant
        ws.Range("H" & currRow).Value = soldeComparatif
        TotalAP = TotalAP + soldeComparatif
        If Not gDictSoldeCodeEF.Exists(codeEF) Then
            rowSommCodeEF = rowSommCodeEF + 1
            gDictSoldeCodeEF.Add codeEF, rowSommCodeEF
            gSoldeCodeEF(rowSommCodeEF, 1) = codeEF
        End If
        gSoldeCodeEF(rowSommCodeEF, 2) = gSoldeCodeEF(rowSommCodeEF, 2) + soldeCourant
        gSoldeCodeEF(rowSommCodeEF, 3) = gSoldeCodeEF(rowSommCodeEF, 3) + soldeComparatif
        
        'Sauvegarde des BNR au début de l'année et Dividendes
        If noCompteGL = "3100" Then
            gBNR_Début_Année_AC = soldeCourant
            gBNR_Début_Année_AP = soldeComparatif
        ElseIf noCompteGL = "3200" Then
            gDividendes_Année_AC = soldeCourant
            gDividendes_Année_AP = soldeComparatif
        End If

        If isDeveloppeur = True Then
            ws.Range("O" & currRow).Value = soldeCourant
            ws.Range("P" & currRow).Value = soldeComparatif
        End If
        
        currRow = currRow + 1
    
    Next i
    
    'Ajuste le format des montants
    ws.Range("F6:F" & currRow).NumberFormat = "###,###,##0.00 ;(###,###,##0.00);0.00"
    ws.Range("H6:H" & currRow).NumberFormat = "###,###,##0.00 ;(###,###,##0.00);0.00"

    ws.Protect UserInterfaceOnly:=True
    ws.EnableSelection = xlUnlockedCells

    Application.EnableEvents = True

    ActiveWindow.ScrollRow = 1

    Application.EnableEvents = False
    ws.Range("C6").Select
    Application.EnableEvents = True

    'Libérer la mémoire
    Set dictPlanComptable = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:CalculerSoldesPourEF", vbNullString, startTime)

End Sub

Sub shpPreparerEF_Click()

    Call AssemblerEtatsFinanciers
    
End Sub

Sub shpRetournerAuMenu_Click()

    Call RetournerAuMenu

End Sub

Sub RetournerAuMenu()

    Call modAppli.QuitterFeuillePourMenu(wshMenuGL, True) '2025-08-19 @ 06:59

End Sub
Sub AssemblerEtatsFinanciers() '2025-08-14 @ 08:05

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerEtatsFinanciers", vbNullString, 0)
    
    Dim dateAC As Date, dateAP As Date
    dateAC = wshGL_PrepEF.Range("F5").Value
    dateAP = wshGL_PrepEF.Range("H5").Value
    
    Call CreerFeuillesEtFormat
    
    Call AssemblerPageTitre0Main(dateAC, dateAP)
    Call AssemblerTM0Main(dateAC, dateAP)
    Call AssemblerER0Main(dateAC, dateAP)
    Call AssemblerBNR0Main(dateAC, dateAP)
    Call AssemblerBilan0Main(dateAC, dateAP)
    Call AssemblerNEFA_0Main(dateAC, dateAP)
    Call AssemblerNEFA2_0Main(dateAC, dateAP)
''    Call AssemblerNEFA3_0Main(dateAC, dateAP)
    
    Dim nomsFeuilles As Variant
    nomsFeuilles = Array("Page titre", "Table des Matières", "État des Résultats", "BNR", "Bilan", "A.tmp", "A2.tmp", "A3.tmp")
    
    Dim ws As Worksheet
    Dim i As Integer
    For i = UBound(nomsFeuilles) To LBound(nomsFeuilles) Step -1
        Set ws = ThisWorkbook.Sheets(nomsFeuilles(i)) 'Vérifier si la feuille existe déjà
        With ws
            'Sélectionner la feuille
            .Activate
            .Visible = xlSheetVisible
            'Afficher en mode aperçu des sauts de page
            ActiveWindow.View = xlPageBreakPreview
            'Affichage de la feuille à 80 %
            ActiveWindow.Zoom = 80
            'Police de base
            .Cells.Font.Name = "Verdana"
            'Remplir toutes les cellules avec la couleur blanche
            .Cells.Interior.Color = RGB(255, 255, 255) 'Blanc
        End With
    Next i
    
    'On se déplace à la première page des états financiers
    ActiveWorkbook.Sheets("Page Titre").Activate
    
    MsgBox "Les états financiers ont été produits" & vbNewLine & vbNewLine & _
            "Voir les onglets respectifs au bas du classeur", vbOKOnly, "Fin de traitement"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerEtatsFinanciers", vbNullString, startTime)

End Sub

Sub CreerFeuillesEtFormat() '2025-08-14 @ 09:32

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:CreerFeuillesEtFormat", vbNullString, 0)
    
    'Liste des feuilles à créer
    Dim nomsFeuilles As Variant
    nomsFeuilles = Array("Page titre", "Table des Matières", "État des Résultats", "BNR", "Bilan", "A.tmp", "A2.tmp", "A3.tmp")

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim i As Integer
    For i = LBound(nomsFeuilles) To UBound(nomsFeuilles)
        Application.StatusBar = "Création de " & nomsFeuilles(i)
        Set ws = Fn_ObtenirOuCreerFeuille(nomsFeuilles(i))

        With ws.PageSetup
            .Orientation = xlPortrait
            .FitToPagesWide = False
            .FitToPagesTall = False
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .CenterHorizontally = False
        End With
        Set ws = Nothing
    Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:CreerFeuillesEtFormat", vbNullString, startTime)
    
End Sub

Sub AssemblerPageTitre0Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerPageTitre0Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Page Titre")
    
    Application.StatusBar = "Construction de la page titre"
        
    Call AssemblerPageTitre1EtArrierePlanEtEntete(ws, dateAC, dateAP)
    
    Application.StatusBar = False
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerPageTitre0Main", vbNullString, startTime)

End Sub

Sub AssemblerPageTitre1EtArrierePlanEtEntete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerPageTitre1EtArrierePlanEtEntete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter
    
    Call PositionnerCellule(ws, UCase$(wsdADMIN.Range("NomEntreprise")), 8, 2, 20, True, xlCenter)
    Call PositionnerCellule(ws, UCase$("États Financiers"), 15, 2, 20, True, xlCenter)
    Call PositionnerCellule(ws, UCase$(Format$(dateAC, "dd mmmm yyyy")), 28, 2, 20, True, xlCenter)
    
    'Ajuster la largeur des colonnes et la hauteur de lignes
    Call ConfigurerColonnesEtLignes(ws, Array(3, 87, 3), "1:28")
    
    'Ajuster la police pour la feuille
    Call AppliquerMiseEnPageEF(ws, 20)

    'Fixer le printArea selon le nombre de lignes ET 3 colonnes
    ws.PageSetup.PrintArea = "$A$1:$C$" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 4
    Debug.Print "Page Titre (Entête) - $A$1:$C$' & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 4"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerPageTitre1EtArrierePlanEtEntete", vbNullString, startTime)

End Sub

Sub AssemblerTM0Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerTM0Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Table des Matières")
    
    Application.StatusBar = "Construction de la table des matières"
    
    Call AssemblerTM1ArrierePlanEtEntete(ws, dateAC, dateAP)
    Call AssemblerTM2Lignes(ws)
    
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerTM0Main", vbNullString, startTime)

End Sub

Sub AssemblerTM1ArrierePlanEtEntete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerTM1ArrierePlanEtEntete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Appliquer le format d'en-tête
    Call AjouterEnteteEF(ws, wsdADMIN.Range("NomEntreprise"), dateAC, 1)
    
    With ws.Range("B5:C5").Borders(xlEdgeBottom)
'    With ws.Range("B6:E6").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    'Ajuster la largeur des colonnes et la hauteur des lignes
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 75
    ws.Columns("C").ColumnWidth = 11
    ws.Columns("D").ColumnWidth = 3
    ws.Rows("1:25").RowHeight = 15
    
    'Fixer le printArea selon le nombre de lignes ET 3 colonnes
    ws.PageSetup.PrintArea = "$A$1:$D$" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3
    Debug.Print "Table des matières (entête) - $A$1:$D$' & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerTM1ArrierePlanEtEntete", vbNullString, startTime)

End Sub

Sub AssemblerTM2Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerTM2Lignes", vbNullString, 0)
    
    'Première ligne
    Dim currRow As Integer
    currRow = 15
    
    With ws
        .Range("C" & currRow).Value = "Page"
        currRow = currRow + 3
        
        .Range("B" & currRow).Value = "États des résultats"
        .Range("C" & currRow).Value = "2"
        currRow = currRow + 2
        
        .Range("B" & currRow).Value = "États des Bénéfices non répartis"
        .Range("C" & currRow).Value = "3"
        currRow = currRow + 2
        
        .Range("B" & currRow).Value = "Bilan"
        .Range("C" & currRow).Value = "4"
        currRow = currRow + 2
        
        .Range("C:C").HorizontalAlignment = xlRight
        
       'Ajuster la police pour la feuille
        Call AppliquerMiseEnPageEF(ws, 11)
    
    End With
    
    'Fixer le printArea selon le nombre de lignes ET 3 colonnes
    ws.PageSetup.PrintArea = "$A$1:$D$" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    Debug.Print "Table des matières (lignes) - $A$1:$D$' & ws.Cells(ws.Rows.count, 2).End(xlUp).Row"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerTM2Lignes", vbNullString, startTime)

End Sub

Sub AssemblerER0Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerER0Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("État des résultats")
    
    Application.StatusBar = "Construction de l'état des résultats"
    
    Call AssemblerER1ArrierePlanEtEntete(ws, dateAC, dateAP)
    Call AssemblerER2Lignes(ws)
    
    'On ajoute le Revenu Net au BNR du bilan via variables Globales
    Dim indice As Integer
    indice = gDictSoldeCodeEF("E02")
    gSoldeCodeEF(indice, 2) = gSoldeCodeEF(indice, 2) - gTotalRevenuNet_AC
    gSoldeCodeEF(indice, 3) = gSoldeCodeEF(indice, 3) - gTotalRevenuNet_AP
    
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerER0Main", vbNullString, startTime)

End Sub

Sub AssemblerER1ArrierePlanEtEntete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerER1ArrierePlanEtEntete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Titre de l'état des résultats
    Dim jourAC As Integer, moisAC As Integer, anneeAC As Integer
    jourAC = day(dateAC)
    moisAC = month(dateAC)
    anneeAC = year(dateAC)
    
    Dim titre As String
    titre = Fn_TitreSelonNombreDeMois(dateAC)
    
    'Appliquer le format d'en-tête
    Call PositionnerCellule(ws, UCase$(wsdADMIN.Range("NomEntreprise")), 1, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("État des Résultats"), 2, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$(titre), 3, 2, 12, True, xlLeft)
    ws.Range("C5:E6").HorizontalAlignment = xlRight
    ws.Range("C5").Value = year(dateAC)
    ws.Range("C5").Font.Bold = True
    ws.Range("E5").Value = year(dateAP)
    ws.Range("E5").Font.Bold = True
    With ws.Range("B5:E5").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -11511710
        .Weight = xlMedium
    End With
    
    ws.Range("C7:C45").NumberFormat = "###,##0 $;(###,##0) $; 0 $"
    ws.Range("E7:E45").NumberFormat = "###,##0 $;(###,##0) $; 0 $"

    'Ajuster la largeur des colonnes et la hauteur de lignes
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 52
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 3
    ws.Rows("1:45").RowHeight = 15

    ws.PageSetup.CenterFooter = 2
     
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerER1ArrierePlanEtEntete", vbNullString, startTime)

End Sub

Sub AssemblerER2Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerER2Lignes", vbNullString, 0)
    
    Dim wsAdmin As Worksheet: Set wsAdmin = wsdADMIN
    
    Dim tbl As ListObject
    Set tbl = wsAdmin.ListObjects("tblÉtatsFinanciersCodes")
    
    Dim LigneEF As String
    Dim codeEF As String
    Dim typeLigne As String
    Dim gras As String
    Dim souligne As String
    Dim ligneTotalDepenses As Long
    Dim size As Long
    'Première ligne
    Dim currRow As Integer
    currRow = 8
    Dim rngRow As ListRow
    For Each rngRow In tbl.ListRows
        LigneEF = rngRow.Range.Cells(1, 1).Value
        codeEF = UCase$(rngRow.Range.Cells(1, 2).Value)
        'On ne traite que les lignes de l'État des résultats (R, D, X & I)
        If InStr("RDXI", Left$(codeEF, 1)) <> 0 Then
            typeLigne = UCase$(rngRow.Range.Cells(1, 3).Value)
            gras = UCase$(rngRow.Range.Cells(1, 4).Value)
            souligne = UCase$(rngRow.Range.Cells(1, 5).Value)
            size = rngRow.Range.Cells(1, 6).Value
            If codeEF = "D99" Then
                ligneTotalDepenses = currRow
            End If
            Call ImprimerLigneEF(ws, currRow, LigneEF, codeEF, typeLigne, gras, souligne, size)
        End If
        
    Next rngRow
    
    'Ajuster la police pour la feuille
    Call AppliquerMiseEnPageEF(ws, 10)

    'Augmenter la taille de police pour les 3 premières lignes
    ws.Range("1:3").Font.size = 12
    
    'Transfère les montants NON arrondis dans les cellules sans les cents
    Dim i As Integer
    For i = 7 To currRow
        If ws.Range("G" & i).Value <> vbNullString Then
            ws.Range("C" & i).Value = ws.Range("G" & i).Value
            ws.Range("E" & i).Value = ws.Range("I" & i).Value
        End If
    Next i
    ws.Range("G:J").Delete
    
    'Tri par ordre descendant une plage
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("C17:C" & ligneTotalDepenses - 2), Order:=xlDescending
        .SetRange ws.Range("B17:E" & ligneTotalDepenses)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Fixer le printArea selon le nombre de lignes ET colonnes
    ActiveSheet.PageSetup.PrintArea = "$A$1:$F$" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3
    Debug.Print "État des Résultats (Lignes) - $A$1:$F$' & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerER2Lignes", vbNullString, startTime)

End Sub

Sub AssemblerBilan0Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBilan0Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Bilan")
    
    Application.StatusBar = "Construction du bilan"
    
    Call AssemblerBilan1ArrierePlanEtEntete(ws, dateAC, dateAP)
    Call AssemblerBilan2Lignes(ws)
    
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBilan0Main", vbNullString, startTime)
    
End Sub

Sub AssemblerBilan1ArrierePlanEtEntete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBilan1ArrierePlanEtEntete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Appliquer le format d'en-tête
    Call PositionnerCellule(ws, UCase$(wsdADMIN.Range("NomEntreprise")), 1, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("Bilan"), 2, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("Au " & Format$(dateAC, "dd mmmm yyyy")), 3, 2, 12, True, xlLeft)
    ws.Range("C5:E6").HorizontalAlignment = xlRight
    ws.Range("C5").Value = year(dateAC)
    ws.Range("C5").Font.Bold = True
    ws.Range("E5").Value = year(dateAP)
    ws.Range("E5").Font.Bold = True
    With ws.Range("B5:E5").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -11511710
        .Weight = xlMedium
    End With
    
    Dim currRow As Integer
    currRow = 8

    ws.Range("C" & currRow & ":C40").NumberFormat = "#,##0 $;(#,##0) $; 0 $"
    ws.Range("E" & currRow & ":E40").NumberFormat = "#,##0 $;(#,##0) $; 0 $"

    'Ajuster la largeur des colonnes et la hauteur des lignes
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 52
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 3
    ws.Rows("1:40").RowHeight = 15
    
    ws.PageSetup.CenterFooter = 4
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBilan1ArrierePlanEtEntete", vbNullString, startTime)

End Sub

Sub AssemblerBilan2Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBilan2Lignes", vbNullString, 0)
    
    Dim wsAdmin As Worksheet
    Set wsAdmin = wsdADMIN
    
    Dim tbl As ListObject
    Set tbl = wsAdmin.ListObjects("tblÉtatsFinanciersCodes")
    
    Dim LigneEF As String, codeEF As String, typeLigne As String, gras As String, souligne As String
    Dim size As Long
    Dim currRow As Integer
    currRow = 8
    Dim rngRow As ListRow
    For Each rngRow In tbl.ListRows
        LigneEF = rngRow.Range.Cells(1, 1).Value
        codeEF = rngRow.Range.Cells(1, 2).Value
        'Ne traite que les lignes du bilan (A, P & E)
        If InStr("APE", Left$(codeEF, 1)) <> 0 Then
            typeLigne = rngRow.Range.Cells(1, 3).Value
            gras = rngRow.Range.Cells(1, 4).Value
            souligne = rngRow.Range.Cells(1, 5).Value
            size = rngRow.Range.Cells(1, 6).Value
            Call ImprimerLigneEF(ws, currRow, LigneEF, codeEF, typeLigne, gras, souligne, size)
        End If
        
    Next rngRow
    
    'Ajuster la police pour la feuille
    Call AppliquerMiseEnPageEF(ws, 10)

    'Augmenter la taille de police pour les 3 premières lignes
    ws.Range("1:3").Font.size = 12
    
    'Transfère les montants NON arrondis dans les cellules sans les cents
    Dim i As Integer
    For i = 7 To currRow
        If ws.Range("G" & i).Value <> vbNullString Then
            ws.Range("C" & i).Value = ws.Range("G" & i).Value
            ws.Range("E" & i).Value = ws.Range("I" & i).Value
        End If
    Next i
    ws.Range("G7:I38").Clear
    
    'Fixer le printArea selon le nombre de lignes ET colonnes
    ActiveSheet.PageSetup.PrintArea = "$A$1:$F$" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3
    Debug.Print "Bilan (lignes) - $A$1:$F$' & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBilan2Lignes", vbNullString, startTime)

End Sub

Sub AssemblerBNR0Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBNR0Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BNR")
    
    Application.StatusBar = "Construction de l'état des bénéfices non répartis"
    
    Call AssemblerBNR1ArrierePlanEtEntete(ws, dateAC, dateAP)
    Call AssemblerBNR2Lignes(ws)
    
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBNR0Main", vbNullString, startTime)
    
End Sub

Sub AssemblerBNR1ArrierePlanEtEntete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBNR1ArrierePlanEtEntete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Titre de l'état des résultats
    Dim jourAC As Integer, moisAC As Integer, anneeAC As Integer
    jourAC = day(dateAC)
    moisAC = month(dateAC)
    anneeAC = year(dateAC)
    
    Dim titre As String
    titre = Fn_TitreSelonNombreDeMois(dateAC)
    
    'Appliquer le format d'en-tête
    Call PositionnerCellule(ws, UCase$(wsdADMIN.Range("NomEntreprise")), 1, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("Bénéfices non répartis"), 2, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$(titre), 3, 2, 12, True, xlLeft)
    ws.Range("C5:E6").HorizontalAlignment = xlRight
    ws.Range("C5").Value = year(dateAC)
    ws.Range("C5").Font.Bold = True
    ws.Range("E5").Value = year(dateAP)
    ws.Range("E5").Font.Bold = True
    With ws.Range("B5:E5").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -11511710
        .Weight = xlMedium
    End With
    
    ws.Range("C7:C20").NumberFormat = "#,##0 $;(#,##0) $; 0 $"
    ws.Range("E7:E20").NumberFormat = "#,##0 $;(#,##0) $; 0 $"

    'Ajuster la largeur des colonnes et la hauteur des lignes
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 52
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 3
    ws.Rows("1:20").RowHeight = 15
    
    ws.PageSetup.CenterFooter = 3
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBNR1ArrierePlanEtEntete", vbNullString, startTime)

End Sub

Sub AssemblerBNR2Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBNR2Lignes", vbNullString, 0)
    
    Dim wsAdmin As Worksheet
    Set wsAdmin = wsdADMIN
    
    Dim tbl As ListObject
    Set tbl = wsAdmin.ListObjects("tblÉtatsFinanciersCodes")
    
    Dim LigneEF As String, codeEF As String, typeLigne As String, gras As String, souligne As String
    Dim size As Long
    Dim currRow As Integer
    currRow = 8
    Dim rngRow As ListRow
    For Each rngRow In tbl.ListRows
        LigneEF = rngRow.Range.Cells(1, 1).Value
        codeEF = rngRow.Range.Cells(1, 2).Value
        'Ne traite que les lignes du bilan (A, P & E)
        If InStr("B", Left$(codeEF, 1)) <> 0 Then
            typeLigne = rngRow.Range.Cells(1, 3).Value
            gras = rngRow.Range.Cells(1, 4).Value
            souligne = rngRow.Range.Cells(1, 5).Value
            size = rngRow.Range.Cells(1, 6).Value
            Call ImprimerLigneEF(ws, currRow, LigneEF, codeEF, typeLigne, gras, souligne, size)
        End If
        
    Next rngRow
    
    'Ajuster la police pour la feuille
    Call AppliquerMiseEnPageEF(ws, 10)
    
    'Augmenter la taille de police pour les 3 premières lignes
    ws.Range("1:3").Font.size = 12

    'Transfère les montants NON arrondis dans les cellules sans les cents
    Dim i As Integer
    For i = 7 To currRow
        If ws.Range("G" & i).Value <> vbNullString Then
            ws.Range("C" & i).Value = ws.Range("G" & i).Value
            ws.Range("E" & i).Value = ws.Range("I" & i).Value
        End If
    Next i
    ws.Range("G:J").Delete '2025-08-01 @ 21:53
    
    'Fixer le printArea selon le nombre de lignes ET colonnes
    ActiveSheet.PageSetup.PrintArea = "$A$1:$F$" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3
    Debug.Print "BNR (lignes) - $A$1:$F$' & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerBNR2Lignes", vbNullString, startTime)

End Sub

Sub AssemblerNEFA_0Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA_0Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("A.tmp")
    
    Application.StatusBar = "Construction des notes"
    
    Call AssemblerNEFA_1ArrierePlanEtEntete(ws, dateAC, dateAP)
    Call AssemblerNEFA_2Lignes(ws)
    
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA_0Main", vbNullString, startTime)
    
End Sub

Sub AssemblerNEFA_1ArrierePlanEtEntete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA_1ArrierePlanEtEntete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Appliquer le format d'en-tête
    ws.Range("A1:E1").Merge
    Call PositionnerCellule(ws, UCase$(wsdADMIN.Range("NomEntreprise")), 1, 1, 12, True, xlLeft)
    ws.Range("A1").IndentLevel = 3
    
    Call PositionnerCellule(ws, UCase$("NOTES AUX ÉTATS FINANCIERS"), 2, 1, 12, True, xlLeft)
    ws.Range("A2").IndentLevel = 3
    
    Call PositionnerCellule(ws, UCase$("Au " & Format$(dateAC, "dd mmmm yyyy")), 3, 1, 12, True, xlLeft)
    ws.Range("A3").IndentLevel = 3
    
    With ws.Range("F4")
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Value = "5"
    End With
    With ws.Range("A4:F4").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -11511710
        .Weight = xlMedium
    End With
    
    'Entête
    With ws.Range("A1:F3")
        .Font.Name = "Verdana"
        .Font.size = 12
        .Font.Color = RGB(140, 131, 117)
    End With

    'Corps
    With ws.Range("A4:F26")
        .Font.Name = "Verdana"
        .Font.size = 11
        .Font.Color = RGB(140, 131, 117)
    End With
    
    'Hauteur de lignes
    ws.Rows("1:3").RowHeight = 15
    ws.Rows("4").RowHeight = 16.5
    ws.Rows("5:7").RowHeight = 14.25
    ws.Rows("8").RowHeight = 40
    ws.Rows("9:14").RowHeight = 14.25
    ws.Rows("15:16").RowHeight = 15
    ws.Rows("17:19").RowHeight = 14.25
    ws.Rows("20").RowHeight = 28.5
    ws.Rows("21:27").RowHeight = 14.25
    
    Dim currRow As Integer
    currRow = 7

    'Note # 1 - Lignes 7 @ 8
    With ws.Range("A7")
        .HorizontalAlignment = xlCenter
        .Value = "1"
        .Font.Bold = True
    End With
    
    With ws.Range("B7")
        .HorizontalAlignment = xlLeft
        .Value = "CONSTITUTION DE LA SOCIÉTÉ"
        .Font.Bold = True
    End With
    
    With ws.Range("B8:E8")
        .MergeCells = True
        .ShrinkToFit = False
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .Font.Bold = False
        .Value = "La société a été constituée le 24 juillet 2008 en vertue de la Partie IA de la Loi " & _
                 "sur les Compagnies du Québec. Elle œuvre dans le domaine de la consultation en fiscalité."
    End With
        
    'Note # 2 - Lignes 11 @ 15
    With ws.Range("A11")
        .HorizontalAlignment = xlCenter
        .Value = "2"
        .Font.Bold = True
    End With
    
    With ws.Range("B11")
        .HorizontalAlignment = xlLeft
        .Value = "FRAIS PAYÉS D'AVANCES"
        .Font.Bold = True
    End With
    
    With ws.Range("B12:E12")
        .MergeCells = True
        .ShrinkToFit = False
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .Value = "Les frais payés d'avance sont constitués des éléments suivants:"
        .Font.Bold = False
    End With
    
    With ws.Range("B13")
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .Value = "Cotisation et Assurances professionnelle"
        .IndentLevel = 2
        .Font.Bold = False
    End With
    
    With ws.Range("D13")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .formula = "=ROUND((2927.5+4846.14)*8/12,0)" '@TODO - 2025-10-24 @ 05:41
    End With
    
    With ws.Range("B14")
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .IndentLevel = 2
        .Value = "Loyer payé d'avance - 1 mois"
        .Font.Bold = False
    End With
    
    With ws.Range("D14")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .Value = 645 '@TODO - 2025-10-24 @ 05:41
    End With
    
    With ws.Range("B15")
        .HorizontalAlignment = xlLeft
        .Value = "TOTAL"
        .Font.Bold = True
    End With
    
    With ws.Range("D15")
        .HorizontalAlignment = xlRight
        .Font.Bold = True
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .formula = "=SUM(D13:D14)" '@TODO - 2025-10-24 @ 05:41
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
    End With
    
    'Note # 3 - Lignes 18 @ 26
    With ws.Range("A18")
        .HorizontalAlignment = xlCenter
        .Value = "3"
        .Font.Bold = True
    End With
    
    With ws.Range("B18")
        .HorizontalAlignment = xlLeft
        .Value = "IMMOBILISATIONS"
        .Font.Bold = True
    End With
    
    With ws.Range("B19:E19")
        .MergeCells = True
        .ShrinkToFit = False
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .Value = "Les immobilisations sont constitués des éléments suivants:"
        .Font.Bold = False
    End With
    
    With ws.Range("C20")
        .HorizontalAlignment = xlCenter
        .Value = "Coût"
        .Font.Bold = True
    End With
    
    With ws.Range("D20")
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .Value = "Amortissement cumulé"
        .Font.Bold = True
    End With
    
    With ws.Range("E20")
        .HorizontalAlignment = xlCenter
        .Value = "Valeur nette"
        .Font.Bold = True
    End With
    
    '@TODO - Aller chercher les soldes
    With ws.Range("E21")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .formula = "=SUM(C21:D21)"
    End With
    
    With ws.Range("E22")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .formula = "=SUM(C22:D22)"
    End With
    
    With ws.Range("E23")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .formula = "=SUM(C23:D23)"
    End With
    
    With ws.Range("E24")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .formula = "=SUM(C24:D24)"
    End With
    
    With ws.Range("C24:E24")
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    With ws.Range("B26")
        .HorizontalAlignment = xlLeft
        .Value = "TOTAL"
        .Font.Bold = True
    End With
    
    With ws.Range("C26")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .Font.Bold = True
        .formula = "=SUM(C21:C24)"
    End With
    
    With ws.Range("D26")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .Font.Bold = True
        .formula = "=SUM(D21:D24)"
    End With
    
    With ws.Range("E26")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .Font.Bold = True
        .formula = "=SUM(E21:E24)"
    End With
    
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 39
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 3.75
    
    ws.PageSetup.CenterFooter = 5
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA_1ArrierePlanEtEntete", vbNullString, startTime)

End Sub

Sub AssemblerNEFA_2Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA_2Lignes", vbNullString, 0)
    
    'Fixer le printArea selon le nombre de lignes ET colonnes
    ActiveSheet.PageSetup.PrintArea = "$A$1:$F$27"
    Debug.Print "Notes A (lignes) - $A$1:$F$27"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA_2Lignes", vbNullString, startTime)

End Sub

Sub AssemblerNEFA2_0Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA2_0Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("A2.tmp")
    
    Application.StatusBar = "Construction des notes"
    
    Call AssemblerNEFA2_1ArrierePlanEtEntete(ws, dateAC, dateAP)
    Call AssemblerNEFA2_2Lignes(ws)
    
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA2_0Main", vbNullString, startTime)
    
End Sub

Sub AssemblerNEFA2_1ArrierePlanEtEntete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA2_1ArrierePlanEtEntete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Appliquer le format d'en-tête
    ws.Range("A1:E1").Merge
    Call PositionnerCellule(ws, UCase$(wsdADMIN.Range("NomEntreprise")), 1, 1, 12, True, xlLeft)
    ws.Range("A1").IndentLevel = 3
    
    ws.Range("A2:E2").Merge
    Call PositionnerCellule(ws, UCase$("NOTES AUX ÉTATS FINANCIERS"), 2, 1, 12, True, xlLeft)
    ws.Range("A2").IndentLevel = 3
    
    Call PositionnerCellule(ws, UCase$("Au " & Format$(dateAC, "dd mmmm yyyy")), 3, 1, 12, True, xlLeft)
    ws.Range("A3").IndentLevel = 3
    
    With ws.Range("H4")
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Value = "6"
    End With
    With ws.Range("A4:H4").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -11511710
        .Weight = xlMedium
    End With
    
    'Entête
    With ws.Range("A1:H3")
        .Font.Name = "Verdana"
        .Font.size = 12
        .Font.Color = RGB(140, 131, 117)
    End With

    'Corps
    With ws.Range("A4:G19")
        .Font.Name = "Verdana"
        .Font.size = 11
        .Font.Color = RGB(140, 131, 117)
    End With
    
    'Hauteur de lignes
    ws.Rows("1:19").RowHeight = 14.25
    ws.Rows("9").RowHeight = 30
    ws.Rows("14").RowHeight = 65
    
    'Note # 4 - Lignes 8 @ 17
    With ws.Range("A8")
        .HorizontalAlignment = xlCenter
        .Value = "4"
        .Font.Bold = True
    End With
    
    With ws.Range("B8")
        .HorizontalAlignment = xlLeft
        .Value = "AMORTISSEMENT"
        .Font.Bold = True
    End With
    
    With ws.Range("B9:G9")
        .MergeCells = True
        .ShrinkToFit = False
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .Font.Bold = False
        .Value = "L'amortissement des immobilisation et des frais de constitution est effectuée de la façon suivante:"
    End With
        
    'Note # 2 - Lignes 11 @ 15
    With ws.Range("E13")
        .HorizontalAlignment = xlCenter
        .Value = "Dégressif"
    End With
    
    With ws.Range("F13")
        .HorizontalAlignment = xlCenter
        .NumberFormat = "##0.00 %"
        .Value = ".55"
    End With
    
    With ws.Range("G13")
        .HorizontalAlignment = xlCenter
        .Value = "Variable"
    End With
    
    With ws.Range("H13")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .Value = 0
    End With
    
    With ws.Range("E14")
        .HorizontalAlignment = xlCenter
        .Value = "Dégressif"
    End With
    
    With ws.Range("F14")
        .HorizontalAlignment = xlCenter
        .NumberFormat = "##0.00 %"
        .Value = ".20"
    End With
    
    With ws.Range("G14")
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .Value = "Demi taux la première année"
    End With
    
    With ws.Range("H14")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .Value = 0
    End With
    
     With ws.Range("E15")
        .HorizontalAlignment = xlCenter
        .Value = "Dégressif"
    End With
    
    With ws.Range("F15")
        .HorizontalAlignment = xlCenter
        .NumberFormat = "##0.00 %"
        .Value = 1
    End With
    
    With ws.Range("G15")
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .Value = "Variable"
    End With
    
    With ws.Range("H15")
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .Value = 0
    End With
    
    With ws.Range("B17")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Value = "TOTAL DES AMORTISSEMENTS"
    End With
    
    With ws.Range("H17")
        .HorizontalAlignment = xlRight
        .Font.Bold = True
        .NumberFormat = "#,##0 $;(#,##0) $; 0 $"
        .formula = "=SUM(H13:H15)"
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
    End With
    
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C").ColumnWidth = 1
    ws.Columns("D").ColumnWidth = 1
    ws.Columns("E").ColumnWidth = 11
    ws.Columns("F").ColumnWidth = 11
    ws.Columns("G").ColumnWidth = 11
    ws.Columns("H").ColumnWidth = 14
    
    ws.PageSetup.CenterFooter = 6
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA2_1ArrierePlanEtEntete", vbNullString, startTime)

End Sub

Sub AssemblerNEFA2_2Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA2_2Lignes", vbNullString, 0)
    
    'Fixer le printArea selon le nombre de lignes ET colonnes
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$19"
    Debug.Print "Notes A2 (lignes) - $A$1:$H$19"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerNEFA2_2Lignes", vbNullString, startTime)

End Sub

Sub PositionnerCellule(ws As Worksheet, cell As String, ligne As Integer, col As Integer, points As Integer, gras As Boolean, alignement As Long)

    With ws.Cells(ligne, col)
        .Value = cell
        .Font.size = points
        .Font.Bold = gras
        .HorizontalAlignment = alignement
    End With
    
End Sub

Sub ImprimerLigneEF(ws As Worksheet, ByRef currRow As Integer, LigneEF As String, codeEF As String, typeLigne As String, gras As String, souligne As String, size As Long)
    
    Dim correcteurSigne As Integer
    Dim section As String
    section = Left$(codeEF, 1)
    correcteurSigne = IIf(InStr("PERIB", section), -1, 1)
    
    Dim doitImprimer As Boolean
    doitImprimer = True
    Dim index As Integer
    Select Case typeLigne
    
        Case "E" 'Entête
            If InStr("E00^D00^", codeEF & "^") = 0 Then 'Saute une ligne AVANT d'imprimer
                currRow = currRow + 1
            End If
            If codeEF = "B00" Then
                ws.Range("G" & currRow).Value = gBNR_Début_Année_AC * correcteurSigne
                ws.Range("I" & currRow).Value = gBNR_Début_Année_AP * correcteurSigne
                gSavePremiereLigne = currRow
            Else
                gSavePremiereLigne = currRow + 1
            End If
            
            If gSavePremiereLigne = 0 Then Stop
            
        Case "G" 'Groupement
            index = gDictSoldeCodeEF(codeEF)
            If index <> 0 Then
                If Round(gSoldeCodeEF(index, 2), 2) <> 0 Or Round(gSoldeCodeEF(index, 3), 2) <> 0 Then
                    ws.Range("G" & currRow).Value = gSoldeCodeEF(index, 2) * correcteurSigne
                    ws.Range("I" & currRow).Value = gSoldeCodeEF(index, 3) * correcteurSigne
                Else
                    doitImprimer = False
                End If
            Else
                doitImprimer = False
            End If
        
        Case "T" 'Totaux
            If InStr("E50^E60^", codeEF & "^") = 0 Then 'Saute une ligne AVANT d'imprimer
                currRow = currRow + 1
            End If
            If codeEF <> "E60" And codeEF <> "B10" Then 'Bordure en haut de la cellule
                With ws.Range("C" & currRow).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Color = -11511710
                    .Weight = xlThin
                End With
                With ws.Range("E" & currRow).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Color = -11511710
                    .Weight = xlThin
                End With
            End If
            
            If codeEF = "E60" Then
                ws.Range("G" & currRow).formula = "=sum(G" & gLigneTotalPassif & ", G" & gLigneTotalADA & ")"
                ws.Range("I" & currRow).formula = "=sum(I" & gLigneTotalPassif & ", I" & gLigneTotalADA & ")"
            ElseIf codeEF = "I01" Then
                ws.Range("G" & currRow).formula = "=sum(G" & gLigneTotalRevenus & " - G" & gLigneTotalDépenses & " + G" & gLigneAutresRevenus & ")"
                ws.Range("I" & currRow).formula = "=sum(I" & gLigneTotalRevenus & " - I" & gLigneTotalDépenses & " + I" & gLigneAutresRevenus & ")"
            ElseIf codeEF = "I03" Then
                ws.Range("G" & currRow).formula = "=sum(G" & gLigneRevenuNetAvantImpôts & ":G" & currRow - 1 & ")"
                ws.Range("I" & currRow).formula = "=sum(I" & gLigneRevenuNetAvantImpôts & ":I" & currRow - 1 & ")"
            Else
                ws.Range("G" & currRow).formula = "=sum(G" & gSavePremiereLigne & ":G" & currRow - 1 & ")"
                ws.Range("I" & currRow).formula = "=sum(I" & gSavePremiereLigne & ":I" & currRow - 1 & ")"
            End If
            'Bordures dans le bas de la cellule
            If codeEF = "I01" Or codeEF = "I03" Then
                With ws.Range("C" & currRow).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Color = -11511710
                    .Weight = xlThin
                End With
                With ws.Range("E" & currRow).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Color = -11511710
                    .Weight = xlThin
                End With
            End If
            
            'Partir un nouveau sous-total, sans entête
            If codeEF = "B10" Then gSavePremiereLigne = currRow
            
    End Select
        
    'Certaines lignes ont besoin d'être notées pour utilisation particulière
    If codeEF = "P99" Then gLigneTotalPassif = currRow
    If codeEF = "E50" Then gLigneTotalADA = currRow
    If codeEF = "R99" Then gLigneTotalRevenus = currRow
    If codeEF = "D99" Then gLigneTotalDépenses = currRow
    If codeEF = "R04" Then gLigneAutresRevenus = currRow
    If codeEF = "I01" Then gLigneRevenuNetAvantImpôts = currRow
    
    'Sauvegarder les 2 montants de Revenu Net
    If codeEF = "I03" Then
        gTotalRevenuNet_AC = ws.Range("G" & currRow).Value2
        gTotalRevenuNet_AP = ws.Range("I" & currRow).Value2
    End If
    
    With ws.Range("B" & currRow & ":E" & currRow).Font
        If UCase$(gras) = "VRAI" Then
            .Bold = True
        End If
        If UCase$(souligne) = "VRAI" Then
            .underline = xlUnderlineStyleSingle
        End If
        If size <> 0 Then
            .size = size
        End If
    End With
    
    If codeEF = "I02" Then
        ws.Range("C" & currRow & ":E" & currRow).Font.Bold = False
    End If
    
    If codeEF = "B01" Then 'Bénéfice net / Revenu net
        ws.Range("B" & currRow).Value = LigneEF
        ws.Range("G" & currRow).Value = gTotalRevenuNet_AC
        ws.Range("I" & currRow).Value = gTotalRevenuNet_AP
        With ws.Range("C" & currRow).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -11511710
            .Weight = xlThin
        End With
        With ws.Range("E" & currRow).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -11511710
            .Weight = xlThin
        End With
        currRow = currRow + 1
    End If
    
    If codeEF = "B20" Then 'Dividendes
        ws.Range("B" & currRow).Value = LigneEF
        ws.Range("G" & currRow).Value = -gDividendes_Année_AC
        ws.Range("I" & currRow).Value = -gDividendes_Année_AP
        currRow = currRow + 1
    End If
    
    If codeEF = "B50" Then 'Solde de fin (BNR)
        With ws.Range("C" & currRow).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With ws.Range("C" & currRow).Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With ws.Range("E" & currRow).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With ws.Range("E" & currRow).Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
    End If
    
    If doitImprimer = True Then
        ws.Range("B" & currRow).Value = LigneEF
        currRow = currRow + 1
    End If
    
    If typeLigne = "T" Then
        currRow = currRow + 1
    End If
    
    If codeEF = "R00" Or codeEF = "D00" Or codeEF = "B00" Then
        currRow = currRow + 1
    End If
    
End Sub

Sub AppliquerMiseEnPageEF(ws As Worksheet, taillePolice As Integer) '2025-08-14 @ 09:12

    With ws.Cells
        .Font.Name = "Verdana"
        .Font.size = taillePolice
        .Font.Color = RGB(140, 131, 117)
    End With

End Sub

Sub ConfigurerColonnesEtLignes(ws As Worksheet, largeurCols As Variant, hauteurLignes As String) '2025-08-14 @ 09:37

    Dim i As Integer
    For i = LBound(largeurCols) To UBound(largeurCols)
        ws.Columns(Chr(65 + i)).ColumnWidth = largeurCols(i)
    Next i
    ws.Rows(hauteurLignes).RowHeight = 20
    
End Sub

Sub AjouterEnteteEF(ws As Worksheet, nomEntreprise As String, dateEF As Date, ligneDépart As Integer) '2025-08-14 @ 09:40

    Call PositionnerCellule(ws, UCase$(nomEntreprise), ligneDépart, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("Table des Matières"), ligneDépart + 1, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("États Financiers"), ligneDépart + 2, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("Au " & Format$(dateEF, "dd mmmm yyyy")), ligneDépart + 3, 2, 12, True, xlLeft)
    
End Sub

Function Fn_TitreSelonNombreDeMois(dateAC As Date) As String '2025-08-14 @ 19:42

    Dim dateFinAnneeFinanciere As Date
    dateFinAnneeFinanciere = Fn_DernierJourAnneeFinanciere(dateAC)
    
    Dim nbMois As Integer
    
    If month(dateAC) > wsdADMIN.Range("MoisFinAnnéeFinancière") Then
        nbMois = month(dateAC) - wsdADMIN.Range("MoisFinAnnéeFinancière")
    Else
        nbMois = month(dateAC) + 12 - wsdADMIN.Range("MoisFinAnnéeFinancière")
    End If
    If month(dateAC) = wsdADMIN.Range("MoisFinAnnéeFinancière") And day(dateAC) = day(dateFinAnneeFinanciere) Then
        Fn_TitreSelonNombreDeMois = "Pour l'exercice financier se terminant le " & Format$(dateAC, "dd mmmm yyyy")
    Else
        Fn_TitreSelonNombreDeMois = "Pour la période de " & nbMois & " mois terminée le " & Format$(dateAC, "dd mmmm yyyy")
    End If
    
End Function

Sub CalculerSoldesCourantEtComparatif(noCompteGL As String, moisCloture As Long, ligne() As Variant, _
                         ByRef soldeCourant As Currency, ByRef soldeComparatif As Currency) '2025-08-14 @ 07:54

    'Initialiser les soldes
    soldeCourant = 0
    soldeComparatif = 0

    'Traitement différent pour les postes du bilan & état des résultats
    Dim k As Long
    If noCompteGL < "4000" Then
        For k = 1 To 13
            soldeComparatif = soldeComparatif + ligne(k)
        Next k
        For k = 1 To 25
            soldeCourant = soldeCourant + ligne(k)
        Next k
    Else
        For k = (13 - moisCloture + 1) To 13
            soldeComparatif = soldeComparatif + ligne(k)
        Next k
        For k = (25 - moisCloture + 1) To 25
            soldeCourant = soldeCourant + ligne(k)
        Next k
    End If

    Debug.Print "# G/L : " & noCompteGL & " " & Right(Space(15) & Format(soldeComparatif, "#,##0.00"), 15) & " " & _
                            Right(Space(15) & Format(soldeCourant, "#,##0.00"), 15)
    
End Sub
