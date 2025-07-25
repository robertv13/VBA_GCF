Attribute VB_Name = "modGL_PrepEF"
Option Explicit

Public gDictgSoldeCodeEF As Object
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

Sub Calculer_Soldes_Pour_EF(ws As Worksheet, dateCutOff As Date) '2025-02-05 @ 04:26
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Calculer_Soldes_Pour_EF", ws.Name & ", " & dateCutOff, 0)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Qui exécute ce programme ?
    Dim isDeveloppeur As Boolean
    If modFunctions.GetNomUtilisateur() = "RobertMV" Or modFunctions.GetNomUtilisateur() = "robertmv" Then
        isDeveloppeur = True
    End If
    
    'Déterminer la date de cutoff pour l'an passé
    Dim cutOffAnPassé As Date
    cutOffAnPassé = dateCutOff
    cutOffAnPassé = DateAdd("yyyy", -1, cutOffAnPassé)
    ws.Range("F5").Value = Format$(dateCutOff, wsdADMIN.Range("B1").Value)
    ws.Range("H5").Value = Format$(cutOffAnPassé, wsdADMIN.Range("B1").Value)
    
    'The Chart of Account will drive the results, so the sort order is determined by COA
    Dim arr As Variant
    arr = Fn_Get_Plan_Comptable(4) 'Retourne un tableau avec 4 colonnes
    
    'Effacer les cellules en place (contenu & format)
    ws.Unprotect
    ws.Range("C6" & ":K" & UBound(arr, 1) + 6 + 2).ClearContents

    'Step # 1 - Use AdvancedFilter on GL_Trans for ALL accounts and transactions between the 2 dates
    Dim rngResultAF As Range
    Call modGL_Stuff.ObtenirSoldeCompteEntreDebutEtFin(vbNullString, #7/31/2024#, dateCutOff, rngResultAF)

    'The SORT method does not sort correctly the GLNo, since there is NUMBER and NUMBER+LETTER !!!
    Dim lastUsedRow As Long
    lastUsedRow = rngResultAF.Rows.count
    If lastUsedRow < 2 Then Exit Sub
    
    'Charge en mémoire (matrice) toutes les transactions du G/L
    Dim arrTrans As Variant
    arrTrans = rngResultAF.Value2
    
    'Dictionary par code d'état financier, pointe vers une matrice
    Dim dictSoldesParGL As Dictionary: Set dictSoldesParGL = New Dictionary
    Dim rowID_to_arrSoldesParGL As Long 'Pointeur à la matrice
    Dim arrSoldesParGL() As Variant 'Soldes Année Courante & Année précédente
    ReDim arrSoldesParGL(1 To UBound(arr, 1), 1 To 3)
    
    'Lire chacune des lignes de transaction du résultat (GL_Trans_AF#1)
    Dim currRowID As Long
    Dim i As Long, glNo As String, MyValue As String, t1 As Currency, t2 As Currency
    For i = 2 To UBound(arrTrans, 1)
        glNo = arrTrans(i, 5)
        If Not dictSoldesParGL.Exists(glNo) Then
            rowID_to_arrSoldesParGL = rowID_to_arrSoldesParGL + 1
            dictSoldesParGL.Add glNo, rowID_to_arrSoldesParGL
            arrSoldesParGL(rowID_to_arrSoldesParGL, 1) = glNo
        End If
        currRowID = dictSoldesParGL(glNo)
        'Mettre à jour la matrice des soldes
        arrSoldesParGL(currRowID, 2) = arrSoldesParGL(currRowID, 2) + arrTrans(i, 7) - arrTrans(i, 8)
        If CDate(arrTrans(i, 2)) <= cutOffAnPassé Then
            arrSoldesParGL(currRowID, 3) = arrSoldesParGL(currRowID, 3) + arrTrans(i, 7) - arrTrans(i, 8)
        End If
    Next i
    
    Dim currRow As Long
    ws.Range("C6:C" & UBound(arr, 1) + 7).HorizontalAlignment = xlCenter
    ws.Range("D6:D" & UBound(arr, 1) + 7).HorizontalAlignment = xlLeft
    ws.Range("E6:E" & UBound(arr, 1) + 7).HorizontalAlignment = xlCenter
    ws.Range("F6:H" & UBound(arr, 1) + 7).HorizontalAlignment = xlRight
    ws.Range("C6:H" & UBound(arr, 1) + 7).Font.Name = "Aptos Narrow"
    ws.Range("C6:H" & UBound(arr, 1) + 7).Font.size = 10
    
    'Maintenant on affiche des résulats, piloté par le plan comptable
    'Utilisation d'un dictionary pour sommariser les lignes de EF
    If Not gDictgSoldeCodeEF Is Nothing Then
        gDictgSoldeCodeEF.RemoveAll
    End If
    If gDictgSoldeCodeEF Is Nothing Then
        Set gDictgSoldeCodeEF = CreateObject("Scripting.Dictionary")
    End If
    Dim rowID_to_gSoldeCodeEF As Long
    ReDim gSoldeCodeEF(1 To UBound(arr, 1), 1 To 3)
    Dim codeEF As String
    Dim dictPreuve As Dictionary
    Set dictPreuve = New Dictionary
    
    'Dictionary de type Global
    Dim dictSectionSub As Object
    If Not dictSectionSub Is Nothing Then
        dictSectionSub.RemoveAll
    End If
    If dictSectionSub Is Nothing Then
        Set dictSectionSub = CreateObject("Scripting.Dictionary")
    End If
    Dim section As String
    
    Dim soldeAC As Currency, soldeAP As Currency, totalAC As Currency, totalAP As Currency
    Dim descGL As String
    currRow = 5
    Dim r As Long
    'arr est la matrice contenant le plan comptable
    For i = LBound(arr, 1) To UBound(arr, 1)
        glNo = arr(i, 1)
        descGL = arr(i, 2)
        codeEF = arr(i, 4)
        
        r = dictSoldesParGL.item(glNo)
        If r <> 0 Then 'r <> 0 indique qu'il y a un solde pour ce G/L
            If arrSoldesParGL(r, 2) <> 0 Or arrSoldesParGL(r, 3) <> 0 Then
                currRow = currRow + 1
                ws.Range("C" & currRow).Value = glNo
                ws.Range("D" & currRow).Value = descGL
                ws.Range("E" & currRow).Value = codeEF
                If isDeveloppeur = True Then
                    ws.Range("M" & currRow).Value = codeEF
                    ws.Range("N" & currRow).Value = glNo
                End If
                'Accumule les montants par ligne d'état financier (codeEF)
                If Not gDictgSoldeCodeEF.Exists(codeEF) Then
                    rowID_to_gSoldeCodeEF = rowID_to_gSoldeCodeEF + 1
                    gDictgSoldeCodeEF.Add codeEF, rowID_to_gSoldeCodeEF
                    gSoldeCodeEF(rowID_to_gSoldeCodeEF, 1) = codeEF
                End If
                currRowID = gDictgSoldeCodeEF(codeEF)
                
                ws.Range("F" & currRow).Value = arrSoldesParGL(r, 2)
                gSoldeCodeEF(currRowID, 2) = gSoldeCodeEF(currRowID, 2) + arrSoldesParGL(r, 2)
                totalAC = totalAC + arrSoldesParGL(r, 2)
                ws.Range("H" & currRow).Value = arrSoldesParGL(r, 3)
                gSoldeCodeEF(currRowID, 3) = gSoldeCodeEF(currRowID, 3) + arrSoldesParGL(r, 3)
                totalAP = totalAP + CCur(arrSoldesParGL(r, 3))
                
                'Preuve
                If Not dictPreuve.Exists(codeEF & "-" & glNo) Then
                    dictPreuve.Add codeEF & "-" & glNo, 0
                End If
                dictPreuve(codeEF & "-" & glNo) = dictPreuve(codeEF & "-" & glNo) + arrSoldesParGL(r, 2)
                
                'Preuve - Sous-total par section
                section = Left$(codeEF, 1)
                If Not dictSectionSub.Exists(section) Then
                    dictSectionSub.Add section, 0
                End If
                dictSectionSub(section) = dictSectionSub(section) + arrSoldesParGL(r, 2)
            End If
        End If
        
        'Sauvegarde des BNR au début de l'année et Dividendes
        If glNo = "3100" Then
            gBNR_Début_Année_AC = ws.Range("F" & currRow).Value
            gBNR_Début_Année_AP = ws.Range("H" & currRow).Value
        ElseIf glNo = "3200" Then
            gDividendes_Année_AC = ws.Range("F" & currRow).Value
            gDividendes_Année_AP = ws.Range("H" & currRow).Value
        End If
    
        If isDeveloppeur = True Then
            ws.Range("O" & currRow).Value = ws.Range("F" & currRow).Value
            ws.Range("P" & currRow).Value = ws.Range("H" & currRow).Value
        End If
    Next i

    currRow = currRow + 2
    
    'Output GL totals
    ws.Range("D" & currRow).Value = "Totaux"
    ws.Range("F" & currRow).Value = totalAC
    ws.Range("H" & currRow).Value = totalAP
    
    'Ajuste le format des montants
    ws.Range("F6:F" & currRow).NumberFormat = "###,###,##0.00 ;(###,###,##0.00);0.00"
    ws.Range("H6:H" & currRow).NumberFormat = "###,###,##0.00 ;(###,###,##0.00);0.00"
    
    ws.Protect userInterfaceOnly:=True
    ws.EnableSelection = xlUnlockedCells
    
    Application.EnableEvents = True
    
    ActiveWindow.ScrollRow = 1
    
    Application.EnableEvents = False
    ws.Range("C6").Select
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set dictPreuve = Nothing
    Set dictSoldesParGL = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Calculer_Soldes_Pour_EF", vbNullString, startTime)

End Sub

Sub shp_GL_PrepEF_Preparer_Click()

    Dim ws As Worksheet
    Set ws = wshGL_PrepEF
    
    Call Assembler_États_Financiers
    
End Sub

Sub shp_GL_PrepEF_Exit_Click()

    Call GL_PrepEF_Back_To_Menu

End Sub

Sub GL_PrepEF_Back_To_Menu()
    
    wshGL_PrepEF.Visible = xlSheetHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
End Sub

Sub Assembler_États_Financiers()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_États_Financiers", vbNullString, 0)
    
    Dim dateAC As Date, dateAP As Date
    dateAC = wshGL_PrepEF.Range("F5").Value
    dateAP = wshGL_PrepEF.Range("H5").Value
    
    Call CreerFeuillesEtFormat
    
    Call Assembler_Page_Titre_0_Main(dateAC, dateAP)
    Call Assembler_TM_0_Main(dateAC, dateAP)
    Call Assembler_ER_0_Main(dateAC, dateAP)
    Call Assembler_BNR_0_Main(dateAC, dateAP)
    Call Assembler_Bilan_0_Main(dateAC, dateAP)
    
    Dim nomsFeuilles As Variant
    nomsFeuilles = Array("Page titre", "Table des Matières", "État des Résultats", "BNR", "Bilan")
    
    Dim ws As Worksheet
    Dim i As Integer
    For i = UBound(nomsFeuilles) To LBound(nomsFeuilles) Step -1
        Set ws = ThisWorkbook.Sheets(nomsFeuilles(i)) 'Vérifier si la feuille existe déjà
        With ws
            'Sélectionner la feuille
            .Activate
            .Visible = xlSheetVisible
            'Affichage de la feuille à 87 %
            ActiveWindow.Zoom = 87
            'Afficher en mode aperçu des sauts de page
            ActiveWindow.View = xlPageBreakPreview
            'Remplir toutes les cellules avec la couleur blanche
            .Cells.Interior.Color = RGB(255, 255, 255) 'Blanc
'            .Cells.Interior.Color = RGB(255, 255, 204) ' Jaune pâle
        End With
    Next i
    
    'On se déplace à la première page des états financiers
    ActiveWorkbook.Sheets("Page Titre").Activate
    
    MsgBox "Les états financiers ont été produits" & vbNewLine & vbNewLine & _
            "Voir les onglets respectifs au bas du classeur", vbOKOnly, "Fin de traitement"
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_États_Financiers", vbNullString, startTime)

End Sub

Sub CreerFeuillesEtFormat()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:CreerFeuillesEtFormat", vbNullString, 0)
    
    'Liste des feuilles à créer
    Dim nomsFeuilles As Variant
    nomsFeuilles = Array("Page titre", "Table des Matières", "État des Résultats", "BNR", "Bilan")

    Application.ScreenUpdating = False
    
    'Création des feuilles et application des formats
    Dim ws As Worksheet
    Dim i As Integer
    For i = LBound(nomsFeuilles) To UBound(nomsFeuilles)
        On Error Resume Next
        Application.StatusBar = "Création de " & nomsFeuilles(i)
        Set ws = ThisWorkbook.Sheets(nomsFeuilles(i)) 'Vérifier si la feuille existe déjà
        On Error GoTo 0
        
        If ws Is Nothing Then ' Si la feuille n'existe pas, la créer
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)) 'ThisWorkbook.Sheets()(Sheets.count))
            ws.Name = nomsFeuilles(i)
        End If
        
        'Appliquer une mise en page standard pour toutes les feuilles
        With ThisWorkbook.Sheets(nomsFeuilles(i)).PageSetup
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

    Application.StatusBar = vbNullString
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:CreerFeuillesEtFormat", vbNullString, startTime)
    
End Sub


Sub Assembler_Page_Titre_0_Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_Page_Titre_0_Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Page Titre")
    
    Application.StatusBar = "Construction de la page titre"
        
    Call AssemblerPageTitre1EtArrièrePlanEtEntête(ws, dateAC, dateAP)
    
    Application.StatusBar = vbNullString
    
    Application.ScreenUpdating = True

    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_Page_Titre_0_Main", vbNullString, startTime)

End Sub

Sub AssemblerPageTitre1EtArrièrePlanEtEntête(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerPageTitre1EtArrièrePlanEtEntête", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter
    
    Call PositionnerCellule(ws, UCase$(wsdADMIN.Range("NomEntreprise")), 8, 2, 20, True, xlCenter)
    Call PositionnerCellule(ws, UCase$("États Financiers"), 15, 2, 20, True, xlCenter)
    Call PositionnerCellule(ws, UCase$(Format$(dateAC, "dd mmmm yyyy")), 28, 2, 20, True, xlCenter)
    
    'Ajuster la largeur des colonnes et la hauteur de lignes
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 87
    ws.Columns("C").ColumnWidth = 3
    ws.Rows("1:28").RowHeight = 20
    
    'Ajuster la police pour la feuille
    With ws.Cells
        .Font.Name = "Calibri"
        .Font.size = 20
        .Font.Color = RGB(98, 88, 80)
    End With

    'Fixer le printArea selon le nombre de lignes ET 3 colonnes
    ActiveSheet.PageSetup.PrintArea = "$A1:$C" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3

    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:AssemblerPageTitre1EtArrièrePlanEtEntête", vbNullString, startTime)

End Sub

Sub Assembler_TM_0_Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_TM_0_Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Table des Matières")
    
    Application.StatusBar = "Construction de la table des matières"
    
    Call Assembler_TM_1_Arrière_Plan_Et_Entete(ws, dateAC, dateAP)
    Call Assembler_TM_2_Lignes(ws)
    
    Application.StatusBar = vbNullString
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_TM_0_Main", vbNullString, startTime)

End Sub

Sub Assembler_TM_1_Arrière_Plan_Et_Entete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_TM_1_Arrière_Plan_Et_Entete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Appliquer le format d'en-tête
    Call PositionnerCellule(ws, UCase$(wsdADMIN.Range("NomEntreprise")), 1, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("Table des Matières"), 2, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("États Financiers"), 3, 2, 12, True, xlLeft)
    Call PositionnerCellule(ws, UCase$("Au " & Format$(dateAC, "dd mmmm yyyy")), 4, 2, 12, True, xlLeft)
    
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
    ActiveSheet.PageSetup.PrintArea = "$A1:$D" & ws.Cells(ws.Rows.count, "B").End(xlUp).Row + 3
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_TM_1_Arrière_Plan_Et_Entete", vbNullString, startTime)

End Sub

Sub Assembler_TM_2_Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_TM_2_Lignes", vbNullString, 0)
    
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
        With .Cells
            .Font.Name = "Calibri"
            .Font.size = 11
            .Font.Color = RGB(98, 88, 80)
        End With
    
    End With
    
    'Fixer le printArea selon le nombre de lignes ET 3 colonnes
    ActiveSheet.PageSetup.PrintArea = "$A1:$D" & ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_TM_2_Lignes", vbNullString, startTime)

End Sub

Sub Assembler_ER_0_Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_ER_0_Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("État des résultats")
    
    Application.StatusBar = "Construction de l'état des résultats"
    
    Call Assembler_ER_1_Arrière_Plan_Et_Entete(ws, dateAC, dateAP)
    Call Assembler_ER_2_Lignes(ws)
    
    'On ajoute le Revenu Net au BNR du bilan via variables Globales
    Dim indice As Integer
    indice = gDictgSoldeCodeEF("E02")
    gSoldeCodeEF(indice, 2) = gSoldeCodeEF(indice, 2) - gTotalRevenuNet_AC
    gSoldeCodeEF(indice, 3) = gSoldeCodeEF(indice, 3) - gTotalRevenuNet_AP
    
    Application.StatusBar = vbNullString
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_ER_0_Main", vbNullString, startTime)

End Sub

Sub Assembler_ER_1_Arrière_Plan_Et_Entete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_ER_1_Arrière_Plan_Et_Entete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Titre de l'état des résultats
    Dim jourAC As Integer, moisAC As Integer, anneeAC As Integer
    jourAC = day(dateAC)
    moisAC = month(dateAC)
    anneeAC = year(dateAC)
    Dim titre As String
    Dim nbMois As Integer
    If moisAC > wsdADMIN.Range("MoisFinAnnéeFinancière") Then
        nbMois = moisAC - wsdADMIN.Range("MoisFinAnnéeFinancière")
    Else
        nbMois = moisAC + 12 - wsdADMIN.Range("MoisFinAnnéeFinancière")
    End If
    If moisAC = wsdADMIN.Range("MoisFinAnnéeFinancière") And jourAC = DateSerial(anneeAC, moisAC + 1, 0) Then
        titre = "Pour l'exercice financier se terminant le "
    Else
        titre = "Pour la période de " & nbMois & " mois terminée le "
    End If
    titre = titre & Format$(dateAC, "dd mmmm yyyy")
    
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
        .ColorIndex = 0
        .TintAndShade = 0
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
     
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_ER_1_Arrière_Plan_Et_Entete", vbNullString, startTime)

End Sub

Sub Assembler_ER_2_Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_ER_2_Lignes", vbNullString, 0)
    
    Dim wsAdmin As Worksheet
    Set wsAdmin = wsdADMIN
    
    Dim tbl As ListObject
    Set tbl = wsAdmin.ListObjects("tblÉtatsFinanciersCodes")
    
    Dim LigneEF As String, codeEF As String, typeLigne As String, gras As String, souligne As String
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
            Call ImprimerLigneEF(ws, currRow, LigneEF, codeEF, typeLigne, gras, souligne, size)
        End If
        
    Next rngRow
    
    'Ajuster la police pour la feuille
    With ws.Cells
        .Font.Name = "Calibri"
        .Font.size = 11
        .Font.Color = RGB(98, 88, 80)
    End With

    'Transfère les montants NON arrondis dans les cellules sans les cents
    Dim i As Integer
    For i = 7 To currRow
        If ws.Range("G" & i).Value <> vbNullString Then
            ws.Range("C" & i).Value = ws.Range("G" & i).Value
            ws.Range("E" & i).Value = ws.Range("I" & i).Value
        End If
    Next i
    ws.Range("G7:I45").Clear
    
    'Tri par ordre descendant une plage
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("C17:C31"), Order:=xlDescending
        .SetRange ws.Range("B17:E31")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Fixer le printArea selon le nombre de lignes ET colonnes
    ActiveSheet.PageSetup.PrintArea = "$A1:$F" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_ER_2_Lignes", vbNullString, startTime)

End Sub

Sub Assembler_Bilan_0_Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_Bilan_0_Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Bilan")
    
    Application.StatusBar = "Construction du bilan"
    
    Call Assembler_Bilan_1_Arrière_Plan_Et_Entete(ws, dateAC, dateAP)
    Call Assembler_Bilan_2_Lignes(ws)
    
    Application.StatusBar = vbNullString
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_Bilan_0_Main", vbNullString, startTime)
    
End Sub

Sub Assembler_Bilan_1_Arrière_Plan_Et_Entete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_Bilan_1_Arrière_Plan_Et_Entete", vbNullString, 0)
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_Bilan_1_Arrière_Plan_Et_Entete", vbNullString, startTime)

End Sub

Sub Assembler_Bilan_2_Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_Bilan_2_Lignes", vbNullString, 0)
    
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
    With ws.Cells
        .Font.Name = "Calibri"
        .Font.size = 11
        .Font.Color = RGB(98, 88, 80)
    End With

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
    ActiveSheet.PageSetup.PrintArea = "$A1:$F" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_Bilan_2_Lignes", vbNullString, startTime)

End Sub

Sub Assembler_BNR_0_Main(dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_BNR_0_Main", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BNR")
    
    Application.StatusBar = "Construction de l'état des bénéfices non répartis"
    
    Call Assembler_BNR_1_Arrière_Plan_Et_Entete(ws, dateAC, dateAP)
    Call Assembler_BNR_2_Lignes(ws)
    
    Application.StatusBar = vbNullString
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_BNR_0_Main", vbNullString, startTime)
    
End Sub

Sub Assembler_BNR_1_Arrière_Plan_Et_Entete(ws As Worksheet, dateAC As Date, dateAP As Date)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_BNR_1_Arrière_Plan_Et_Entete", vbNullString, 0)
    
    'Effacer le contenu existant
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    'Titre de l'état des résultats
    Dim jourAC As Integer, moisAC As Integer, anneeAC As Integer
    jourAC = day(dateAC)
    moisAC = month(dateAC)
    anneeAC = year(dateAC)
    Dim titre As String
    Dim nbMois As Integer
    If moisAC > wsdADMIN.Range("MoisFinAnnéeFinancière") Then
        nbMois = moisAC - wsdADMIN.Range("MoisFinAnnéeFinancière")
    Else
        nbMois = moisAC + 12 - wsdADMIN.Range("MoisFinAnnéeFinancière")
    End If
    If moisAC = wsdADMIN.Range("MoisFinAnnéeFinancière") And jourAC = DateSerial(anneeAC, moisAC + 1, 0) Then
        titre = "Pour l'exercice financier se terminant le "
    Else
        titre = "Pour la période de " & nbMois & " mois terminée le "
    End If
    titre = titre & Format$(dateAC, "dd mmmm yyyy")
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_BNR_1_Arrière_Plan_Et_Entete", vbNullString, startTime)

End Sub

Sub Assembler_BNR_2_Lignes(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_BNR_2_Lignes", vbNullString, 0)
    
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
    With ws.Cells
        .Font.Name = "Calibri"
        .Font.size = 11
        .Font.Color = RGB(98, 88, 80)
    End With

    
    'Transfère les montants NON arrondis dans les cellules sans les cents
    Dim i As Integer
    For i = 7 To currRow
        If ws.Range("G" & i).Value <> vbNullString Then
            ws.Range("C" & i).Value = ws.Range("G" & i).Value
            ws.Range("E" & i).Value = ws.Range("I" & i).Value
        End If
    Next i
    ws.Range("G7:I25").Clear
    
    'Fixer le printArea selon le nombre de lignes ET colonnes
    ActiveSheet.PageSetup.PrintArea = "$A1:$F" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row + 3
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_PrepEF:Assembler_BNR_2_Lignes", vbNullString, startTime)

End Sub

Sub PositionnerCellule(ws As Worksheet, cell As String, ligne As Integer, col As Integer, points As Integer, gras As Boolean, alignement As Long)

    With ws.Cells(ligne, col)
        .Value = cell
        .Font.size = points
        .Font.Bold = gras
        .HorizontalAlignment = alignement
    End With
    
End Sub

Sub AdditionnerSoldes(r1 As Range, r2 As Range, comptes As String)

    If comptes = vbNullString Then
        Exit Sub
    End If
    
    Dim compte() As String
    compte = Split(comptes, "^")
    
    Dim i As Integer
    For i = 0 To UBound(compte, 1) - 1
        r1.Value = r1.Value + ChercherSoldes(compte(i), 1)
    Next i

    r1.Value = Round(r1.Value, 0)
    
End Sub

Function ChercherSoldes(valeur As String, colonne As Integer) As Currency

    Dim ws As Worksheet
    Set ws = wshGL_PrepEF
    
    Dim r As Range
    Set r = ws.Range("C6:C" & ws.Cells(ws.Rows.count, "C").End(xlUp).Row).Find(valeur, LookAt:=xlWhole)
    
    If Not r Is Nothing Then
        ChercherSoldes = r.offset(0, 3).Value
    Else
        ChercherSoldes = 0
    End If
    
End Function

Sub ImprimerLigneEF(ws As Worksheet, ByRef currRow As Integer, LigneEF As String, codeEF As String, typeLigne As String, gras As String, souligne As String, size As Long)
    
'    Debug.Print "#7-"; currRow; Tab(10); codeEF; Tab(18); typeLigne; Tab(25); gras; Tab(33); souligne; Tab(41); size
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
            index = gDictgSoldeCodeEF(codeEF)
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
                    .Weight = xlMedium
                End With
                With ws.Range("E" & currRow).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Color = -11511710
                    .Weight = xlMedium
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
                    .Weight = xlMedium
                End With
                With ws.Range("E" & currRow).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Color = -11511710
                    .Weight = xlMedium
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
        ws.Range("G" & currRow).Value = gTotalRevenuNet_AC
        ws.Range("I" & currRow).Value = gTotalRevenuNet_AP
        With ws.Range("C" & currRow).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -11511710
            .Weight = xlMedium
        End With
        With ws.Range("E" & currRow).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -11511710
            .Weight = xlMedium
        End With
    End If
    
    If codeEF = "B20" Then 'Dividendes
        ws.Range("G" & currRow).Value = -gDividendes_Année_AC
        ws.Range("I" & currRow).Value = -gDividendes_Année_AP
    End If
    
    If codeEF = "B50" Then 'Solde de fin (BNR)
        With ws.Range("C" & currRow).Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Color = -11511710
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With ws.Range("E" & currRow).Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Color = -11511710
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

Sub TrierDictionaryParCle(ByRef dict As Object)

    'Récupérer les clés dans un tableau
    Dim keys() As Variant, values() As Variant
    keys = dict.keys
    values = dict.items
    
    'Trier les clés et réarranger les valeurs
    Dim i As Integer, j As Integer
    Dim tempKey As Variant, tempValue As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then ' Tri ascendant (A ? Z)
                ' Échanger les clés
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
                
                ' Échanger les valeurs correspondantes
                tempValue = values(i)
                values(i) = values(j)
                values(j) = tempValue
            End If
        Next j
    Next i

    'Afficher le dictionnaire trié
    Debug.Print "Dictionnaire trié par clé :"
    For i = LBound(keys) To UBound(keys)
        Debug.Print keys(i) & " - " & Format$(values(i), "###,##0.00")
    Next i
End Sub


