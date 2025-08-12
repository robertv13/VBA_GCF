Attribute VB_Name = "modTBD"
Option Explicit

'Sub AdditionnerSoldes(r1 As Range, r2 As Range, comptes As String)
'
'    If comptes = vbNullString Then
'        Exit Sub
'    End If
'
'    Dim compte() As String
'    compte = Split(comptes, "^")
'
'    Dim i As Integer
'    For i = 0 To UBound(compte, 1) - 1
'        r1.Value = r1.Value + Fn_ChercherSoldes(compte(i), 1)
'    Next i
'
'    r1.Value = Round(r1.Value, 0)
'
'End Sub
'
'Ajustements à la feuille DB_Clients (Ajout du contactdans le nom du client)
'Sub zz_AjouterContactDansNomClient()
'
'    Declare and open the closed workbook
'    Dim wb As Workbook: Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx")
'
'    Define the worksheet you want to work with
'    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
'
'    Find the last used row with data in column A
'    Dim lastUsedRow As Long
'    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
'
'    Loop through each row starting from row 2 (headers are 1 row)
'    Dim client As String, clientID As String, contactFacturation As String
'    Dim posOpenSquareBracket As Integer, posCloseSquareBracket As Integer
'    Dim numberOpenSquareBracket As Integer, numberCloseSquareBracket As Integer
'    Dim i As Long
'    For i = 2 To lastUsedRow
'        Load data into variables
'        client = ws.Cells(i, fClntFMClientNom).Value
'        clientID = ws.Cells(i, fClntFMClientID).Value
'        contactFacturation = Trim$(ws.Cells(i, fClntFMContactFacturation).Value)
'
'        Process the data and make adjustments if necessary
'        posOpenSquareBracket = InStr(client, "[")
'        posCloseSquareBracket = InStr(client, "]")
'
'        If posOpenSquareBracket = 0 And posCloseSquareBracket = 0 Then
'            If contactFacturation <> vbNullString And InStr(client, contactFacturation) = 0 Then
'                client = Trim$(client) & " [" & contactFacturation & "]"
'                ws.Cells(i, 1).Value = client
'                Debug.Print "#065 - " & i & " - " & client
'            End If
'        End If
'
'    Next i
'
'    wb.Save
'
'    Libérer la mémoire
'    Set wb = Nothing
'    Set ws = Nothing
'
'    MsgBox "Le traitement est complété sur " & i - 1 & " lignes"
'
'End Sub
'
'Ajustements à la feuille DB_Clients (*) ---> [*]
'Sub zz_AjusterNomClientBD()
'
'    Declare and open the closed workbook
'    Dim wb As Workbook: Set wb = Workbooks.Open("C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_Entrée.xlsx")
'
'    Define the worksheet you want to work with
'    Dim ws As Worksheet: Set ws = wb.Worksheets("Clients")
'
'    Find the last used row with data in column A
'    Dim lastUsedRow As Long
'    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
'
'    Loop through each row starting from row 2 (headers are 1 row)
'    Dim client As String, clientID As String, contactFacturation As String
'    Dim posOpenParenthesis As Integer, posCloseParenthesis As Integer
'    Dim numberOpenParenthesis As Integer, numberCloseParenthesis As Integer
'    Dim i As Long
'    For i = 2 To lastUsedRow
'        Load data into variables
'        client = ws.Cells(i, fClntFMClientNom).Value
'        clientID = ws.Cells(i, fClntFMClientID).Value
'        contactFacturation = ws.Cells(i, fClntFMContactFacturation).Value
'
'        Process the data and make adjustments if necessary
'        posOpenParenthesis = InStr(client, "(")
'        posCloseParenthesis = InStr(client, ")")
'        numberOpenParenthesis = Fn_CompteOccurenceCaractere(client, "(")
'        numberCloseParenthesis = Fn_CompteOccurenceCaractere(client, ")")
'
'        If numberOpenParenthesis = 1 And numberCloseParenthesis = 1 Then
'            If posCloseParenthesis > posOpenParenthesis + 5 Then
'                client = Replace(client, "(", "[")
'                client = Replace(client, ")", "]")
'                ws.Cells(i, 1).Value = client
'                Debug.Print "#064 - " & i & " - " & client
'            End If
'        End If
'
'    Next i
'
'    wb.Save
'
'    Libérer la mémoire
'    Set wb = Nothing
'    Set ws = Nothing
'
'    MsgBox "Le traitement est complété sur " & i - 1 & " lignes"
'
'End Sub
'
'Sub zz_AnalyserImagesEnteteFactureExcel() '2025-05-27 @ 14:40
'
'    Dim dossier As String, fichier As String
'    Dim wb As Workbook, ws As Worksheet
'    Dim img As Shape
'    Dim largeurOrig As Double, hauteurOrig As Double
'    Dim largeurActuelle As Double, hauteurActuelle As Double
'    Dim cheminComplet As String
'    Dim nomImageCible As String
'
'    Demande à l'utilisateur de choisir un dossier
'    With Application.fileDialog(msoFileDialogFolderPicker)
'        .Title = "Choisissez un dossier contenant les fichiers Excel"
'        If .show <> -1 Then Exit Sub 'Annuler
'        dossier = .SelectedItems(1)
'    End With
'
'    Nom exact de l'image à trouver (ou utiliser un critère partiel)
'    nomImageCible = "Image 1" '? Modifier si nécessaire
'
'    Recherche tous les fichiers .xlsx dans le dossier
'    Dim dateSeuilMinimum As Date
'    dateSeuilMinimum = DateSerial(2024, 8, 1)
'    fichier = Dir(dossier & "\*.xlsx")
'
'    Do While fichier <> vbNullString
'        cheminComplet = dossier & "\" & fichier
'        If FileDateTime(cheminComplet) < dateSeuilMinimum Then
'            fichier = Dir
'            GoTo SkipFile
'        End If
'        Set wb = Workbooks.Open(cheminComplet, ReadOnly:=True)
'
'        On Error Resume Next
'        Set ws = wb.Worksheets(wb.Worksheets.count)
'        If ws.Name = "Activités" Then
'            GoTo SkipFile
'        End If
'        On Error GoTo 0
'
'        If Not ws Is Nothing Then
'            For Each img In ws.Shapes
'                If img.Type = msoPicture Then
'                    If img.Name = nomImageCible Then
'                        largeurActuelle = img.Width
'                        hauteurActuelle = img.Height
'
'                        Lire la taille originale estimée
'                        Call FN_LireTailleOriginaleImage(img, largeurOrig, hauteurOrig)
'
'                        Debug.Print "Fichier : " & fichier
'                        Debug.Print "  Image : " & img.Name
'                        Debug.Print "  Taille actuelle : " & largeurActuelle & " x " & hauteurActuelle
'                        Debug.Print "  Taille originale : " & largeurOrig & " x " & hauteurOrig
'                        Debug.Print String(40, "-")
'                    End If
'                End If
'            Next img
'        End If
'
'        wb.Close SaveChanges:=False
'        fichier = Dir
'SkipFile:
'    Loop
'
'    MsgBox "Analyse terminée."
'
'End Sub
'
'Public Sub RelancerSurveillance() '2025-07-02 @ 07:41
'
'    If gMODE_DEBUG Then Debug.Print "[modAppli:RelancerSurveillance] *** Surveillance relancée manuellement à " & Format(Now, "hh:mm:ss")
'
'    On Error Resume Next
'    If gFermeturePlanifiee = 0 Then
'        If gMODE_DEBUG Then Debug.Print "[modAppli:RelancerSurveillance] gFermeturePlanifiee est nul — aucun OnTime à annuler"
'    End If
'
'    Application.OnTime gFermeturePlanifiee, "FermerApplicationInactive", , False
'    Application.OnTime ufConfirmationFermeture.ProchainTick, "RelancerTimer", , False
'    On Error GoTo 0
'
'    gDerniereActivite = Now
'    Call VerifierDerniereActivite
'
'End Sub
'
'Function Fn_CreerCopieTemporaireSolide(onglet As String) As String
'
'    Dim wsSrc As Worksheet, wsDest As Worksheet
'    Dim wbTmp As Workbook
'    Dim sPath As String, sFichier As String
'    Dim oldScreenUpdating As Boolean
'    Dim lastRow As Long, lastCol As Long
'
'    On Error GoTo ErrHandler
'
'    sPath = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & "\"
'    If Dir(sPath, vbDirectory) = vbNullString Then
'        MsgBox "Le répertoire n'existe pas : " & vbCrLf & sPath, vbCritical
'        Fn_CreerCopieTemporaireSolide = vbNullString
'        Exit Function
'    End If
'
'    sFichier = "GL_Temp_" & Environ("Username") & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
'    oldScreenUpdating = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    Set wsSrc = ThisWorkbook.Worksheets(onglet)
'    Set wbTmp = Workbooks.Add(xlWBATWorksheet)
'    Set wsDest = wbTmp.Sheets(1)
'
'    ' Déterminer la zone utilisée
'    With wsSrc
'        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
'        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
'    End With
'
'    ' Copier les valeurs uniquement
'    wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(lastRow, lastCol)).Value = _
'        wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Value
'
'    ' Optionnel : nommer la feuille comme l’originale
'    On Error Resume Next: wsDest.Name = wsSrc.Name: On Error GoTo 0
'
'    ' Sauvegarde
'    Application.DisplayAlerts = False
'    wbTmp.SaveAs fileName:=sPath & sFichier, FileFormat:=xlOpenXMLWorkbook
'    wbTmp.Close SaveChanges:=False
'    Application.DisplayAlerts = True
'
'    Application.ScreenUpdating = oldScreenUpdating
'    Fn_CreerCopieTemporaireSolide = sPath & sFichier
'    Exit Function
'
'ErrHandler:
'    Application.ScreenUpdating = oldScreenUpdating
'    MsgBox "Erreur lors de la création du fichier temporaire : " & Err.description, vbCritical
'    Fn_CreerCopieTemporaireSolide = vbNullString
'
'End Function
'
'
'Fonction pour estimer la taille originale d'une image
'Sub FN_LireTailleOriginaleImage(img As Shape, ByRef largeurOrig As Double, ByRef hauteurOrig As Double)
'
'    Dim ws As Worksheet
'    Dim copie As Shape
'
'    Set ws = img.Parent
'    img.Copy
'    ws.Paste
'    Set copie = ws.Shapes(ws.Shapes.count) 'la dernière collée
'
'    With copie
'        .ScaleWidth 1, msoTrue, msoScaleFromTopLeft
'        .ScaleHeight 1, msoTrue, msoScaleFromTopLeft
'        largeurOrig = .Width
'        hauteurOrig = .Height
'        .Delete
'    End With
'
'End Sub
'
'Function Fn_CelluleAPartirColonneDansFeuille(ws As Worksheet, search As String, searchCol As Long, returnCol As Long) As Variant
'
'    Dim foundCell As Range
'
'    'Utilisation de la méthode Find pour rechercher dans la première colonne
'    Set foundCell = ws.Columns(searchCol).Find(What:=search, LookIn:=xlValues, LookAt:=xlWhole)
'
'    If Not foundCell Is Nothing Then
'        Fn_CelluleAPartirColonneDansFeuille = ws.Cells(foundCell.row, returnCol)
'    Else
'        Fn_CelluleAPartirColonneDansFeuille = vbNullString
'    End If
'
'    'Libérer la mémoire
'    Set foundCell = Nothing
'
'End Function
'
'
'Function Fn_ObtenirTECFacturesPourFacture(invNo As String) As Variant
'
'    Dim wsTEC As Worksheet: Set wsTEC = wsdTEC_Local
'
'    Dim lastUsedRow As Long
'    lastUsedRow = wsTEC.Cells(wsTEC.Rows.count, 1).End(xlUp).Row '2024-08-18 @ 06:37
'
'    Dim resultArr() As Variant
'    ReDim resultArr(1 To 1000)
'
'    Dim rowCount As Long
'    Dim i As Long
'    For i = 3 To lastUsedRow
'        If wsTEC.Cells(i, 16).Value = invNo And UCase$(wsTEC.Cells(i, 14).Value) <> "VRAI" Then
'            rowCount = rowCount + 1
'            resultArr(rowCount) = i
'        End If
'    Next i
'
'    If rowCount > 0 Then
'        ReDim Preserve resultArr(1 To rowCount)
'    End If
'
'    If rowCount = 0 Then
'        Fn_ObtenirTECFacturesPourFacture = Array()
'    Else
'        Fn_ObtenirTECFacturesPourFacture = resultArr
'    End If
'
'    'Libérer la mémoire
'    Set wsTEC = Nothing
'
'End Function
'
'
'Function Fn_PeriodeAging(age As Long, days1 As Long, days2 As Long, days3 As Long, days4 As Long)
'
'    Select Case age
'        Case Is < days1
'            Fn_PeriodeAging = 0
'        Case Is < days2
'            Fn_PeriodeAging = 1
'        Case Is < days3
'            Fn_PeriodeAging = 2
'        Case Is < days4
'            Fn_PeriodeAging = 3
'        Case Else
'            Fn_PeriodeAging = 4
'    End Select
'
'End Function
'
'Function Fn_ChercherSoldes(valeur As String, colonne As Integer) As Currency
'
'    Dim ws As Worksheet
'    Set ws = wshGL_PrepEF
'
'    Dim r As Range
'    Set r = ws.Range("C6:C" & ws.Cells(ws.Rows.count, "C").End(xlUp).Row).Find(valeur, LookAt:=xlWhole)
'
'    If Not r Is Nothing Then
'        Fn_ChercherSoldes = r.offset(0, 3).Value
'    Else
'        Fn_ChercherSoldes = 0
'    End If
'
'End Function
'
'Function Fn_NumeroSemaineSelonAnneeFinanciere(DateDonnee As Date) As Long
'
'    Dim DebutAnneeFinanciere As Date
'    DebutAnneeFinanciere = wsdADMIN.Range("AnneeDe")
'
'    'Trouver le jour de la semaine du début de l'année financière (1 = dimanche, 2 = lundi, etc.)
'    Dim JourSemaineDebut As Long
'    JourSemaineDebut = Weekday(DebutAnneeFinanciere, vbMonday)
'
'    ' Ajuster le début de l'année financière au lundi précédent si ce n'est pas un lundi
'    If JourSemaineDebut > 1 Then
'        DebutAnneeFinanciere = DebutAnneeFinanciere - (JourSemaineDebut - 1)
'    End If
'
'    ' Calculer le nombre de jours entre la date donnée et le début ajusté de l'année financière
'    Dim NbJours As Long
'    NbJours = DateDonnee - DebutAnneeFinanciere
'
'    ' Calculer le numéro de la semaine (diviser par 7 et arrondir)
'    Dim Semaine As Integer
'    Semaine = Int(NbJours / 7) + 1
'
'    ' Retourner le numéro de la semaine
'    Fn_NumeroSemaineSelonAnneeFinanciere = Semaine
'
'End Function
'
'Function Fn_ObtenirSoldeOuvertureGLAvecAF(glNo As String, d As Date) As Double
'
'    'Using AdvancedFilter # 1 in wsdGL_Trans
'
'    Fn_ObtenirSoldeOuvertureGLAvecAF = 0
'
'    Dim ws As Worksheet: Set ws = wsdGL_Trans
'
'    Application.EnableEvents = False
'
'    'Effacer les données de la dernière utilisation
'    ws.Range("M6:M10").ClearContents
'    ws.Range("M6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
'
'    'Définir le range pour la source des données en utilisant un tableau
'    Dim rngData As Range
'    Set rngData = ws.Range("l_tbl_GL_Trans[#All]")
'    ws.Range("M7").Value = rngData.Address
'
'    'Définir le range des critères
'    Dim rngCriteria As Range
'    Set rngCriteria = ws.Range("L2:N3")
'    ws.Range("L3").FormulaR1C1 = glNo
'    ws.Range("M3").FormulaR1C1 = ">=" & CLng(#7/31/2024#)
'    ws.Range("N3").FormulaR1C1 = "<" & CLng(d)
'    ws.Range("M8").Value = rngCriteria.Address
'
'    'Définir le range des résultats et effacer avant le traitement
'    Dim rngResult As Range
'    Set rngResult = ws.Range("P1").CurrentRegion
'    rngResult.offset(1, 0).Clear
'    Set rngResult = ws.Range("P1:Y1")
'    ws.Range("M9").Value = rngResult.Address
'
'    rngData.AdvancedFilter _
'                action:=xlFilterCopy, _
'                criteriaRange:=rngCriteria, _
'                CopyToRange:=rngResult, _
'                Unique:=False
'
'    'Quels sont les résultats ?
'    Dim lastUsedRow As Long
'    lastUsedRow = ws.Cells(ws.Rows.count, "P").End(xlUp).Row
'    ws.Range("M10").Value = lastUsedRow - 1 & " lignes"
'
'    Application.EnableEvents = True
'
'    'Pas de tri nécessaire pour calculer le solde
'    If lastUsedRow < 2 Then
'        Exit Function
'    End If
'
'    'Méthode plus rapide pour obtenir une somme
'    Set rngResult = ws.Range("P2:Y" & lastUsedRow)
'    Fn_ObtenirSoldeOuvertureGLAvecAF = Application.WorksheetFunction.Sum(rngResult.Columns(7)) _
'                                           - Application.WorksheetFunction.Sum(rngResult.Columns(8))
'
'    'Libérer la mémoire
'    Set rngCriteria = Nothing
'    Set rngData = Nothing
'    Set rngResult = Nothing
'    Set ws = Nothing
'
'End Function
'
'Function Fn_ServeurEstDisponible() As Boolean
'
'    DoEvents
'
'    On Error Resume Next
'    'Tester l'existence d'un fichier ou d'un répertoire sur le lecteur P:
'    Fn_ServeurEstDisponible = Dir("P:\", vbDirectory) <> vbNullString
'    On Error GoTo 0
'
'End Function
'
'Function Fn_SommePlageTableau(tableau As Variant, ligne As Long, debutCol As Long, finCol As Long) As Currency
'
'    Dim plage() As Long
'    Dim i As Long
'
'    'Construire un tableau d’indices colonnes
'    ReDim plage(1 To finCol - debutCol + 1)
'    For i = debutCol To finCol
'        plage(i - debutCol + 1) = i
'    Next i
'
'    Dim extrait As Variant
'    extrait = Application.index(tableau, ligne, plage)
'    Debug.Print Join(extrait, ", ")
'
'    'Somme via Index + Sum
'    Fn_SommePlageTableau = WorksheetFunction.Sum(Application.index(tableau, ligne, plage))
'
'End Function
'
'Function Fn_CompteOccurenceCaractere(ByVal inputString As String, ByVal charToCount As String) As Long
'
'    'Ensure charToCount is a single character
'    If Len(charToCount) <> 1 Or Len(inputString) = 0 Then
'        Fn_CompteOccurenceCaractere = -1 ' Return -1 for invalid input
'        Exit Function
'    End If
'
'    'Loop through each character in the string
'    Dim i As Long, count As Long
'    For i = 1 To Len(inputString)
'        If Mid$(inputString, i, 1) = charToCount Then
'            count = count + 1
'        End If
'    Next i
'
'    Fn_CompteOccurenceCaractere = count
'
'End Function
'
'Function Fn_DernierJourAnneeFinanciere(maxDate As Date) As Date
'
'    Dim dt As Date
'    Dim annee As Integer
'
'    'Calcul de l'année du début d'exercice
'    If month(maxDate) > wsdADMIN.Range("MoisFinAnnéeFinancière") Then
'        annee = year(maxDate) + 1
'    Else
'        annee = year(maxDate)
'    End If
'
'    'Retourner la date de la fin de l'année financière
'    dt = DateSerial(annee, wsdADMIN.Range("MoisFinAnnéeFinancière") + 1, 0)
'
'    Fn_DernierJourAnneeFinanciere = dt
'
'End Function
'
'Function Fn_CalculerMoisAnneeFinanciere(m As Long) As Long
'
'    Dim dernierMoisAnneeFinanciere As Long
'    dernierMoisAnneeFinanciere = wsdADMIN.Range("MoisFinAnnéeFinancière").Value
'
'    If m > dernierMoisAnneeFinanciere Then
'        m = m - dernierMoisAnneeFinanciere
'    Else
'        m = m + 12 - dernierMoisAnneeFinanciere
'    End If
'
'    Fn_CalculerMoisAnneeFinanciere = m
'
'End Function
'
'
