Attribute VB_Name = "modGL_BV"
Option Explicit

Public dynamicShape As Shape

Sub shpActualiserBV_Click()

    Call EffacerZoneTransactionsDetailleesBV(wshGL_BV)
    
    Call ActualiserBV

End Sub

Sub ActualiserBV() '2025-07-21 @ 13:01

    Dim ws As Worksheet: Set ws = wshGL_BV
    
    Dim dateBV As Date
    dateBV = ws.Range("J1").Value
    
    Application.ScreenUpdating = True
    Application.EnableEvents = False
    wshGL_BV.Range("C2").Value = "Au " & Format$(dateBV, wsdADMIN.Range("B1").Value)
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
    If lastUsedRow > 3 Then
        ws.Range("D4:G" & lastUsedRow).Clear
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    Dim soldes As Object
    Set soldes = CreateObject("Scripting.Dictionary")
    
    Set soldes = modGL_Stuff.Fn_SoldesParCompteAvecADO("0000", "9999", dateBV, False)
    If soldes Is Nothing Then
        MsgBox "Impossible d'obtenir les soldes par numéro de compte" & vbNewLine & vbNewLine & _
                "en date du " & Format$(dateBV, wsdADMIN.Range("B1").Value) & _
                "VEUILLEZ CONTACTER LE DÉVELOPPEUR SANS TARDER", _
                vbCritical, _
                "Les soldes ne peuvent être calculés !!!"
        
        Exit Sub
    End If
    
    Call AfficherSoldesBV(soldes)
    
    Dim dateFinExercice As Date
    dateFinExercice = Fn_DateFinExercice(dateBV)
    ws.Range("B12").Value = dateFinExercice
    If dateBV = dateFinExercice Then
        ws.Shapes("shpEcritureCloture").Visible = True
    Else
        ws.Shapes("shpEcritureCloture").Visible = False
    End If
    
    'Libérer la mémoire
    Set ws = Nothing

End Sub

Sub AfficherSoldesBV(soldes As Dictionary, Optional ligneDépart As Long = 4) '2025-06-03 @ 20:18

    Dim i As Long
    Dim ligne As Long
    Dim globalDebit As Currency
    Dim globalCredit As Currency
    Const COL_CODE = 1, COL_DESC = 2, COL_DEBIT = 3, COL_CREDIT = 4

    ligne = ligneDépart
    Application.EnableEvents = False

    'Parcours du dictionaire 'soldes'
    Dim cpte As Variant
    Dim descCompte As String
    Dim montant As Currency
    For Each cpte In soldes.keys
        montant = soldes(cpte)
'        If montant <> 0 Then
            'Montant inverse pour solder le compte
            descCompte = Fn_DescriptionAPartirNoCompte(CStr(cpte))
            wshGL_BV.Range("D" & ligne).Value = CStr(cpte)
            wshGL_BV.Range("E" & ligne).Value = descCompte
            If montant >= 0 Then
                wshGL_BV.Range("F" & ligne).Value = Format$(montant, "#,##0.00 $")
                globalDebit = globalDebit + montant
            Else
                wshGL_BV.Range("G" & ligne).Value = Format$(-montant, "#,##0.00 $")
                globalCredit = globalCredit - montant
            End If
            ligne = ligne + 1
'        End If
    Next cpte

    ' Écriture des totaux
    ligne = ligne + 1
    With wshGL_BV.Cells(ligne, 4)
        .Value = "TOTALS"
        .Font.Bold = True
    End With
    wshGL_BV.Cells(ligne, 6).Value = globalDebit
    wshGL_BV.Cells(ligne, 7).Value = globalCredit

    With wshGL_BV.Range("F" & ligne & ":G" & ligne)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        .Font.Bold = True
        .NumberFormat = "#,##0.00 $"
    End With

    wshGL_BV.Range("D" & ligneDépart & ":D" & ligne).HorizontalAlignment = xlCenter
    wshGL_BV.Range("F" & ligneDépart & ":G" & ligne).HorizontalAlignment = xlRight

    If globalDebit <> globalCredit Then
        MsgBox "Il y a une différence entre le total des débits et des crédits : " & _
               Format(globalDebit - globalCredit, "0.00"), vbExclamation
    End If

    Application.EnableEvents = True
    
End Sub

Sub AfficherTransactionsPourUnCompteDeGL(compte As String, description As String, dateMin As Date, dateMax As Date) '2025-05-27 @ 19:40 - v6.C.7 - ChatGPT

    Dim rsInit As Object
    Dim wsTrans As Worksheet, wsResult As Worksheet
    Dim strSQL As String
    Dim ligne As Long, lastRow As Long
    Dim debit As Currency, credit As Currency, solde As Currency, soldeInitial As Currency

    'Feuilles
    Set wsTrans = wsdGL_Trans
    Set wsResult = wshGL_BV

    'Compte & description (passés en paramètre)
    If compte = vbNullString Then
        MsgBox "Aucun compte sélectionné.", vbExclamation
        Exit Sub
    End If

    'Nettoyer la zone existante (M5 vers le bas) & ajuster l'entête
    Call EffacerZoneTransactionsDetailleesBV(wsResult)
    
    Application.EnableEvents = False
    wsResult.Range("L2").Value = "Du " & Format$(dateMin, wsdADMIN.Range("B1").Value) & " au " & Format$(dateMax, wsdADMIN.Range("B1").Value)
    
    'Écrire NoCompte & Description en L4
    With wsResult.Range("L4")
        .Value = compte & IIf(description <> vbNullString, " - " & description, vbNullString)
        .Font.Name = "Aptos Narrow"
        .Font.size = 10
        .Font.Bold = True
    End With
    'Sauvegarder le numéro du compte sélectionné ainsi que la description
    wshGL_BV.Range("B6").Value = compte
    wshGL_BV.Range("B7").Value = description
    Application.EnableEvents = True
    
    'Connexion ADO à MASTER
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";"

    'Calcul du solde d'ouverture avant DateMin
    dateMin = wsResult.Range("B8").Value
    Set rsInit = CreateObject("ADODB.Recordset")
    
    strSQL = "SELECT SUM(IIF(Débit IS NULL, 0, Débit)) as TotalDebit, SUM(IIF(Crédit IS NULL, 0, Crédit)) AS TotalCredit " & _
             "FROM [GL_Trans$] " & _
             "WHERE NoCompte = '" & compte & "' AND Date < #" & Format$(dateMin, "yyyy-mm-dd") & "#"
    
    rsInit.Open strSQL, conn, 1, 1
    If Not rsInit.EOF Then
        soldeInitial = modGL_Stuff.Nz(rsInit.Fields("TotalDebit").Value) - modGL_Stuff.Nz(rsInit.Fields("TotalCredit").Value)
    End If
    Debug.Print "Solde d'ouverture pour '" & compte & "' est de " & Format$(soldeInitial, "#,##0.00 $")
    rsInit.Close: Set rsInit = Nothing
    
    'Solde d'ouverture
    Application.EnableEvents = False
    With wsResult
        .Range("L4").Value = compte & IIf(description <> vbNullString, " - " & description, vbNullString)
        .Range("P4").Value = "Solde d'ouverture au " & Format(dateMin, wsdADMIN.Range("B1"))
        .Range("S4").Value = soldeInitial
        With .Range("P4:S4")
            .Font.Name = "Aptos Narrow"
            .Font.size = 9
            .Font.Bold = True
        End With
    End With
    
    solde = soldeInitial
    ligne = 1 'Commencer les écritures de transactions à la 1ère ligne du tableau
    Application.EnableEvents = True
    
    'Requête SQL complète (toutes les dates) pour le compte
    strSQL = "SELECT Date, NoEntrée, Description, Source, Débit, Crédit, AutreRemarque FROM [GL_Trans$] " & _
             "WHERE NoCompte = '" & Replace(compte, "'", "''") & "'" & _
             "AND Date >= #" & Format(dateMin, "yyyy-mm-dd") & "# " & _
             "AND Date <= #" & Format(dateMax, "yyyy-mm-dd") & "# " & _
             "ORDER BY Date, NoEntrée"
    Debug.Print "#777 - strSQL2 (Transactions pour la période) = " & strSQL
    
    'Exécuter la requête
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")
    recSet.Open strSQL, conn, 1, 1

    'Utilisation d'un tableau pour performance optimale avec ligne 'Solde ouverture'
    If Not recSet.EOF Then
        recSet.MoveLast
        Dim nbLignes As Long
        nbLignes = recSet.RecordCount
        recSet.MoveFirst

        'Tableau recevra les données à partir du rs
        Dim tableau() As Variant
        ReDim tableau(1 To nbLignes, 1 To 8) 'Colonnes M à S

        Do While Not recSet.EOF
            debit = Nz(recSet.Fields("Débit").Value)
            credit = Nz(recSet.Fields("Crédit").Value)
            solde = solde + debit - credit

            tableau(ligne, 1) = recSet.Fields("Date").Value
            tableau(ligne, 2) = recSet.Fields("NoEntrée").Value
            tableau(ligne, 3) = recSet.Fields("Description").Value
            tableau(ligne, 4) = recSet.Fields("Source").Value
            tableau(ligne, 5) = IIf(debit > 0, debit, vbNullString)
            tableau(ligne, 6) = IIf(credit > 0, credit, vbNullString)
            tableau(ligne, 7) = solde
            tableau(ligne, 8) = recSet.Fields("AutreRemarque")

            ligne = ligne + 1
            recSet.MoveNext
        Loop

        'Écriture de tableau dans la plage, en commençant à M5 - @TODO - 2025-07-11 @ 03:14
        Application.EnableEvents = False
        ActiveWindow.FreezePanes = False
        
        'Positionner la cellule d’ancrage juste à droite du volet figé
        wsResult.Activate

        wsResult.Range("M5").Resize(nbLignes, 8).Value = tableau
        
        Application.ScreenUpdating = True

        With wsResult.Range("M5:T" & (4 + nbLignes)).Font
            .Name = "Aptos Narrow"
            .size = 9
        End With
        wsResult.Range("M5:N" & (4 + nbLignes)).HorizontalAlignment = xlCenter
        wsResult.Range("S" & (4 + nbLignes)).Font.Bold = True
        Application.EnableEvents = True
    Else
        MsgBox "Aucune transaction à afficher pour ce" & vbNewLine & vbNewLine & _
                "compte, avec la période choisie", vbExclamation, "Transactions pour la période"
    End If
    
    Call AjusterAffichageTransactionsDetailleesBV
    
    Call AjouterFormeRetourEnHaut

    'Nettoyage
    recSet.Close: Set recSet = Nothing
    conn.Close: Set conn = Nothing
    
End Sub

Sub AjusterAffichageTransactionsDetailleesBV()

    Dim ws As Worksheet
    Set ws = wshGL_BV
    
    With ws
        'Date
        .Columns("M").ColumnWidth = 9
        .Columns("M").HorizontalAlignment = xlCenter
    
        'No écriture
        .Columns("N").ColumnWidth = 7
        
        'Description
        .Columns("O").ColumnWidth = 44
        
        'Source
        .Columns("P").ColumnWidth = 20
        
        'Débit et crédit
        .Columns("Q:R").ColumnWidth = 14
        
        'Solde
        .Columns("S").ColumnWidth = 15
        
        'Autre remarque
        .Columns("T").ColumnWidth = 30
    End With
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "M").End(xlUp).Row
    Dim visibleRows As Long
    visibleRows = ActiveWindow.VisibleRange.Rows.count
    If lastUsedRow > visibleRows Then
        ActiveWindow.ScrollRow = lastUsedRow - visibleRows + 5 'Move to the bottom of the worksheet
    Else
        ActiveWindow.ScrollRow = 1
    End If

    'Ajouter un fond alternatif pour faciliter la lecture, s'il y a des transactions
    If lastUsedRow >= 5 Then
        With ws.Range("M5:T" & lastUsedRow)
            On Error Resume Next
            .FormatConditions.Add _
                Type:=xlExpression, _
                Formula1:="=ET($M5<>"""";MOD(LIGNE();2)=1)"
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
            End With
            .FormatConditions(1).StopIfTrue = False
            On Error GoTo 0
        End With
    End If

End Sub

'@Description "Procédure pour obtenir les soldes en date de la fin d'année financière et"
'             "Effectuer l'écriture de clôture pour l'exercice"
Sub ComptabiliserEcritureCloture() '2025-08-04 @ 07:15

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_BV:ComptabiliserEcritureCloture", vbNullString, 0)
    
    Dim ws As Worksheet
    Set ws = wshGL_BV
    
    Dim dateCloture As Date
    dateCloture = ws.Range("B12").Value
    
    '1. Efface l'écriture si elle existe dans MASTER + Reimporter MASTER dans Local
    Call SupprimerEcritureClotureCourante(dateCloture)
    
    Call modImport.ImporterGLTransactions 'Reimporte de MASTER
    
    '2. Construire les soldes à la date de clôture
    Dim soldes As Object
    Set soldes = CreateObject("Scripting.Dictionary")
    
    Dim cheminFichier As String
    cheminFichier = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    Dim compteBNR As String
    compteBNR = Fn_NoCompteAPartirIndicateurCompte("Bénéfices Non Répartis")

    'Récupération des soldes par ADO (classeur, feuille, premierGL, dernierGL, dateLimite, rejet écriture clôture)
    Set soldes = Fn_SoldesParCompteAvecADO("4000", "9999", dateCloture, False)
    If soldes Is Nothing Then
        MsgBox "Impossible d'effectuer l'écriture de clôture pour" & vbNewLine & vbNewLine & _
                "l'exercice se terminant le " & Format$(dateCloture, wsdADMIN.Range("B1").Value) & _
                "VEUILLEZ CONTACTER LE DÉVELOPPEUR SANS TARDER", _
                vbCritical, _
                "Les soldes de clôture ne peuvent être calculés !!!"
        
        Exit Sub
    End If

    '3. Création de l'écriture à partir de soldes (dictionary)
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & cheminFichier & ";" & _
              "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    Dim cmd As Object: Set cmd = CreateObject("ADODB.Command")
              
    Dim cpte As Variant
    Dim montant As Currency
    Dim totalResultat As Currency
    Dim ecr As clsGL_Entry
    
    'Instanciation de l'écrituire globale
    Set ecr = New clsGL_Entry
    ecr.DateEcriture = dateCloture
    ecr.description = "Écriture de clôture annuelle"
    ecr.source = "Clôture Annuelle"
    
    'Parcours du dictionaire
    Dim descCompte As String
    For Each cpte In soldes.keys
        montant = soldes(cpte)
        If montant <> 0 Then
            'Montant inverse pour solder le compte
            descCompte = modFunctions.Fn_DescriptionAPartirNoCompte(CStr(cpte))
            ecr.AjouterLigne CStr(cpte), descCompte, -montant, "Générée par l'application" 'Inverse pour solder
            totalResultat = totalResultat + montant
        End If
    Next cpte

    'Ligne de contrepartie pour BNR
    If totalResultat <> 0 Then
        descCompte = Fn_DescriptionAPartirNoCompte(compteBNR)
        ecr.AjouterLigne CStr(compteBNR), descCompte, totalResultat, "Générée par l'application"
    End If
    
    Call AjouterEcritureGLADOPlusLocale(ecr, False)
    
    MsgBox "L'écriture de clôture en date du " & Format$(dateCloture, wsdADMIN.Range("B1").Value) & vbNewLine & vbNewLine & _
           "a été complétée avec succès", _
           vbInformation, _
           "Écriture ANNUELLE de clôture"
    
    ws.Shapes("shpEcritureCloture").Visible = False
    
    'Libérer la mémoire
    Set cmd = Nothing
    Set conn = Nothing
    Set ecr = Nothing
    Set soldes = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_BV:ComptabiliserEcritureCloture", vbNullString, startTime)
    
End Sub

Public Sub SupprimerEcritureClotureCourante(dateCloture As Date) '2025-07-21 @ 11:56

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_BV:SupprimerEcritureClotureCourante", vbNullString, 0)
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cheminMaster As String
    
    cheminMaster = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(cheminMaster, ReadOnly:=False)
    Set ws = wb.Sheets("GL_Trans")

    Dim i As Long
    'Boucle INVERSÉE pour supprimer l'écriture de clôture courante
    With ws
        For i = .Cells(.Rows.count, "A").End(xlUp).Row To 2 Step -1
            If .Cells(i, fGlTDate).Value = dateCloture And _
               .Cells(i, fGlTSource).Value = "Clôture Annuelle" Then
                .Rows(i).Delete
            End If
        Next i
    End With
    
    wb.Close SaveChanges:=True
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_BV:SupprimerEcritureClotureCourante", vbNullString, startTime)

End Sub

Sub shpImprimerBV_Click()

    Call ImprimerBV

End Sub

Sub ImprimerBV()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_BV:ImprimerBV", vbNullString, 0)
    
    Dim lastRow As Long
    lastRow = wshGL_BV.Cells(wshGL_BV.Rows.count, "D").End(xlUp).Row + 2
    If lastRow < 4 Then Exit Sub
    
    Dim printRange As Range
    Set printRange = wshGL_BV.Range("D1:G" & lastRow)
    
    Dim pagesRequired As Long
    pagesRequired = Int((lastRow - 1) / 60) + 1
    
    Dim shp As Shape: Set shp = wshGL_BV.Shapes("GL_BV_Print")
    shp.Visible = msoFalse
    
    Call MettreEnPageEtPrevisualiserBVOuTrans(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    'Libérer la mémoire
    Set printRange = Nothing
    Set shp = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_BV:ImprimerBV", vbNullString, startTime)

End Sub

Sub shpImprimerBVTrans_Click()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_BV:shpImprimerBVTrans_Click", vbNullString, 0)
    
    Call ImprimerBVTransactions

    Call modDev_Utils.EnregistrerLogApplication("modGL_BV:shpImprimerBVTrans_Click", vbNullString, startTime)

End Sub

Sub ImprimerBVTransactions()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("   modGL_BV:ImprimerBVTransactions", vbNullString, 0)
    
    Dim lastRow As Long
    lastRow = wshGL_BV.Cells(wshGL_BV.Rows.count, "M").End(xlUp).Row
    If lastRow < 4 Then Exit Sub
    
    Dim printRange As Range
    Set printRange = wshGL_BV.Range("L1:T" & lastRow)
    
    Dim pagesRequired As Long
    pagesRequired = Int((lastRow - 1) / 80) + 1
    
    Dim shp As Shape: Set shp = ActiveSheet.Shapes("GL_BV_Print_Trans")
    shp.Visible = msoFalse
    
    Call MettreEnPageEtPrevisualiserBVOuTrans(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    'Libérer la mémoire
    Set printRange = Nothing
    Set shp = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("   modGL_BV:ImprimerBVTransactions", vbNullString, startTime)

End Sub

Sub MettreEnPageEtPrevisualiserBVOuTrans(myPrintRange As Range, pagesTall As Long)
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("      modGL_BV:MettreEnPageEtPrevisualiserBVOuTrans", vbNullString, 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    With ActiveSheet.PageSetup
        .PaperSize = xlPaperLetter
        .Orientation = xlPortrait
        .PrintArea = myPrintRange.Address 'Parameter 1
        .FitToPagesWide = 1
        .FitToPagesTall = pagesTall 'Parameter 2
        
        'Page Header & Footer
        .CenterHeader = "&""Aptos Narrow,Gras""&18 " & wsdADMIN.Range("NomEntreprise").Value
        
        .LeftFooter = "&9&D - &T"
        .RightFooter = "&9Page &P de &N"
        
        'Page Margins
        .LeftMargin = Application.InchesToPoints(0.16)
        .RightMargin = Application.InchesToPoints(0.16)
         
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
         
        .CenterHorizontally = True
        .CenterVertically = False
         
        'Header and Footer margins
        .HeaderMargin = Application.InchesToPoints(0.16)
        .FooterMargin = Application.InchesToPoints(0.16)
        
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    wshGL_BV.PrintPreview '2024-08-15 @ 14:53
 
    Call modDev_Utils.EnregistrerLogApplication("      modGL_BV:MettreEnPageEtPrevisualiserBVOuTrans", vbNullString, startTime)
 
End Sub

Sub EffacerFormesNonRequises() '2024-08-15 @ 14:42

    Dim ws As Worksheet: Set ws = wshGL_BV
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        If InStr(shp.Name, "Rounded Rectangle ") Then
            shp.Delete
        End If
    Next shp

    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub AfficherDetailEcritureAvecForme()

    Call CreerFormeDynamiquePourAfficherDetailEcriture
    Call AjusterFormeDynamiqueBV
    Call AfficherFormeDynamiqueBV
    
End Sub

Sub CreerFormeDynamiquePourAfficherDetailEcriture()

    'Check if the shape has already been created
    If dynamicShape Is Nothing Then
        'Create the text box shape
        wshGL_BV.Unprotect
        Set dynamicShape = wshGL_BV.Shapes.AddShape(msoShapeRoundedRectangle, 2000, 100, 600, 100)
    End If

End Sub

Sub AjusterFormeDynamiqueBV()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_BV:AjusterFormeDynamiqueBV", vbNullString, 0)
    
    Dim lastResultRow As Long
    lastResultRow = wsdGL_Trans.Cells(wsdGL_Trans.Rows.count, "AC").End(xlUp).Row
    If lastResultRow < 2 Then Exit Sub
    
    Dim rowSelected As Long
    rowSelected = wshGL_BV.Range("B10").Value
    
    Dim texteOneLine As String, texteFull As String
    
    Dim i As Long, maxLength As Long
    With wsdGL_Trans
        For i = 2 To lastResultRow
            If i = 2 Then
                texteFull = "Entrée #: " & .Range("AC2").Value & Space$(43) & "(" & .Range("AL2").Value & ")" & vbCrLf
                texteFull = texteFull & "Desc    : " & .Range("AE2").Value & vbCrLf
                If Trim$(.Range("AF2").Value) <> vbNullString Then
                    texteFull = texteFull & "Source  : " & .Range("AF2").Value & vbCrLf & vbCrLf
                Else
                    texteFull = texteFull & vbCrLf
                End If
            End If
            texteOneLine = Fn_ChaineRemplie(.Range("AG" & i).Value, " ", 5, "R") & _
                            " - " & Fn_ChaineRemplie(.Range("AH" & i).Value, " ", 35, "R") & _
                            "  " & Fn_ChaineRemplie(Format$(.Range("AI" & i).Value, "#,##0.00 $"), " ", 14, "L") & _
                            "  " & Fn_ChaineRemplie(Format$(.Range("AJ" & i).Value, "#,##0.00 $"), " ", 14, "L")
            If Trim$(.Range("AF" & i).Value) = Trim$(wshGL_BV.Range("B6").Value) Then
                texteOneLine = " * " & texteOneLine
            Else
                texteOneLine = "   " & texteOneLine
            End If
            texteOneLine = Fn_ChaineRemplie(texteOneLine, " ", 79, "R")
            If Trim$(.Range("AK" & i).Value) <> vbNullString Then
                texteOneLine = texteOneLine & Trim$(.Range("AK" & i).Value)
            End If
            If Len(texteOneLine) > maxLength Then
                maxLength = Len(texteOneLine)
            End If
            texteFull = texteFull & texteOneLine & vbCrLf
        Next i
    End With
    If Right$(texteFull, Len(texteFull) - 1) = vbCrLf Then
        texteFull = Left$(texteFull, Len(texteFull) - 2)
    End If
    
    Dim dynamicShape As Shape: Set dynamicShape = wshGL_BV.Shapes("JE_Detail_Trans")

    'Set shape properties
    With dynamicShape
        .Fill.ForeColor.RGB = RGB(249, 255, 229)
        .Fill.Transparency = 0
        .Line.Weight = 2
        .Line.ForeColor.RGB = vbBlue
        .TextFrame.Characters.text = texteFull
        .TextFrame.Characters.Font.Color = vbBlack
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.size = 10
        .TextFrame.MarginLeft = 4
        .TextFrame.MarginRight = 4
        .TextFrame.MarginTop = 3
        .TextFrame.MarginBottom = 3
        If maxLength < 80 Then maxLength = 80
        .Width = ((maxLength * 6))
'            .Height = ((lastResultRow + 4) * 12) + 3 + 3
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .Left = wshGL_BV.Range("N" & rowSelected).Left + 4
        .Top = wshGL_BV.Range("N" & rowSelected + 1).Top + 4
    End With
        
    'Libérer la mémoire
    Set dynamicShape = Nothing
      
    Call modDev_Utils.EnregistrerLogApplication("modGL_BV:AjusterFormeDynamiqueBV", vbNullString, startTime)
      
End Sub

Sub AfficherFormeDynamiqueBV()

    Dim shp As Shape: Set shp = wshGL_BV.Shapes("JE_Detail_Trans")
    shp.Visible = msoTrue
    
    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub EffacerFormeDynamique()

    Dim shp As Shape: Set shp = wshGL_BV.Shapes("JE_Detail_Trans")
    shp.Visible = msoFalse

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub EffacerZoneTransactionsDetailleesBV(w As Worksheet)

    Application.EnableEvents = False
    Dim lastUsedRow As Long
    lastUsedRow = w.Cells(w.Rows.count, "M").End(xlUp).Row
    If lastUsedRow < 4 Then
        lastUsedRow = 4
    End If
    
    Application.EnableEvents = False
    w.Range("L4:T" & lastUsedRow).Clear
    Application.EnableEvents = True
    
    'Supprimer les formes 'shpRetour'
    Call SupprimerToutesFormesRetour(w)

    Application.EnableEvents = True

End Sub

Public Sub shpEcritureCloture_Click()

    Call ComptabiliserEcritureCloture

End Sub

Sub shpSortieBV_Click()

    Dim ws As Worksheet
    Set ws = wshGL_BV
    
    Call EffacerZoneTransactionsDetailleesBV(ws)
    Call EffacerZoneBV(ws)
    Call SupprimerToutesFormesRetour(ws)
    
    ws.Shapes("shpEcritureCloture").Visible = False
    
    DoEvents
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call RetournerMenuGL

End Sub

Sub RetournerMenuGL()
    
    Call EffacerFormesNonRequises
    
    wshGL_BV.Visible = xlSheetHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
End Sub


