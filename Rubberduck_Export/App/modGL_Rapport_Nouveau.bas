Attribute VB_Name = "modGL_Rapport_Nouveau"
Option Explicit

Public Sub GenererRapportGL_Compte(wsRapport As Worksheet, dateDebut As Date, dateFin As Date)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_Compte", "", 0)
   
    'Crée une collection pour tous les postes de GL sélectionnés
    Dim collGL_Selectionnes As Collection
    Set collGL_Selectionnes = New Collection

    'Construction de la chaîne avec séparateur ".|."
    Dim i As Integer
    For i = 0 To ufGL_Rapport.lsbComptes.ListCount - 1
        If ufGL_Rapport.lsbComptes.Selected(i) Then
            collGL_Selectionnes.Add ufGL_Rapport.lsbComptes.List(i)
        End If
    Next i
    
    'Setup report header
    Call SetUpGLReportHeadersAndColumns_Compte(wsRapport)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Process one account at the time...
    Dim item As Variant
    Dim compte As String
    Dim descGL As String
    Dim GL As String
    For Each item In collGL_Selectionnes
        compte = item
        GL = Left(compte, InStr(compte, " ") - 1)
        descGL = Right(compte, Len(compte) - InStr(compte, " "))
        'Informe l'utilisateur de la progression
        Application.StatusBar = "Traitement du compte " & GL & " - " & descGL
        'Obtenir le solde d'ouverture & les transactions
        Dim soldeOuverture As Currency
        soldeOuverture = Fn_Get_GL_Account_Balance(GL, dateDebut - 1)
        
        'Impression des résultats
        Call Print_Results_From_GL_Trans(wsRapport, GL, descGL, soldeOuverture, dateDebut, dateFin)
    
    Next item
    
    Application.StatusBar = ""
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Dim h1 As String, h2 As String, h3 As String
    h1 = wshAdmin.Range("NomEntreprise")
    h2 = "Rapport des transactions du Grand Livre"
    h3 = "(Du " & dateDebut & " au " & dateFin & ")"
    Call GL_Rapport_Wrap_Up_Compte(wsRapport, h1, h2, h3)
    
    Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_Compte", "", startTime)
    
    ufGL_Rapport.lblProgressBar = ""
    Unload ufGL_Rapport
    
    MsgBox "Le rapport a été généré avec succès", vbInformation, "Rapport des transactions du Grand Livre"

'    wshGL_Trans.Visible = xlSheetVisible
    wsRapport.Visible = xlSheetVisible
    wsRapport.Activate
    ActiveWindow.SplitRow = 2
    'Placer le curseur en haut du rapport (par exemple, cellule A3)
    wsRapport.Range("A3").Select
    
    'Libérer la mémoire
    Set collGL_Selectionnes = Nothing
    Set wsRapport = Nothing
    
End Sub

Public Sub GenererRapportGL_Ecriture(wsRapport As Worksheet, noEcritureDebut As Long, noEcritureFin As Long)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_Ecriture", "", 0)
   
    'Référence à la feuille source (les données de base)
    Dim wsSource As Worksheet
    Set wsSource = wshGL_Trans
    
    'Setup report header
    Call SetUpGLReportHeadersAndColumns_Ecriture(wsRapport)
    
    Dim rowRapport As Long
    rowRapport = 2 'Commencer à remplir les données à partir de la 2ème ligne

    'Trouver la dernière ligne de données dans la source
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).row

    'Filtrer les données par numéro d'écriture
    On Error Resume Next
    wsSource.AutoFilterMode = False
    On Error GoTo 0
    With wsSource.ListObjects("l_tbl_GL_Trans")
        .Range.AutoFilter Field:=1, Criteria1:=">=" & noEcritureDebut, Operator:=xlAnd, Criteria2:="<=" & noEcritureFin
    End With
    
    'Assigner le résultat à un range
    Dim filteredRange As Range
    On Error Resume Next
    Set filteredRange = wsSource.ListObjects("l_tbl_GL_Trans").DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    'Appliquer le tri uniquement sur les lignes visibles (par No. Écriture et Débit (D) et Crédit (D)
    If Not filteredRange Is Nothing Then
        With wsSource.Sort
            .SortFields.Clear
            .SortFields.Add2 key:=filteredRange.Columns(1), Order:=xlAscending ' Numéro d'Écriture
            .SortFields.Add2 key:=filteredRange.Columns(7), Order:=xlDescending ' Montant Débit
            .SortFields.Add2 key:=filteredRange.Columns(8), Order:=xlDescending ' Montant Crédit
            .SetRange filteredRange
            .Header = xlNo
            .Orientation = xlTopToBottom
            .Apply
        End With
    
        'Parcourir chaque ligne visible de filteredRange
        Dim row As Range
        Dim i As Long
        Dim colDesc As Integer
        Dim debit As Currency, credit As Currency
        Dim tDebit As Currency, tCredit As Currency
        Dim dateTrans As Date
        Dim currentEcriture As Long
        i = 1 'Numéro de ligne à traiter
        For Each row In filteredRange.Rows
            If Not row.Hidden Then
                'Traitement des données visibles seulement
                If row.Cells(1).value <> currentEcriture Then
                    currentEcriture = row.Cells(1, fGlTNoEntrée).value
                    'Informe l'utilisateur de la progression
                    If currentEcriture Mod 25 = 0 Then
                        Application.StatusBar = "Traitement de l'écriture numéro " & currentEcriture
                    End If
                    dateTrans = row.Cells(1, fGlTDate).value
                    rowRapport = rowRapport + 1
                    'Ajouter la ligne d'entête pour le No. Écriture
                    wsRapport.Cells(rowRapport, 1).value = row.Cells(fGlTNoEntrée).value
                    wsRapport.Cells(rowRapport, 2).value = dateTrans
                    wsRapport.Cells(rowRapport, 3).value = row.Cells(fGlTSource).value & ", " & row.Cells(fGlTDescription).value
                    wsRapport.Cells(rowRapport, 3).Font.Bold = True
                    rowRapport = rowRapport + 1 'Passer à la ligne suivante pour les détails
                End If
                'Détermine la colonne pour la description du GL et le montant
                If row.Cells(fGlTDébit).value <> 0 Then
                    debit = row.Cells(fGlTDébit).value
                    credit = 0
                    colDesc = 5
                Else
                    debit = 0
                    credit = row.Cells(fGlTCrédit).value
                    colDesc = 6
                End If
                'Ajouter les lignes de détail pour chaque compte
                wsRapport.Cells(rowRapport, 4).value = row.Cells(fGlTNoCompte).value
                wsRapport.Cells(rowRapport, colDesc).value = row.Cells(fGlTCompte).value
                wsRapport.Cells(rowRapport, 7).value = row.Cells(fGlTAutreRemarque).value
                'Déterminer s'il y a un débit ou un crédit
                If debit <> 0 Then
                    wsRapport.Cells(rowRapport, 8).value = debit
                    tDebit = tDebit + debit
                Else
                    wsRapport.Cells(rowRapport, 9).value = credit
                    tCredit = tCredit + credit
                End If
                rowRapport = rowRapport + 1
                i = i + 1
            Else
                Stop
            End If
        Next row
    End If
    
    Application.StatusBar = ""

    'Impression des totaux
    rowRapport = rowRapport + 1
    
    'Total Débit
    With wsRapport.Cells(rowRapport, 8)
        .value = tDebit
        .Font.Bold = True
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    'Total Crédit
    With wsRapport.Cells(rowRapport, 9)
        .value = tCredit
        .Font.Bold = True
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    'Désactiver le filtre
    wsSource.AutoFilterMode = False
    DoEvents
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Dim h1 As String, h2 As String, h3 As String
    h1 = wshAdmin.Range("NomEntreprise")
    h2 = "Rapport des transactions du Grand Livre par numéro d'écriture"
    h3 = "(Pour les numéros d'écriture de " & noEcritureDebut & " à " & noEcritureFin & ")"
    Call GL_Rapport_Wrap_Up_Ecriture(wsRapport, h1, h2, h3)
    
    Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_Ecriture", "", startTime)
    
    ufGL_Rapport.lblProgressBar = "Traitement de l'écriture numéro '" & currentEcriture
    Unload ufGL_Rapport
    
    MsgBox "Le rapport a été généré avec succès", vbInformation, "Rapport des transactions du Grand Livre"
    
'    wshGL_Trans.Visible = xlSheetVisible
    wsRapport.Visible = xlSheetVisible
    wsRapport.Activate
    ActiveWindow.SplitRow = 2
    'Placer le curseur en haut du rapport (par exemple, cellule A3)
    wsRapport.Range("A3").Select
    
    'Libérer la mémoire
    Set wsRapport = Nothing
    Set wsSource = Nothing
    
End Sub

Sub SetUpGLReportHeadersAndColumns_Compte(ws As Worksheet)

    'Efface le contenu de la feuille
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    With ws
        .Cells(1, 1) = "Compte"
        .Cells(1, 2) = "Date"
        .Cells(1, 3) = "Description"
        .Cells(1, 4) = "Source"
        .Cells(1, 5) = "No.Écr."
        .Cells(1, 6) = "Débit"
        .Cells(1, 7) = "Crédit"
        .Cells(1, 8) = "SOLDE"
        With .Range("A1:H1")
            .Font.Italic = True
            .Font.Bold = True
            .Font.Name = "Aptos Narrow"
            .Font.size = 9
            .HorizontalAlignment = xlCenter
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
        End With
    
        With .Columns("A")
            .ColumnWidth = 5
        End With
        
        With .Columns("B")
            .ColumnWidth = 11
            .HorizontalAlignment = xlCenter
        End With
        
        With .Columns("C")
            .ColumnWidth = 50
        End With
        
        With .Columns("D")
            .ColumnWidth = 20
        End With
        
        With .Columns("E")
            .ColumnWidth = 9
            .HorizontalAlignment = xlCenter
        End With
        
        With .Columns("F")
            .ColumnWidth = 15
        End With
        
        With .Columns("G")
            .ColumnWidth = 15
        End With
        
        With .Columns("H")
            .ColumnWidth = 15
        End With
    End With
    
End Sub

Sub SetUpGLReportHeadersAndColumns_Ecriture(ws As Worksheet)

    'Efface le contenu de la feuille
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    With ws
        .Cells(1, 1).value = "# Écriture"
        .Cells(1, 2).value = "Date"
        .Cells(1, 4).value = "# G/L"
        .Cells(1, 5).value = "Description"
        .Cells(1, 7).value = "Autre Remarque"
        .Cells(1, 8).value = "Débits"
        .Cells(1, 9).value = "Crédits"
        With .Range("A1:I1")
            .Font.Italic = True
            .Font.Bold = True
            .Font.Name = "Aptos Narrow"
            .Font.size = 9
            .HorizontalAlignment = xlCenter
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
        End With
    
        With .Columns("A")
            .ColumnWidth = 9
            .HorizontalAlignment = xlCenter
        End With
        
        With .Columns("B")
            .ColumnWidth = 12
            .HorizontalAlignment = xlCenter
        End With
        
        With .Columns("C")
            .ColumnWidth = 2
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("D")
            .ColumnWidth = 8
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("E")
            .ColumnWidth = 2
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("F")
            .ColumnWidth = 30
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("G")
            .ColumnWidth = 20
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("H")
            .ColumnWidth = 15
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00"
        End With
        
        With .Columns("I")
            .ColumnWidth = 15
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00"
        End With
        
    End With

End Sub

Sub GL_Rapport_Wrap_Up_Compte(ws As Worksheet, h1 As String, h2 As String, h3 As String)

    Application.PrintCommunication = False
    
    'Determine the active cells & setup Print Area
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "H").End(xlUp).row + 1
    Range("A3:H" & lastUsedRow).Select
    
    With ws.PageSetup
        .PrintArea = "$A$3:$H$" & lastUsedRow
        .PrintTitleRows = "$1:$2"
        .LeftMargin = Application.InchesToPoints(0.15)
        .RightMargin = Application.InchesToPoints(0.15)
        .TopMargin = Application.InchesToPoints(0.85)
        .BottomMargin = Application.InchesToPoints(0.45)
        .HeaderMargin = Application.InchesToPoints(0.15)
        .FooterMargin = Application.InchesToPoints(0.15)
        .LeftHeader = ""
        .CenterHeader = "&""-,Gras""&16" & h1 & _
                        Chr(10) & "&11" & h2 & _
                        Chr(10) & "&11" & h3
        .RightHeader = ""
        .LeftFooter = "&9&D - &T"
        .CenterFooter = ""
        .RightFooter = "&9Page &P de &N"
        .FitToPagesWide = 1
        .FitToPagesTall = 99
    End With
    
    Application.PrintCommunication = True

    'Keep header rows always displayed
    ActiveWindow.SplitRow = 2
    ws.Range("A" & lastUsedRow).Select
    
End Sub

Sub GL_Rapport_Wrap_Up_Ecriture(ws As Worksheet, h1 As String, h2 As String, h3 As String)

    Application.PrintCommunication = False
    
    'Determine the active cells & setup Print Area
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "H").End(xlUp).row + 1
    Range("A3:I" & lastUsedRow).Select
    
    Range("A3:I" & lastUsedRow).Font.Name = "Aptos Narrow"
    Range("A3:I" & lastUsedRow).Font.size = 10
    
    With ws.PageSetup
        .PrintArea = "$A$3:$I$" & lastUsedRow
        .PrintTitleRows = "$1:$2"
        .LeftMargin = Application.InchesToPoints(0.15)
        .RightMargin = Application.InchesToPoints(0.15)
        .TopMargin = Application.InchesToPoints(0.85)
        .BottomMargin = Application.InchesToPoints(0.45)
        .HeaderMargin = Application.InchesToPoints(0.15)
        .FooterMargin = Application.InchesToPoints(0.15)
        .LeftHeader = ""
        .CenterHeader = "&""-,Gras""&16" & h1 & _
                        Chr(10) & "&11" & h2 & _
                        Chr(10) & "&11" & h3
        .RightHeader = ""
        .LeftFooter = "&9&D - &T"
        .CenterFooter = ""
        .RightFooter = "&9Page &P de &N"
        .FitToPagesWide = 1
        .FitToPagesTall = 99
    End With
    
    Application.PrintCommunication = True

    'Keep header rows always displayed
    ActiveWindow.SplitRow = 2

    ws.Range("A" & lastUsedRow).Select
    
End Sub

Public Sub Print_Results_From_GL_Trans(ws As Worksheet, compte As String, descGL As String, soldeOuverture As Currency, dateDebut As Date, dateFin As Date)

    Dim lastRowUsed_AB As Long, lastRowUsed_A As Long, lastRowUsed_B As Long
    Dim saveFirstRow As Long
    Dim solde As Currency
    lastRowUsed_A = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    lastRowUsed_B = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lastRowUsed_A > lastRowUsed_B Then
        lastRowUsed_AB = lastRowUsed_A
    Else
        lastRowUsed_AB = lastRowUsed_B
    End If
    
    If lastRowUsed_AB <> 1 Then
        lastRowUsed_AB = lastRowUsed_AB + 3
    Else
        lastRowUsed_AB = lastRowUsed_AB + 2
    End If
    ws.Range("A" & lastRowUsed_AB).value = compte & " - " & descGL
    ws.Range("A" & lastRowUsed_AB).Font.Bold = True
    
    'Solde d'ouverture pour ce compte
    Dim glNo As String
    glNo = compte
    solde = soldeOuverture
    ws.Range("D" & lastRowUsed_AB).value = "Solde d'ouverture"
    
    ws.Range("H" & lastRowUsed_AB).value = solde
    ws.Range("H" & lastRowUsed_AB).Font.Bold = True
    lastRowUsed_AB = lastRowUsed_AB + 1
    saveFirstRow = lastRowUsed_AB

    Dim rngResult As Range
    Call GL_Get_Account_Trans_AF(glNo, dateDebut, dateFin, rngResult)
    
    Dim lastUsedTrans As Long
    lastUsedTrans = wshGL_Trans.Cells(wshGL_Trans.Rows.count, "P").End(xlUp).row '2024-11-08 @ 09:15
    If lastUsedTrans > 1 Then
        Dim i As Long, sumDT As Currency, sumCT As Currency
        'Read thru the rows
        For i = 2 To lastUsedTrans
            ws.Cells(lastRowUsed_AB, 2).value = wshGL_Trans.Range("Q" & i).value
            ws.Cells(lastRowUsed_AB, 2).NumberFormat = wshAdmin.Range("B1").value
            ws.Cells(lastRowUsed_AB, 3).value = wshGL_Trans.Range("R" & i).value
            ws.Cells(lastRowUsed_AB, 4).value = wshGL_Trans.Range("S" & i).value
            ws.Cells(lastRowUsed_AB, 5).value = wshGL_Trans.Range("P" & i).value
            ws.Cells(lastRowUsed_AB, 6).value = wshGL_Trans.Range("V" & i).value
            ws.Cells(lastRowUsed_AB, 6).NumberFormat = "###,###,##0.00 $"
            ws.Cells(lastRowUsed_AB, 7).value = wshGL_Trans.Range("W" & i).value
            ws.Cells(lastRowUsed_AB, 7).NumberFormat = "###,###,##0.00 $"
            
            solde = solde + CCur(wshGL_Trans.Range("V" & i).value) - CCur(wshGL_Trans.Range("W" & i).value)
            ws.Cells(lastRowUsed_AB, 8).value = solde
            
            sumDT = sumDT + wshGL_Trans.Range("V" & i).value
            sumCT = sumCT + wshGL_Trans.Range("W" & i).value
            
            lastRowUsed_AB = lastRowUsed_AB + 1
        Next i
    Else
        GoTo No_Transaction
    End If
    
No_Transaction:

    'Ajoute le formatage conditionnel pour les transactions
    With Range("B" & saveFirstRow & ":H" & lastRowUsed_AB - 1)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(LIGNE();2)=1"
        .FormatConditions(Range("B" & saveFirstRow & ":H" & lastRowUsed_AB - 1).FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.14996795556505
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
    
    ws.Range("H" & lastRowUsed_AB - 1).Font.Bold = True
    With ws.Range("F" & lastRowUsed_AB, "G" & lastRowUsed_AB)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    ws.Range("F" & lastRowUsed_AB).value = sumDT
    ws.Range("F" & lastRowUsed_AB).NumberFormat = "###,###,##0.00 $"
    ws.Range("G" & lastRowUsed_AB).value = sumCT
    ws.Range("G" & lastRowUsed_AB).NumberFormat = "###,###,##0.00 $"
    
    With ws.Range("A" & saveFirstRow - 1 & ":H" & lastRowUsed_AB).Font
        .Name = "Aptos Narrow"
        .size = 10
    End With
    
    'Libérer la mémoire
    Set rngResult = Nothing
    
End Sub


