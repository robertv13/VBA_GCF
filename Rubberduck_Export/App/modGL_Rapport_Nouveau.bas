Attribute VB_Name = "modGL_Rapport_Nouveau"
Option Explicit

Public Sub GenererRapportGL_Compte(wsRapport As Worksheet, dateDebut As Date, dateFin As Date)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_Compte", dateDebut & " @ " & dateFin, 0)
   
    'Cr�e une collection pour tous les postes de GL s�lectionn�s
    Dim collGL_Selectionnes As Collection
    Set collGL_Selectionnes = New Collection

    'Construction de la cha�ne avec s�parateur ".|."
    Dim i As Integer
    For i = 0 To ufGL_Rapport.lsbComptes.ListCount - 1
        If ufGL_Rapport.lsbComptes.Selected(i) Then
            collGL_Selectionnes.Add ufGL_Rapport.lsbComptes.List(i)
        End If
    Next i
    
    'Setup report header
    Call SetUpGLReportHeadersAndColumns_Compte(wsRapport)
    Dim rowRapport As Integer, saveFirstRow As Integer
    rowRapport = 3
    
    Application.ScreenUpdating = False
    
    'Filter et trier toutes les transactions du G/L
    Dim rngResultAll As Range
    Call GL_Get_Account_Trans_AF("", #1/1/2024#, dateFin, rngResultAll)
    
    'Process one account at the time...
    Dim GL As String, descGL As String
    Dim soldeOuverture As Currency, solde As Currency
    Dim item As Variant
    For Each item In collGL_Selectionnes
        GL = Left$(item, InStr(item, " ") - 1)
        descGL = Right$(item, Len(item) - InStr(item, " "))
        soldeOuverture = 0
        'Informe l'utilisateur de la progression
        Application.StatusBar = "Traitement du compte " & GL & " - " & descGL
        
        'Extraire les lignes pertinentes pour un compte de GL - arr()
        Dim arr() As Variant
        arr = ExtraireTransactionsPourUnCompte(rngResultAll, GL)
        Dim arrTrans() As Variant
        arrTrans = Array()
        If UBound(arr, 1) > 0 Then
            'Traitement de toutes les lignes pertinentes
            Dim j As Long, r As Long
            r = 0
            For i = 1 To UBound(arr, 1)
                If arr(i, fGlTDate) < dateDebut Then
                    soldeOuverture = soldeOuverture + arr(i, fGlTD�bit) - arr(i, fGlTCr�dit)
                Else
                    If r = 0 Then
                        ReDim arrTrans(1 To UBound(arr, 1), 1 To UBound(arr, 2))
                    End If
                    r = r + 1
                    For j = 1 To UBound(arr, 2)
                        arrTrans(r, j) = arr(i, j)
                    Next j
                End If
            Next i
            
            If r > 0 Then
                Call Array_2D_Resizer(arrTrans, r, UBound(arrTrans, 2))
            End If

        End If
        
        'Solde d'ouverture
        solde = soldeOuverture
        wsRapport.Range("A" & rowRapport).value = GL & " - " & descGL
        wsRapport.Range("A" & rowRapport).Font.Bold = True
        wsRapport.Range("D" & rowRapport).value = "Solde d'ouverture"
        wsRapport.Range("H" & rowRapport).value = solde
        wsRapport.Range("H" & rowRapport).Font.Bold = True
        rowRapport = rowRapport + 1
        
        'Impression des transactions pertinentes
        Dim sumDT As Currency, sumCT As Currency
        sumDT = 0
        sumCT = 0
        
        If UBound(arrTrans, 1) > 0 Then
            saveFirstRow = rowRapport
            For i = 1 To UBound(arrTrans, 1)
                wsRapport.Range("B" & rowRapport).value = arrTrans(i, 2)
                wsRapport.Range("B" & rowRapport).NumberFormat = wsdADMIN.Range("B1").value
                wsRapport.Range("C" & rowRapport).value = arrTrans(i, 3)
                wsRapport.Range("D" & rowRapport).value = arrTrans(i, 4)
                wsRapport.Range("E" & rowRapport).value = arrTrans(i, 1)
                wsRapport.Range("F" & rowRapport).value = arrTrans(i, 7)
'                wsRapport.Range("F" & rowRapport).NumberFormat = "###,###,##0.00"
                wsRapport.Range("G" & rowRapport).value = arrTrans(i, 8)
'                wsRapport.Range("G" & rowRapport).NumberFormat = "###,###,##0.00"
                
                solde = solde + CCur(arrTrans(i, 7)) - CCur(arrTrans(i, 8))
                wsRapport.Range("H" & rowRapport).value = solde
                
                sumDT = sumDT + arrTrans(i, 7)
                sumCT = sumCT + arrTrans(i, 8)
                rowRapport = rowRapport + 1
            Next i
        End If
        
        With wsRapport.Range("F" & rowRapport & ":G" & rowRapport)
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        wsRapport.Range("F" & rowRapport).value = sumDT
        wsRapport.Range("G" & rowRapport).value = sumCT
        
        'Ajoute le formatage conditionnel pour les transactions
        If saveFirstRow <> -1 Then
            Dim isPair As Integer 'Touujours laisser la premi�re ligne de d�tail sans surbrillance
            isPair = IIf(saveFirstRow Mod 2 = 0, 1, 0)
            With ActiveSheet.Range("B" & saveFirstRow & ":H" & rowRapport - 1)
                .FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(LIGNE();2)=" & isPair
                .FormatConditions(ActiveSheet.Range("B" & saveFirstRow & ":H" & rowRapport - 1).FormatConditions.count).SetFirstPriority
                With .FormatConditions(1).Interior
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.14996795556505
                End With
                .FormatConditions(1).StopIfTrue = False
            End With
        End If
        
        rowRapport = rowRapport + 2
        saveFirstRow = -1
        
    Next item
    
    With wsRapport.Range("A3:H" & rowRapport).Font
        .Name = "Aptos Narrow"
        .size = 10
    End With
    
    Application.StatusBar = ""
    
    Application.ScreenUpdating = True
    
    Dim h1 As String, h2 As String, h3 As String
    h1 = wsdADMIN.Range("NomEntreprise")
    h2 = "Rapport des transactions du Grand Livre"
    h3 = "(Du " & dateDebut & " au " & dateFin & ")"
    Call GL_Rapport_Wrap_Up_Compte(wsRapport, h1, h2, h3)
    
    Unload ufGL_Rapport
    
    Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_Compte", "", startTime)
    
    MsgBox "Le rapport a �t� g�n�r� avec succ�s", vbInformation, "Rapport des transactions du Grand Livre"

    wsRapport.Visible = xlSheetVisible
    wsRapport.Activate
    ActiveWindow.SplitRow = 2
    'Placer le curseur en haut du rapport (par exemple, cellule A3)
    wsRapport.Range("A3").Select
    
    'Lib�rer la m�moire
    Set collGL_Selectionnes = Nothing
    Set wsRapport = Nothing
    
End Sub

Public Sub GenererRapportGL_Ecriture(wsRapport As Worksheet, noEcritureDebut As Long, noEcritureFin As Long)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_Ecriture", noEcritureDebut & " � " & noEcritureFin, 0)
   
    'R�f�rence � la feuille source (les donn�es de base)
    Dim wsSource As Worksheet
    Set wsSource = wsdGL_Trans
    
    'Setup report header
    Call SetUpGLReportHeadersAndColumns_Ecriture(wsRapport)
    
    Application.ScreenUpdating = False
    
    Dim rowRapport As Long
    rowRapport = 2 'Commencer � remplir les donn�es � partir de la 2�me ligne

    'Trouver la derni�re ligne de donn�es dans la source
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).row

    'Filtrer les donn�es par num�ro d'�criture
    On Error Resume Next
    wsSource.AutoFilterMode = False
    On Error GoTo 0
    With wsSource.ListObjects("l_tbl_GL_Trans")
        .Range.AutoFilter Field:=1, Criteria1:=">=" & noEcritureDebut, Operator:=xlAnd, Criteria2:="<=" & noEcritureFin
    End With
    
    'Assigner le r�sultat � un range
    Dim filteredRange As Range
    On Error Resume Next
    Set filteredRange = wsSource.ListObjects("l_tbl_GL_Trans").DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    'Appliquer le tri uniquement sur les lignes visibles (par No. �criture et D�bit (D) et Cr�dit (D)
    If Not filteredRange Is Nothing Then
        With wsSource.Sort
            .SortFields.Clear
            .SortFields.Add2 key:=filteredRange.Columns(1), Order:=xlAscending ' Num�ro d'�criture
            .SortFields.Add2 key:=filteredRange.Columns(7), Order:=xlDescending ' Montant D�bit
            .SortFields.Add2 key:=filteredRange.Columns(8), Order:=xlDescending ' Montant Cr�dit
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
        i = 1 'Num�ro de ligne � traiter
        Debug.Print "X - ", ufGL_Rapport.chkDebourse, ufGL_Rapport.chkDepotClient, ufGL_Rapport.chkEJ, ufGL_Rapport.chkEncaissement, ufGL_Rapport.chkFacture, ufGL_Rapport.chkRegularisation
        For Each row In filteredRange.Rows
'            If InStr(row.Cells(fGlTSource).value, "D�BOURS�:") Or InStr(row.Cells(fGlTSource).value, "ENCAISSEMENT:") Or InStr(row.Cells(fGlTSource).value, "FACTURE:") Then Stop
            If Not row.Hidden And Fn_ValiderSiDoitImprimerTransaction(row.Cells(fGlTSource).value) = True Then
                'Traitement des donn�es visibles seulement
                If row.Cells(1).value <> currentEcriture Then
                    currentEcriture = row.Cells(1, fGlTNoEntr�e).value
                    'Informe l'utilisateur de la progression
                    If currentEcriture Mod 25 = 0 Then
                        Application.StatusBar = "Traitement de l'�criture num�ro " & currentEcriture
                    End If
                    dateTrans = row.Cells(1, fGlTDate).value
                    rowRapport = rowRapport + 1
                    'Ajouter la ligne d'ent�te pour le No. �criture
                    wsRapport.Cells(rowRapport, 1).value = row.Cells(fGlTNoEntr�e).value
                    wsRapport.Cells(rowRapport, 2).value = dateTrans
                    wsRapport.Cells(rowRapport, 3).value = row.Cells(fGlTSource).value & ", " & row.Cells(fGlTDescription).value
                    wsRapport.Cells(rowRapport, 3).Font.Bold = True
                    rowRapport = rowRapport + 1 'Passer � la ligne suivante pour les d�tails
                End If
                'D�termine la colonne pour la description du GL et le montant
                If row.Cells(fGlTD�bit).value <> 0 Then
                    debit = row.Cells(fGlTD�bit).value
                    credit = 0
                    colDesc = 5
                Else
                    debit = 0
                    credit = row.Cells(fGlTCr�dit).value
                    colDesc = 6
                End If
                'Ajouter les lignes de d�tail pour chaque compte
                wsRapport.Cells(rowRapport, 4).value = row.Cells(fGlTNoCompte).value
                wsRapport.Cells(rowRapport, colDesc).value = row.Cells(fGlTCompte).value
                wsRapport.Cells(rowRapport, 7).value = row.Cells(fGlTAutreRemarque).value
                'D�terminer s'il y a un d�bit ou un cr�dit
                If debit <> 0 Then
                    wsRapport.Cells(rowRapport, 8).value = debit
                    tDebit = tDebit + debit
                Else
                    wsRapport.Cells(rowRapport, 9).value = credit
                    tCredit = tCredit + credit
                End If
                rowRapport = rowRapport + 1
                i = i + 1
            End If
        Next row
    End If
    
    Application.StatusBar = ""

    'Impression des totaux
    rowRapport = rowRapport + 1
    
    'Total D�bit
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
    
    'Total Cr�dit
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
    
    'D�sactiver le filtre
    wsSource.AutoFilterMode = False
    DoEvents
    
    Application.ScreenUpdating = True
    
    Dim h1 As String, h2 As String, h3 As String
    h1 = wsdADMIN.Range("NomEntreprise")
    h2 = "Rapport des transactions du Grand Livre par num�ro d'�criture"
    h3 = "(Pour les num�ros d'�criture de " & noEcritureDebut & " � " & noEcritureFin & ")"
    Call GL_Rapport_Wrap_Up_Ecriture(wsRapport, h1, h2, h3)
    
    Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_Ecriture", "", startTime)
    
    Unload ufGL_Rapport
    
    MsgBox "Le rapport a �t� g�n�r� avec succ�s", vbInformation, "Rapport des transactions du Grand Livre"
    
'    wsdGL_Trans.Visible = xlSheetVisible
    wsRapport.Visible = xlSheetVisible
    wsRapport.Activate
    ActiveWindow.SplitRow = 2
    'Placer le curseur en haut du rapport (par exemple, cellule A3)
    wsRapport.Range("A3").Select
    
    'Lib�rer la m�moire
    Set wsRapport = Nothing
    Set wsSource = Nothing
    
End Sub

Public Sub GenererRapportGL_DateSaisie(wsRapport As Worksheet, dtSaisieDebut As Date, dtSaisieFin As Date)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_DateSaisie", dtSaisieDebut & " � " & dtSaisieFin, 0)
   
    'R�f�rence � la feuille source (les donn�es de base)
    Dim wsSource As Worksheet
    Set wsSource = wsdGL_Trans
    
    'Setup report header
    Call SetUpGLReportHeadersAndColumns_DateSaisie(wsRapport)
    
    Application.ScreenUpdating = False
    
    Dim rowRapport As Long
    rowRapport = 2 'Commencer � remplir les donn�es � partir de la 2�me ligne

    'Trouver la derni�re ligne de donn�es dans la source
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).row

    'Filtrer les donn�es par num�ro d'�criture
    On Error Resume Next
    wsSource.AutoFilterMode = False
    On Error GoTo 0
    With wsSource.ListObjects("l_tbl_GL_Trans")
        .Range.AutoFilter Field:=10, _
                          Criteria1:=">=" & CLng(dtSaisieDebut), _
                          Operator:=xlAnd, _
                          Criteria2:="<" & CLng(dtSaisieFin) + 1
    End With
    
    'Assigner le r�sultat � un range
    Dim filteredRange As Range
    On Error Resume Next
    Set filteredRange = wsSource.ListObjects("l_tbl_GL_Trans").DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    'Appliquer le tri uniquement sur les lignes visibles (par No. �criture et D�bit (D) et Cr�dit (D)
    If Not filteredRange Is Nothing Then
        With wsSource.Sort
            .SortFields.Clear
            .SortFields.Add2 key:=filteredRange.Columns(10), Order:=xlAscending ' Date de saisie
            .SortFields.Add2 key:=filteredRange.Columns(1), Order:=xlAscending  ' Num�ro d'�criture
            .SortFields.Add2 key:=filteredRange.Columns(7), Order:=xlDescending ' Montant D�bit
            .SortFields.Add2 key:=filteredRange.Columns(8), Order:=xlDescending ' Montant Cr�dit
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
        Dim currentTimeStamp As Double
        Debug.Print "X - ", ufGL_Rapport.chkDebourse, ufGL_Rapport.chkDepotClient, ufGL_Rapport.chkEJ, ufGL_Rapport.chkEncaissement, ufGL_Rapport.chkFacture, ufGL_Rapport.chkRegularisation
        For Each row In filteredRange.Rows
'            If InStr(row.Cells(fGlTSource).value, "D�BOURS�:") Or InStr(row.Cells(fGlTSource).value, "ENCAISSEMENT:") Or InStr(row.Cells(fGlTSource).value, "FACTURE:") Then Stop
            If Not row.Hidden And Fn_ValiderSiDoitImprimerTransaction(row.Cells(fGlTSource).value) = True Then
                'Traitement des donn�es visibles seulement
                If CDbl(row.Cells(10).value) <> currentTimeStamp Then
                    currentTimeStamp = CDbl(row.Cells(1, fGlTTimeStamp).value)
                    'Informe l'utilisateur de la progression
                    i = i + 1
                    If i Mod 25 = 0 Then
                        Application.StatusBar = "Traitement de l'�criture saisie le " & currentTimeStamp
                    End If
                    dateTrans = row.Cells(1, fGlTDate).value
                    rowRapport = rowRapport + 1
                    'Ajouter la ligne d'ent�te pour le No. �criture
                    wsRapport.Cells(rowRapport, 1).value = currentTimeStamp
                    wsRapport.Cells(rowRapport, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                    wsRapport.Cells(rowRapport, 2).value = row.Cells(fGlTNoEntr�e).value
                    wsRapport.Cells(rowRapport, 3).value = dateTrans
                    wsRapport.Cells(rowRapport, 4).value = row.Cells(fGlTSource).value & ", " & row.Cells(fGlTDescription).value
                    wsRapport.Cells(rowRapport, 4).Font.Bold = True
                    rowRapport = rowRapport + 1 'Passer � la ligne suivante pour les d�tails
                End If
                'D�termine la colonne pour la description du GL et le montant
                If row.Cells(fGlTD�bit).value <> 0 Then
                    debit = row.Cells(fGlTD�bit).value
                    credit = 0
                    colDesc = 6
                Else
                    debit = 0
                    credit = row.Cells(fGlTCr�dit).value
                    colDesc = 7
                End If
                'Ajouter les lignes de d�tail pour chaque compte
                wsRapport.Cells(rowRapport, 5).value = row.Cells(fGlTNoCompte).value
                wsRapport.Cells(rowRapport, colDesc).value = row.Cells(fGlTCompte).value
                wsRapport.Cells(rowRapport, 8).value = row.Cells(fGlTAutreRemarque).value
                'D�terminer s'il y a un d�bit ou un cr�dit
                If debit <> 0 Then
                    wsRapport.Cells(rowRapport, 9).value = debit
                    tDebit = tDebit + debit
                Else
                    wsRapport.Cells(rowRapport, 10).value = credit
                    tCredit = tCredit + credit
                End If
                rowRapport = rowRapport + 1
                i = i + 1
            End If
        Next row
    End If
    
    Application.StatusBar = ""

    'Impression des totaux
    rowRapport = rowRapport + 1
    
    'Total D�bit
    With wsRapport.Cells(rowRapport, 9)
        .value = tDebit
        .Font.Bold = True
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    'Total Cr�dit
    With wsRapport.Cells(rowRapport, 10)
        .value = tCredit
        .Font.Bold = True
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    'D�sactiver le filtre
    wsSource.AutoFilterMode = False
    
    DoEvents
    
    Application.ScreenUpdating = True
    
    Dim h1 As String, h2 As String, h3 As String
    h1 = wsdADMIN.Range("NomEntreprise")
    h2 = "Rapport des transactions du Grand Livre par date de saisie"
    h3 = "(Pour les �critures saisies entre le " & dtSaisieDebut & " et le " & dtSaisieFin & ")"
    Call GL_Rapport_Wrap_Up_DateSaisie(wsRapport, h1, h2, h3)
    
    Call Log_Record("modGL_Rapport_Nouveau:GenererRapportGL_DateSaisie", "", startTime)
    
    Unload ufGL_Rapport
    
    MsgBox "Le rapport a �t� g�n�r� avec succ�s", vbInformation, "Rapport des transactions du Grand Livre"
    
'    wsdGL_Trans.Visible = xlSheetVisible
    wsRapport.Visible = xlSheetVisible
    wsRapport.Activate
    ActiveWindow.SplitRow = 2
    'Placer le curseur en haut du rapport (par exemple, cellule A3)
    wsRapport.Range("A3").Select
    
    'Lib�rer la m�moire
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
        .Cells(1, 5) = "No.�cr."
        .Cells(1, 6) = "D�bit"
        .Cells(1, 7) = "Cr�dit"
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
            .NumberFormat = "###,###,##0.00"
        End With
        
        With .Columns("G")
            .ColumnWidth = 15
            .NumberFormat = "###,###,##0.00"
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
        .Cells(1, 1).value = "# �criture"
        .Cells(1, 2).value = "Date"
        .Cells(1, 4).value = "# G/L"
        .Cells(1, 5).value = "Description"
        .Cells(1, 7).value = "Autre Remarque"
        .Cells(1, 8).value = "D�bits"
        .Cells(1, 9).value = "Cr�dits"
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

Sub SetUpGLReportHeadersAndColumns_DateSaisie(ws As Worksheet)

    'Efface le contenu de la feuille
    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    
    With ws
        .Cells(1, 1).value = "Date Saisie"
        .Cells(1, 2).value = "# �criture"
        .Cells(1, 3).value = "Date"
        .Cells(1, 5).value = "# G/L"
        .Cells(1, 6).value = "Description"
        .Cells(1, 8).value = "Autre Remarque"
        .Cells(1, 9).value = "D�bits"
        .Cells(1, 10).value = "Cr�dits"
        With .Range("A1:J1")
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
            .ColumnWidth = 18
            .HorizontalAlignment = xlCenter
        End With
        
        With .Columns("B")
            .ColumnWidth = 9
            .HorizontalAlignment = xlCenter
        End With
        
        With .Columns("C")
            .ColumnWidth = 12
            .HorizontalAlignment = xlCenter
        End With
        
        With .Columns("D")
            .ColumnWidth = 2
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("E")
            .ColumnWidth = 8
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("F")
            .ColumnWidth = 2
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("G")
            .ColumnWidth = 30
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("H")
            .ColumnWidth = 20
            .HorizontalAlignment = xlLeft
        End With
        
        With .Columns("I")
            .ColumnWidth = 15
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00"
        End With
        
        With .Columns("J")
            .ColumnWidth = 15
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00"
        End With
        
    End With

End Sub

Sub GL_Rapport_Wrap_Up_Compte(ws As Worksheet, h1 As String, h2 As String, h3 As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport_Nouveau:GL_Rapport_Wrap_Up_Compte", "", 0)
    
    Application.PrintCommunication = False
    
    'Determine the active cells & setup Print Area
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "H").End(xlUp).row + 1
    ActiveSheet.Range("A3:H" & lastUsedRow).Select
    
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
                        Chr$(10) & "&11" & h2 & _
                        Chr$(10) & "&11" & h3
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
    
    Call Log_Record("modGL_Rapport_Nouveau:GL_Rapport_Wrap_Up_Compte", "", startTime)

End Sub

Sub GL_Rapport_Wrap_Up_Ecriture(ws As Worksheet, h1 As String, h2 As String, h3 As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport_Nouveau:GL_Rapport_Wrap_Up_Ecriture", "", 0)
    
    Application.PrintCommunication = False
    
    'Determine the active cells & setup Print Area
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "H").End(xlUp).row + 1
    ActiveSheet.Range("A3:I" & lastUsedRow).Select
    
    ActiveSheet.Range("A3:I" & lastUsedRow).Font.Name = "Aptos Narrow"
    ActiveSheet.Range("A3:I" & lastUsedRow).Font.size = 10
    
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
                        Chr$(10) & "&11" & h2 & _
                        Chr$(10) & "&11" & h3
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
    
    Call Log_Record("modGL_Rapport_Nouveau:GL_Rapport_Wrap_Up_Ecriture", "", startTime)

End Sub

Sub GL_Rapport_Wrap_Up_DateSaisie(ws As Worksheet, h1 As String, h2 As String, h3 As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport_Nouveau:GL_Rapport_Wrap_Up_DateSaisie", "", 0)
    
    Application.PrintCommunication = False
    
    'Determine the active cells & setup Print Area
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row + 1
    ActiveSheet.Range("A3:J" & lastUsedRow).Select
    
    ActiveSheet.Range("A3:J" & lastUsedRow).Font.Name = "Aptos Narrow"
    ActiveSheet.Range("A3:J" & lastUsedRow).Font.size = 10
    
    With ws.PageSetup
        .PrintArea = "$A$3:$J$" & lastUsedRow
        .PrintTitleRows = "$1:$2"
        .LeftMargin = Application.InchesToPoints(0.15)
        .RightMargin = Application.InchesToPoints(0.15)
        .TopMargin = Application.InchesToPoints(0.85)
        .BottomMargin = Application.InchesToPoints(0.45)
        .HeaderMargin = Application.InchesToPoints(0.15)
        .FooterMargin = Application.InchesToPoints(0.15)
        .LeftHeader = ""
        .CenterHeader = "&""-,Gras""&16" & h1 & _
                        Chr$(10) & "&11" & h2 & _
                        Chr$(10) & "&11" & h3
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
    
    Call Log_Record("modGL_Rapport_Nouveau:GL_Rapport_Wrap_Up_DateSaisie", "", startTime)

End Sub

Function Fn_ValiderSiDoitImprimerTransaction(ByVal source As String) As Boolean '2025-03-03 @ 10:21

    'Variable pour v�rifier si la transaction est valide
    Dim aImprimer As Boolean
    aImprimer = False

    'Traitement de la transaction selon fGlTSource et l'�tat des cases
    If InStr(source, "D�BOURS�:") = 1 Or InStr(source, "RENV/D�BOURS�:") = 1 Then
        If ufGL_Rapport.chkDebourse.value = True Then aImprimer = True
    ElseIf InStr(source, "D�P�T DE CLIENT:") = 1 Then
        If ufGL_Rapport.chkDepotClient.value = True Then aImprimer = True
    ElseIf InStr(source, "ENCAISSEMENT:") = 1 Then
        If ufGL_Rapport.chkEncaissement.value = True Then aImprimer = True
    ElseIf InStr(source, "FACTURE:") = 1 Then
        If ufGL_Rapport.chkFacture.value = True Then aImprimer = True
    ElseIf InStr(source, "R�GULARISATION:") = 1 Then
        If ufGL_Rapport.chkRegularisation.value = True Then aImprimer = True
    Else
        If ufGL_Rapport.chkEJ.value = True Then aImprimer = True
    End If

    'Retourne True si la transaction doit �tre trait�e, sinon False
    Fn_ValiderSiDoitImprimerTransaction = aImprimer
    
End Function

