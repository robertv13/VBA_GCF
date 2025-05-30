Attribute VB_Name = "modGL_BV"
Option Explicit

Public dynamicShape As Shape

Sub shp_GL_BV_Actualiser_Click()

'    Dim ws As Worksheet
'    Set ws = wshGL_BV
'
    Call GL_Trial_Balance_Build

End Sub

Sub GL_Trial_Balance_Build() '2025-05-27 @ 18:03 - v6.C.7 - ChatPGT

    Dim wsResult As Worksheet
    Dim rs As ADODB.Recordset
    Dim ligne As Long
    Dim dateMin As Date, dateMax As Date
    Dim rsKey As String
    Dim totalDebit As Double, totalCredit As Double
    Dim globalDebit As Double, globalCredit As Double
    Dim tblData() As Variant
    Dim key As Variant
    Dim i As Long

    Set wsResult = wshGL_BV
    
    Application.EnableEvents = False
    
    'Zone BV
    Call GL_BV_EffacerZoneBV(wsResult)
    
    'Zone transactions détaillées
    Call GL_BV_EffacerZoneTransactionsDetaillees(wsResult)
    
    'Formes 'shpRetour'
    Call GL_BV_SupprimerToutesLesFormes_shpRetour(wsResult)
    
    'Dates à traiter
    dateMin = DateSerial(2024, 7, 31)
    dateMax = wsResult.Range("J1").value

    'Add the cut-off date in the header (printing purposes)
    Application.EnableEvents = False
    wsResult.Range("C2").value = "Au " & Format$(dateMax, wsdADMIN.Range("B1").value)
    wsResult.Range("L2").value = "Du " & Format$(dateMin, wsdADMIN.Range("B1").value) & " au " & Format$(dateMax, wsdADMIN.Range("B1").value)
    Application.EnableEvents = True
    
    wsResult.Range("T2").value = "Mois"
    DoEvents
    
    'Dictionnaire des comptes du Plan Comptable
    Dim dPlanComptable As Object: Set dPlanComptable = CreateObject("Scripting.Dictionary")
    Dim arr As Variant
    arr = Fn_Get_Plan_Comptable(2) 'Returns array with 2 columns (Code, Description)
    For i = 1 To UBound(arr, 1)
        If Not dPlanComptable.Exists(arr(i, 1)) Then
            dPlanComptable.Add arr(i, 1), arr(i, 2)
        End If
    Next i

    'Lecture des transactions par compte via ADO
    Set rs = Get_Summary_By_GL_Account(dateMin, dateMax)

    Dim solde As Currency
    Dim tDebit As Currency, tCredit As Currency
    Dim dSoldeParGL As Object: Set dSoldeParGL = CreateObject("Scripting.Dictionary")
    
    Do While Not rs.EOF
        rsKey = rs.Fields("NoCompte").value
        tDebit = Nz(rs.Fields("TotalDébit"))
        tCredit = Nz(rs.Fields("TotalCrédit"))
        If tDebit <> 0 Or tCredit <> 0 Then
            If Not dSoldeParGL.Exists(rsKey) Then
                dSoldeParGL.Add rsKey, Array(tDebit, tCredit)
            End If
        End If
        rs.MoveNext
    Loop
    
    'Créer tableau pour tri
    Dim soldes As Variant
    ReDim tblData(1 To dPlanComptable.count, 1 To 4)
    i = 1
    For Each key In dPlanComptable.keys
        soldes = Array(0, 0)
        If dSoldeParGL.Exists(key) Then
            soldes = dSoldeParGL(key)
            solde = soldes(0) - soldes(1)
        Else
            solde = 0
        End If

        If soldes(0) <> 0 Or soldes(1) <> 0 Then
            tblData(i, 1) = key
            tblData(i, 2) = dPlanComptable(key)
            If solde >= 0 Then
                tblData(i, 3) = solde
            Else
                tblData(i, 4) = -solde
            End If
            i = i + 1
        End If
    Next key
    
    'Enlève les lignes qui n'ont pas MINIMALEMENT un débit ou un crédit
    Dim tblFinal() As Variant
    Dim j As Long
    If i > 1 Then
        ReDim tblFinal(1 To i - 1, 1 To 4)
        For j = 1 To i - 1
            tblFinal(j, 1) = tblData(j, 1)
            tblFinal(j, 2) = tblData(j, 2)
            tblFinal(j, 3) = tblData(j, 3)
            tblFinal(j, 4) = tblData(j, 4)
        Next j
        'Utilisez tblFinal à la place de tblData
        tblData = tblFinal
    Else
        Erase tblData
    End If
    Erase tblFinal
    
    'Trier tableau par NoCompte
    Call Sort2DArray(tblData, 1, True)

    'Écrire résultats + calculer totaux
    ligne = 4
    Application.EnableEvents = False
    For i = 1 To UBound(tblData, 1)
        wsResult.Cells(ligne, 4).Resize(1, 4).value = Array(tblData(i, 1), tblData(i, 2), tblData(i, 3), tblData(i, 4))
        globalDebit = globalDebit + tblData(i, 3)
        globalCredit = globalCredit + tblData(i, 4)
        ligne = ligne + 1
    Next i

    'Totaux
    ligne = ligne + 1
    With wsResult.Cells(ligne, 4)
        .value = "TOTALS"
        .Font.Bold = True
    End With
    wsResult.Cells(ligne, 6).value = globalDebit
    wsResult.Cells(ligne, 7).value = globalCredit
    
    With wsResult.Range("F" & ligne & ":" & "G" & ligne)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        .Font.Bold = True
        .NumberFormat = "#,##0.00 $"
    End With
    
    wsResult.Range("D4:D" & ligne).HorizontalAlignment = xlCenter

    'Vérification intégrité (DT ?= CT)
    If Round(globalDebit, 2) <> Round(globalCredit, 2) Then
        MsgBox "Il y a une différence entre le total des débits et le total des crédits : " & Format(globalDebit - globalCredit, "0.00"), vbExclamation
    End If
    
    Application.EnableEvents = True '2025-05-27 @ 20:02
    
End Sub

Function Get_Summary_By_GL_Account(dateMin As Date, dateMax As Date) As ADODB.Recordset '2025-05-27 @ 17:51 - v6.C.7 - ChatPGT

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:Get_Summary_By_GL_Account", "", 0)
    
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String

    'Fichier actif
    Dim sWBPath As String
    sWBPath = ThisWorkbook.FullName

    'Requête SQL
    strSQL = "SELECT [NoCompte], " & _
           "SUM([Débit]) AS TotalDébit, " & _
           "SUM([Crédit]) AS TotalCrédit " & _
           "FROM [GL_Trans$] " & _
           "WHERE [Date] >= #" & Format(dateMin, "yyyy-mm-dd") & "# " & _
           "AND [Date] <= #" & Format(dateMax, "yyyy-mm-dd") & "# " & _
           "GROUP BY [NoCompte] " & _
           "ORDER BY [NoCompte];"
    Debug.Print "Calcul de la BV" & vbNewLine & "Get_Summary_By_GL_Account - strSQL = '" & strSQL & "'"
    
    'Connexion ADO
    Set cn = New ADODB.Connection
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & sWBPath & ";" & _
                            "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
        .Open
    End With

    'Recordset
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenStatic, adLockReadOnly

    'Retour
    Set Get_Summary_By_GL_Account = rs

    'Libérer
    Set cn = Nothing
    
    Call Log_Record("modGL_BV:Get_Summary_By_GL_Account", "", startTime)
    
End Function

Private Sub Sort2DArray(arr As Variant, sortColumn As Long, ascending As Boolean) '2025-05-27 @ 18:05 - v6.C.7 - ChatPGT

    Dim i As Long, j As Long
    Dim temp As Variant
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If (ascending And arr(i, sortColumn) > arr(j, sortColumn)) _
            Or (Not ascending And arr(i, sortColumn) < arr(j, sortColumn)) Then
                temp = arr(i, 1)
                arr(i, 1) = arr(j, 1)
                arr(j, 1) = temp

                temp = arr(i, 2): arr(i, 2) = arr(j, 2): arr(j, 2) = temp
                temp = arr(i, 3): arr(i, 3) = arr(j, 3): arr(j, 3) = temp
                temp = arr(i, 4): arr(i, 4) = arr(j, 4): arr(j, 4) = temp
            End If
        Next j
    Next i

End Sub

Sub GL_BV_Display_Trans_For_Selected_Account(compte As String, description As String, dateMin As Date, dateMax As Date) '2025-05-27 @ 19:40 - v6.C.7 - ChatGPT

    Dim cn As Object, rs As Object, rsInit As Object
    Dim wsTrans As Worksheet, wsResult As Worksheet
    Dim strSQL As String
    Dim ligne As Long, lastRow As Long
    Dim debit As Currency, credit As Currency, solde As Currency, soldeInitial As Currency

    'Feuilles
    Set wsTrans = wsdGL_Trans
    Set wsResult = wshGL_BV

    'Compte & description (passés en paramètre)
    If compte = "" Then
        MsgBox "Aucun compte sélectionné.", vbExclamation
        Exit Sub
    End If

    'Nettoyer la zone existante (M5 vers le bas) & ajuster l'entête
    Call GL_BV_EffacerZoneTransactionsDetaillees(wsResult)
    
    Application.EnableEvents = False
    wsResult.Range("L2").value = "Du " & Format$(dateMin, wsdADMIN.Range("B1").value) & " au " & Format$(dateMax, wsdADMIN.Range("B1").value)
    
    'Écrire NoCompte & Description en L4
    With wsResult.Range("L4")
        .value = compte & IIf(description <> "", " - " & description, "")
        .Font.Name = "Aptos Narrow"
        .Font.size = 10
        .Font.Bold = True
    End With
    'Sauvegarder le numéro du compte sélectionné ainsi que la description
    wshGL_BV.Range("B6").value = compte
    wshGL_BV.Range("B7").value = description
    Application.EnableEvents = True
    
    'Connexion ADO
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & ThisWorkbook.FullName & ";" & _
                            "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";"
        .Open
    End With

    'Calcul du solde initial avant DateMin
    dateMin = wsResult.Range("B8").value
'    wsResult.Range("L2").value = "Du " & Format$(dateMin, wsdADMIN.Range("B1").value) & " au " & Format$(dateMax, wsdADMIN.Range("B1").value)
    Set rsInit = CreateObject("ADODB.Recordset")
    
    strSQL = "SELECT SUM(Débit) AS TotalDebit, SUM(Crédit) AS TotalCredit FROM [GL_Trans$] " & _
             "WHERE NoCompte = '" & compte & "'AND Date < #" & Format(dateMin, "mm/dd/yyyy") & "#"
    Debug.Print "GL_BV_Display_Trans_For_Selected_Account - strSQL1 = '" & strSQL & "'"
    
    rsInit.Open strSQL, cn, 1, 1
    If Not rsInit.EOF Then
        soldeInitial = Nz(rsInit.Fields("TotalDebit").value) - Nz(rsInit.Fields("TotalCredit").value)
    End If
    rsInit.Close: Set rsInit = Nothing
    
    'Requête SQL complète (toutes les dates) pour le compte
    strSQL = "SELECT Date, NoEntrée, Description, Source, Débit, Crédit, AutreRemarque FROM [GL_Trans$] " & _
             "WHERE NoCompte = '" & compte & "'" & _
             "AND Date >= #" & Format(dateMin, "yyyy-mm-dd") & "# " & _
             "AND Date <= #" & Format(dateMax, "yyyy-mm-dd") & "# " & _
             "ORDER BY Date, NoEntrée"
    Debug.Print "GL_BV_Display_Trans_For_Selected_Account - strSQL2 = '" & strSQL & "'"
    
    'Exécuter la requête
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, cn, 1, 1

    'Utilisation d'un tableau pour performance optimale avec ligne 'Solde ouverture'
    If Not rs.EOF Then
        rs.MoveLast
        Dim nbLignes As Long
        nbLignes = rs.RecordCount
        rs.MoveFirst

        'Tableau recevra les données à partir du rs
        Dim tableau() As Variant
        ReDim tableau(1 To nbLignes, 1 To 8) 'Colonnes M à S

        'Solde d'ouverture
        Application.EnableEvents = False
        wsResult.Range("P4").value = "Solde d'ouverture au " & Format(dateMin, wsdADMIN.Range("B1"))
        wsResult.Range("S4").value = soldeInitial
        With wsResult.Range("P4:S4")
            .Font.Name = "Aptos Narrow"
            .Font.size = 9
            .Font.Bold = True
        End With
        solde = soldeInitial
        ligne = 1 'Commencer les écritures de transactions à la 1ère ligne du tableau
        Application.EnableEvents = True

        Do While Not rs.EOF
            debit = Nz(rs.Fields("Débit").value)
            credit = Nz(rs.Fields("Crédit").value)
            solde = solde + debit - credit

            tableau(ligne, 1) = rs.Fields("Date").value
            tableau(ligne, 2) = rs.Fields("NoEntrée").value
            tableau(ligne, 3) = rs.Fields("Description").value
            tableau(ligne, 4) = rs.Fields("Source").value
            tableau(ligne, 5) = IIf(debit > 0, debit, "")
            tableau(ligne, 6) = IIf(credit > 0, credit, "")
            tableau(ligne, 7) = solde
            tableau(ligne, 8) = rs.Fields("AutreRemarque")

            ligne = ligne + 1
            rs.MoveNext
        Loop

        'Écriture de tableau dans la plage, en commençant à M5
        Application.EnableEvents = False
        wsResult.Range("M5").Resize(nbLignes, 8).value = tableau
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
        Application.EnableEvents = False
        wsResult.Range("L4").value = ""
        Application.EnableEvents = True
    End If
    
    Call GL_BV_AjustementAffichageTransactionsDetaillees
    
    Call GL_BV_Ajouter_Shape_Retour

    'Nettoyage
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing
    
End Sub

Sub GL_BV_AjustementAffichageTransactionsDetaillees()

    'Ajuster la largeur des colonnes de la section
    Dim ws As Worksheet
    Set ws = wshGL_BV
    
    Dim rng As Range
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "M").End(xlUp).row
    
    'Date
    Set rng = ws.Range("M5:M" & lastUsedRow)
    rng.ColumnWidth = 9
    rng.HorizontalAlignment = xlCenter
    
    'No écriture
    Set rng = ws.Range("N5:N" & lastUsedRow)
    rng.ColumnWidth = 7
    
    'Description
    Set rng = ws.Range("O5:O" & lastUsedRow)
    rng.ColumnWidth = 45
    
    'Source
    Set rng = ws.Range("P5:P" & lastUsedRow)
    rng.ColumnWidth = 20
    
    'Débit & Crédit
    Set rng = ws.Range("Q5:R" & lastUsedRow)
    rng.ColumnWidth = 14
    
    'Solde
    Set rng = ws.Range("S5:S" & lastUsedRow)
    rng.ColumnWidth = 15
    
    'Autre remarque
    Set rng = ws.Range("T5:T" & lastUsedRow)
    rng.ColumnWidth = 30

    Dim visibleRows As Long
    visibleRows = ActiveWindow.visibleRange.Rows.count
    If lastUsedRow > visibleRows Then
        ActiveWindow.ScrollRow = lastUsedRow - visibleRows + 5 'Move to the bottom of the worksheet
    Else
        ActiveWindow.ScrollRow = 1
    End If

    'Ajouter un fond alternatif pour faciliter la lecture
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

End Sub

'CommentOut - 2025-05-27 @ 17:50 - v6.C.7
'Sub GL_Trial_Balance_Build(ws As Worksheet, dateCutOff As Date) '2024-11-18 @ 07:50
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_Trial_Balance_Build", ws.Name & " " & dateCutOff, 0)
'
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'
'    'Clear TB cells - Contents & formats
'    Dim lastUsedRow As Long
'    lastUsedRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row
'    ws.Unprotect '2024-08-24 @ 16:38
'    Application.EnableEvents = False
'    ws.Range("D4" & ":G" & lastUsedRow + 2).Clear
'    Application.EnableEvents = True
'
'    'Clear Detail transaction section
'    ws.Range("L4").CurrentRegion.offset(3, 0).Clear
'
'    'Add the cut-off date in the header (printing purposes)
'    Dim minDate As Date
'    ws.Range("C2").value = "Au " & Format$(dateCutOff, wsdADMIN.Range("B1").value)
'
'    Application.EnableEvents = False
'    ws.Range("B2").value = 3
'    ws.Range("B10").value = 0
'    Application.EnableEvents = True
'
'    'Step # 1 - Use AdvancedFilter on GL_Trans for ALL accounts and transactions between the 2 dates
'    Dim rngResultAF As Range
'    Call GL_Get_Account_Trans_AF("", #7/31/2024#, dateCutOff, rngResultAF)
'
'    'The SORT method does not sort correctly the GLNo, since there is NUMBER and NUMBER+LETTER !!!
'    lastUsedRow = rngResultAF.Rows.count
'    If lastUsedRow < 2 Then Exit Sub
'
'    'The Chart of Account will drive the results, so the sort order is determined by COA
'    Dim arr As Variant
'    arr = Fn_Get_Plan_Comptable(2) 'Returns array with 2 columns (Code, Description)
'
'    Dim dictSoldesParGL As Dictionary: Set dictSoldesParGL = New Dictionary
'    Dim arrSolde() As Variant 'GLbalances
'    ReDim arrSolde(1 To UBound(arr, 1), 1 To 2)
'    Dim newRowID As Long: newRowID = 1
'    Dim currRowID As Long
'
'    'Parse every line of the result (AdvancedFilter in GL_Trans)
'    Dim i As Long, glNo As String, MyValue As String, t1 As Currency, t2 As Currency
'    For i = 2 To lastUsedRow
'        glNo = rngResultAF.Cells(i, 5)
'        If Not dictSoldesParGL.Exists(glNo) Then
'            dictSoldesParGL.Add glNo, newRowID
'            arrSolde(newRowID, 1) = glNo
'            newRowID = newRowID + 1
'        End If
'        currRowID = dictSoldesParGL(glNo)
'        'Update the summary array
'        arrSolde(currRowID, 2) = arrSolde(currRowID, 2) + rngResultAF.Cells(i, 7).value - rngResultAF.Cells(i, 8).value
'    Next i
'
'    t1 = Application.WorksheetFunction.Sum(rngResultAF.Columns(7))
'    t2 = Application.WorksheetFunction.Sum(rngResultAF.Columns(8))
'
'    Dim sumDT As Currency, sumCT As Currency, GLNoPlusDesc As String
'    Dim currRow As Long: currRow = 4
'    ws.Range("D4:D" & UBound(arrSolde, 1)).HorizontalAlignment = xlCenter
'    ws.Range("F4:G" & UBound(arrSolde, 1) + 3).HorizontalAlignment = xlRight
'
'    Dim r As Long
'    For i = LBound(arr, 1) To UBound(arr, 1)
'        glNo = arr(i, 1)
'        If glNo <> "" Then
'            r = dictSoldesParGL.item(glNo) 'Get the value of the item associated with GLNo
'            If r <> 0 Then
'                ws.Range("D" & currRow).value = glNo
'                ws.Range("E" & currRow).value = arr(i, 2)
'                If arrSolde(r, 2) >= 0 Then
'                    ws.Range("F" & currRow).value = Format$(arrSolde(r, 2), "###,###,##0.00")
'                    sumDT = sumDT + arrSolde(r, 2)
'                Else
'                    ws.Range("G" & currRow).value = Format$(-arrSolde(r, 2), "###,###,##0.00")
'                    sumCT = sumCT - arrSolde(r, 2)
'                End If
'                currRow = currRow + 1
'            End If
'        End If
'    Next i
'
'    currRow = currRow + 1
'    ws.Range("B2").value = currRow
'
'    'Unprotect the active cells of the TB area
'    With ws '2024-08-21 @ 07:10
'        .Unprotect
'        .Range("D4:G" & currRow - 2).Locked = False
'        .Protect UserInterfaceOnly:=True
'        .EnableSelection = xlUnlockedCells
'    End With
'
'    'Output Debit total
'    Dim rng As Range
'    Set rng = ws.Range("F" & currRow)
'    Call GL_BV_Display_TB_Totals(rng, sumDT) 'Débit total - 2024-06-09 @ 07:51
'
'    'Output Credit total
'    Set rng = ws.Range("G" & currRow)
'    Call GL_BV_Display_TB_Totals(rng, sumCT) 'Débit total - 2024-06-09 @ 07:51
'
'    'Setup page for printing purposes
'    Dim CenterHeaderTxt As String
'    CenterHeaderTxt = wsdADMIN.Range("NomEntreprise")
'    With ActiveSheet.PageSetup
'        .CenterHeader = "&""Calibri,Bold""&16 " & CenterHeaderTxt
'        .PrintArea = "$D$1:$G$" & currRow
'        .Orientation = xlPortrait
'        .FitToPagesWide = 1
'        .FitToPagesTall = 1
'    End With
'
'    Application.EnableEvents = True
'
'    ActiveWindow.ScrollRow = 4
'
'    Application.EnableEvents = False
'    ws.Range("C4").Select
'    Application.EnableEvents = True
'
'    'Libérer la mémoire
'    Set dictSoldesParGL = Nothing
'    Set rng = Nothing
'
'    Call Log_Record("modGL_BV:GL_Trial_Balance_Build", "", startTime)
'
'End Sub
'
'Sub GL_BV_Display_TB_Totals(rng As Range, t As Currency) '2024-06-09 @ 07:45
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Display_TB_Totals", "", 0)
'
'    With rng
'        With .Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .ColorIndex = 0
'            .TintAndShade = 0
'            .Weight = xlThin
'        End With
'        With .Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .ColorIndex = 0
'            .TintAndShade = 0
'            .Weight = xlThick
'        End With
'        .value = t
'        .Font.Bold = True
'        .NumberFormat = "#,##0.00 $"
'    End With
'
'    Call Log_Record("modGL_BV:GL_BV_Display_TB_Totals", "", startTime)
'
'End Sub
'
'Sub GL_BV_Display_Trans_For_Selected_Account(GLAcct As String, GLDesc As String, minDate As Date, maxDate As Date) 'Display GL Trans for a specific account
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Display_Trans_For_Selected_Account", GLAcct & " De " & minDate & " à " & maxDate, 0)
'
'    Dim ws As Worksheet: Set ws = wshGL_BV
'
'    'Clear the display area & display the account number & description
'    With ws
'        .Range("L4:T99999").Clear '2024-06-08 @ 15:28
'        .Range("L2").value = "Du " & Format$(minDate, wsdADMIN.Range("B1").value) & " au " & Format$(maxDate, wsdADMIN.Range("B1").value)
'
'        .Range("L4").Font.Bold = True
'        .Range("L4").value = GLAcct & " - " & GLDesc
'        .Range("B6").value = GLAcct
'        .Range("B7").value = GLDesc
'    End With
'
'    'Use the AdvancedFilter Result already prepared for TB
'    Dim row As Range, foundRow As Long, lastResultUsedRow As Long
'    lastResultUsedRow = wsdGL_Trans.Cells(wsdGL_Trans.Rows.count, "P").End(xlUp).row
'    If lastResultUsedRow <= 2 Then
'        GoTo Exit_Sub
'    End If
'    foundRow = 0
'
'    'Find the first occurence of GlACct in AdvancedFilter Results on GL_Trans
'    Dim searchRange As Range: Set searchRange = wsdGL_Trans.Range("T1:T" & lastResultUsedRow)
'    Dim foundCell As Range: Set foundCell = searchRange.Find(What:=GLAcct, LookIn:=xlValues, LookAt:=xlWhole)
'    foundRow = foundCell.row
'
'    'Check if the target value was found
'    If foundRow = 0 Then
'        MsgBox "Il n'existe aucune transaction pour ce compte (période choisie)."
'        Exit Sub
'    End If
'
'    Dim rowGLDetail As Long
'    rowGLDetail = 5
'    With ws.Range("S4")
'        .value = 0
'        .Font.Bold = True
'        .NumberFormat = "#,##0.00 $"
'        With .Interior
'            .Pattern = xlSolid
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorDark1
'            .TintAndShade = -0.149998474074526
'            .PatternTintAndShade = 0
'        End With
'    End With
'
'    Dim d As Date, OK As Long
'
'    Application.ScreenUpdating = False
'
'    With ws
'        'On assume que les résultats de GL_Trans sont triés par numéro de compte, par date & par no écriture
'        Do Until wsdGL_Trans.Range("T" & foundRow).value <> GLAcct
'            'Traitement des transactions détaillées
'            d = Format$(wsdGL_Trans.Range("Q" & foundRow).Value2, wsdADMIN.Range("B1").value)
'            If d >= minDate And d <= maxDate Then
'                .Range("M" & rowGLDetail).value = wsdGL_Trans.Range("Q" & foundRow).Value2
'                .Range("M" & rowGLDetail).NumberFormat = wsdADMIN.Range("B1").value
'                .Range("N" & rowGLDetail).value = wsdGL_Trans.Range("P" & foundRow).value
'                .Range("N" & rowGLDetail).HorizontalAlignment = xlCenter
'                .Range("O" & rowGLDetail).value = wsdGL_Trans.Range("R" & foundRow).value
'                .Range("P" & rowGLDetail).value = wsdGL_Trans.Range("S" & foundRow).value
'                .Range("Q" & rowGLDetail).NumberFormat = "#,##0.00"
'                .Range("Q" & rowGLDetail).value = wsdGL_Trans.Range("V" & foundRow).value
'                .Range("R" & rowGLDetail).NumberFormat = "#,##0.00"
'                .Range("R" & rowGLDetail).value = wsdGL_Trans.Range("W" & foundRow).value
'                .Range("S" & rowGLDetail).value = ws.Range("S" & rowGLDetail - 1).value + _
'                    wsdGL_Trans.Range("V" & foundRow).value - wsdGL_Trans.Range("W" & foundRow).value
'                .Range("T" & rowGLDetail).Value2 = wsdGL_Trans.Range("X" & foundRow).value
'                foundRow = foundRow + 1
'                rowGLDetail = rowGLDetail + 1
'                OK = OK + 1
'            Else
'                foundRow = foundRow + 1
'            End If
'        Loop
'    End With
'
'    With ws.Range("S" & rowGLDetail - 1)
'        .Font.Bold = True
'        With .Interior
'            .Pattern = xlSolid
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorDark1
'            .TintAndShade = -0.149998474074526
'            .PatternTintAndShade = 0
'        End With
'    End With
'
'    Dim rng As Range
'    lastResultUsedRow = ws.Cells(ws.Rows.count, "M").End(xlUp).row
'    Set rng = ws.Range("M5:T" & lastResultUsedRow)
'
'    'Fix font size & Family for the detailled transactions list
'    Call Fix_Font_Size_And_Family(rng, "Aptos Narrow", 9)
'
'    'Set columns width for the detailled transactions list
'    Set rng = ws.Range("M5:M" & lastResultUsedRow)
'    rng.ColumnWidth = 9
'    rng.HorizontalAlignment = xlCenter
'
'    Set rng = ws.Range("N5:N" & lastResultUsedRow)
'    rng.ColumnWidth = 7
'    Set rng = ws.Range("O5:O" & lastResultUsedRow)
'    rng.ColumnWidth = 45
'    Set rng = ws.Range("P5:P" & lastResultUsedRow)
'    rng.ColumnWidth = 20
'    Set rng = ws.Range("Q5:S" & lastResultUsedRow)
'    rng.ColumnWidth = 15
'    Set rng = ws.Range("T5:T" & lastResultUsedRow)
'    rng.ColumnWidth = 35
'
'    Dim visibleRows As Long
'    visibleRows = ActiveWindow.visibleRange.Rows.count
'    If lastResultUsedRow > visibleRows Then
'        ActiveWindow.ScrollRow = lastResultUsedRow - visibleRows + 5 'Move to the bottom of the worksheet
'    Else
'        ActiveWindow.ScrollRow = 1
'    End If
'
'    'Create a Conditional Formating for the displayed transactions
'    ws.Unprotect
'    With ws.Range("M5:T" & lastResultUsedRow)
'        On Error Resume Next
'        .FormatConditions.Add _
'            Type:=xlExpression, _
'            Formula1:="=ET($M5<>"""";MOD(LIGNE();2)=1)"
'        .FormatConditions(.FormatConditions.count).SetFirstPriority
'        With .FormatConditions(1).Interior
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorAccent1
'            .TintAndShade = 0.799981688894314
'        End With
'        .FormatConditions(1).StopIfTrue = False
'        On Error GoTo 0
'    End With
'
'    'Unprotect the active cells of the transactions details area
'    With wshGL_BV '2024-08-21 @ 07:15
'        .Unprotect
'        .Range("L4:T" & lastResultUsedRow).Locked = False
'        .Protect UserInterfaceOnly:=True
'        .EnableSelection = xlUnlockedCells
'    End With
'
'    Call GL_BV_Ajouter_Shape_Retour
'
'Exit_Sub:
'
'    Application.ScreenUpdating = True
'
'    'Libérer la mémoire
'    Set foundCell = Nothing
'    Set rng = Nothing
'    Set searchRange = Nothing
'    Set ws = Nothing
'
'    Call Log_Record("modGL_BV:GL_BV_Display_Trans_For_Selected_Account", GLAcct & " De " & minDate & " à " & maxDate, startTime)
'
'End Sub
'
'Sub GL_BV_Sub_Totals(glNo As String, GLDesc As String, s As Currency)
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Sub_Totals", "", 0)
'
'    Dim r As Long
'    With wshGL_BV
'        r = .Range("B2").value + 1
'        .Range("D" & r).HorizontalAlignment = xlCenter
'        .Range("D" & r).value = glNo
'        .Range("E" & r).value = GLDesc
'        If s > 0 Then
'            .Range("F" & r).value = s
'        ElseIf s < 0 Then
'            .Range("G" & r).value = -s
'        End If
'        .Range("B2").value = wshGL_BV.Range("B2").value + 1
'    End With
'
'    Call Log_Record("modGL_BV:GL_BV_Sub_Totals", "", startTime)
'
'End Sub
'
'Sub GL_BV_ChoixPériodeAImprimer()
'
'    Range("T2").Select
'
'    Select Case period
'        Case "Mois"
'            wshGL_BV.Range("B8").value = wsdADMIN.Range("MoisDe").value
'            wshGL_BV.Range("B9").value = wsdADMIN.Range("MoisA").value
'        Case "Mois dernier"
'            wshGL_BV.Range("B8").value = wsdADMIN.Range("MoisPrecDe").value
'            wshGL_BV.Range("B9").value = wsdADMIN.Range("MoisPrecA").value
'        Case "Trimestre"
'            wshGL_BV.Range("B8").value = wsdADMIN.Range("TrimDe").value
'            wshGL_BV.Range("B9").value = wsdADMIN.Range("TrimA").value
'        Case "Trimestre dernier"
'            wshGL_BV.Range("B8").value = wsdADMIN.Range("TrimPrecDe").value
'            wshGL_BV.Range("B9").value = wsdADMIN.Range("TrimPrecA").value
'        Case "Année"
'            wshGL_BV.Range("B8").value = wsdADMIN.Range("AnneeDe").value
'            wshGL_BV.Range("B9").value = wsdADMIN.Range("AnneeA").value
'        Case "Année dernière"
'            wshGL_BV.Range("B8").value = wsdADMIN.Range("AnneePrecDe").value
'            wshGL_BV.Range("B9").value = wsdADMIN.Range("AnneePrecA").value
'        Case "Dates Manuelles"
'            wshGL_BV.Range("B8").value = CDate(Format$("07-31-2024", "dd/mm/yyyy"))
'            wshGL_BV.Range("B9").value = CDate(Format$("07-31-2025", "dd/mm/yyyy"))
'        Case "Toutes les dates"
'            wshGL_BV.Range("B8").value = CDate(Format$(wshGL_BV.Range("B3").value, "dd/mm/yyyy"))
'            wshGL_BV.Range("B9").value = CDate(Format$(wshGL_BV.Range("B4").value, "dd/mm/yyyy"))
'    End Select
'
'End Sub
'
Sub shp_GL_BV_Impression_BV_Click()

    Call GL_BV_Setup_And_Print

End Sub

Sub GL_BV_Setup_And_Print()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Setup_And_Print", "", 0)
    
    Dim lastRow As Long
    lastRow = wshGL_BV.Cells(wshGL_BV.Rows.count, "D").End(xlUp).row + 2
    If lastRow < 4 Then Exit Sub
    
    Dim printRange As Range
    Set printRange = wshGL_BV.Range("D1:G" & lastRow)
    
    Dim pagesRequired As Long
    pagesRequired = Int((lastRow - 1) / 60) + 1
    
    Dim shp As Shape: Set shp = wshGL_BV.Shapes("GL_BV_Print")
    shp.Visible = msoFalse
    
    Call GL_BV_SetUp_And_Print_Document(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    'Libérer la mémoire
    Set printRange = Nothing
    Set shp = Nothing
    
    Call Log_Record("modGL_BV:GL_BV_Setup_And_Print", "", startTime)

End Sub

Sub shp_GL_BV_Setup_And_Print_Trans_Click()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:shp_GL_BV_Setup_And_Print_Trans_Click", "", 0)
    
    Call GL_BV_Setup_And_Print_Trans

    Call Log_Record("modGL_BV:shp_GL_BV_Setup_And_Print_Trans_Click", "", startTime)

End Sub

Sub GL_BV_Setup_And_Print_Trans()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_Setup_And_Print_Trans", "", 0)
    
    Dim lastRow As Long
    lastRow = wshGL_BV.Cells(wshGL_BV.Rows.count, "M").End(xlUp).row
    If lastRow < 4 Then Exit Sub
    
    Dim printRange As Range
    Set printRange = wshGL_BV.Range("L1:T" & lastRow)
    
    Dim pagesRequired As Long
    pagesRequired = Int((lastRow - 1) / 80) + 1
    
    Dim shp As Shape: Set shp = ActiveSheet.Shapes("GL_BV_Print_Trans")
    shp.Visible = msoFalse
    
    Call GL_BV_SetUp_And_Print_Document(printRange, pagesRequired)
    
    shp.Visible = msoTrue
    
    'Libérer la mémoire
    Set printRange = Nothing
    Set shp = Nothing
    
    Call Log_Record("modGL_BV:GL_BV_Setup_And_Print_Trans", "", startTime)

End Sub

Sub GL_BV_SetUp_And_Print_Document(myPrintRange As Range, pagesTall As Long)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:GL_BV_SetUp_And_Print_Document", "", 0)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    With ActiveSheet.PageSetup
'        .PrintTitleRows = ""
'        .PrintTitleColumns = ""
        .PaperSize = xlPaperLetter
        .Orientation = xlPortrait
        .PrintArea = myPrintRange.Address 'Parameter 1
        .FitToPagesWide = 1
        .FitToPagesTall = pagesTall 'Parameter 2
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1 is completed", -1)
        
        'Page Header & Footer
'        .LeftHeader = ""
        .CenterHeader = "&""Aptos Narrow,Gras""&18 " & wsdADMIN.Range("NomEntreprise").value
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1.A is completed", -1)
        
'        .RightHeader = ""
        .LeftFooter = "&9&D - &T"
'        .CenterFooter = ""
        .RightFooter = "&9Page &P de &N"
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1.B is completed", -1)
        
        'Page Margins
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 is starting", -1)
        .LeftMargin = Application.InchesToPoints(0.16)
        .RightMargin = Application.InchesToPoints(0.16)
         Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Left & Right) margins", -1)
         
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
         Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Top & Bottom) margins", -1)
         
        .CenterHorizontally = True
        .CenterVertically = False
         Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Center Horizontal & Vertical)", -1)
         
        'Header and Footer margins
        .HeaderMargin = Application.InchesToPoints(0.16)
        .FooterMargin = Application.InchesToPoints(0.16)
        Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Header & Footer) margins", -1)
        
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    Call Log_Record("   modGL_BV:GL_BV_SetUp_And_Print_Document - Speed Measure", -1)
    
    wshGL_BV.PrintPreview '2024-08-15 @ 14:53
 
    Call Log_Record("modGL_BV:GL_BV_SetUp_And_Print_Document", "", startTime)
 
End Sub

Sub Erase_Non_Required_Shapes() '2024-08-15 @ 14:42

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

Sub Test_Get_All_Shapes() '2024-08-15 @ 14:42

    Dim ws As Worksheet: Set ws = wshGL_BV
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub wshGL_BV_Display_JE_Trans_With_Shape()

    Call wshGL_BV_Create_Dynamic_Shape
    Call wshGL_BV_Adjust_The_Shape
    Call GL_BV_Show_Dynamic_Shape
    
End Sub

Sub wshGL_BV_Create_Dynamic_Shape()

    'Check if the shape has already been created
    If dynamicShape Is Nothing Then
        'Create the text box shape
        wshGL_BV.Unprotect
        Set dynamicShape = wshGL_BV.Shapes.AddShape(msoShapeRoundedRectangle, 2000, 100, 600, 100)
    End If

    'Libérer la mémoire
'    Set dynamicShape = Nothing
    
End Sub

Sub wshGL_BV_Adjust_The_Shape()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_BV:wshGL_BV_Adjust_The_Shape", "", 0)
    
    Dim lastResultRow As Long
    lastResultRow = wsdGL_Trans.Cells(wsdGL_Trans.Rows.count, "AC").End(xlUp).row
    If lastResultRow < 2 Then Exit Sub
    
    Dim rowSelected As Long
    rowSelected = wshGL_BV.Range("B10").value
    
    Dim texteOneLine As String, texteFull As String
    
    Dim i As Long, maxLength As Long
    With wsdGL_Trans
        For i = 2 To lastResultRow
            If i = 2 Then
                texteFull = "Entrée #: " & .Range("AC2").value & vbCrLf
                texteFull = texteFull & "Desc    : " & .Range("AE2").value & vbCrLf
                If Trim$(.Range("AF2").value) <> "" Then
                    texteFull = texteFull & "Source  : " & .Range("AF2").value & vbCrLf & vbCrLf
                Else
                    texteFull = texteFull & vbCrLf
                End If
            End If
            texteOneLine = Fn_Pad_A_String(.Range("AG" & i).value, " ", 5, "R") & _
                            " - " & Fn_Pad_A_String(.Range("AH" & i).value, " ", 35, "R") & _
                            "  " & Fn_Pad_A_String(Format$(.Range("AI" & i).value, "#,##0.00 $"), " ", 14, "L") & _
                            "  " & Fn_Pad_A_String(Format$(.Range("AJ" & i).value, "#,##0.00 $"), " ", 14, "L")
            If Trim$(.Range("AF" & i).value) = Trim$(wshGL_BV.Range("B6").value) Then
                texteOneLine = " * " & texteOneLine
            Else
                texteOneLine = "   " & texteOneLine
            End If
            texteOneLine = Fn_Pad_A_String(texteOneLine, " ", 79, "R")
            If Trim$(.Range("AK" & i).value) <> "" Then
                texteOneLine = texteOneLine & Trim$(.Range("AK" & i).value)
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
        .Line.Weight = 2
        .Line.ForeColor.RGB = vbBlue
        .TextFrame.Characters.Text = texteFull
        .TextFrame.Characters.Font.Color = vbBlack
        .TextFrame.Characters.Font.Name = "Consolas"
        .TextFrame.Characters.Font.size = 10
        .TextFrame.MarginLeft = 4
        .TextFrame.MarginRight = 4
        .TextFrame.MarginTop = 3
        .TextFrame.MarginBottom = 3
        If maxLength < 80 Then maxLength = 80
        .Width = ((maxLength * 6.1))
'            .Height = ((lastResultRow + 4) * 12) + 3 + 3
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .Left = wshGL_BV.Range("N" & rowSelected).Left + 4
        .Top = wshGL_BV.Range("N" & rowSelected + 1).Top + 4
    End With
        
    'Libérer la mémoire
    Set dynamicShape = Nothing
      
    Call Log_Record("modGL_BV:wshGL_BV_Adjust_The_Shape", "", startTime)
      
End Sub

Sub GL_BV_Show_Dynamic_Shape()

    Dim shp As Shape: Set shp = wshGL_BV.Shapes("JE_Detail_Trans")
    shp.Visible = msoTrue
    
    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub GL_BV_Hide_Dynamic_Shape()

    Dim shp As Shape: Set shp = wshGL_BV.Shapes("JE_Detail_Trans")
    shp.Visible = msoFalse

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Private Function Nz(val As Variant) As Double '2025-05-27 @ 17:55 - v6.C.7 - ChatGPT

    If IsNull(val) Or IsEmpty(val) Then
        Nz = 0
    Else
        Nz = val
    End If
    
End Function

Sub shp_GL_BV_Exit_Click()

    Dim ws As Worksheet
    Set ws = wshGL_BV
    
    Call GL_BV_EffacerZoneTransactionsDetaillees(ws)
    Call GL_BV_EffacerZoneBV(ws)
    Call GL_BV_SupprimerToutesLesFormes_shpRetour(ws)
    DoEvents
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call GL_BV_Back_To_Menu

End Sub

Sub GL_BV_Back_To_Menu()
    
    Call Erase_Non_Required_Shapes
    
    wshGL_BV.Visible = xlSheetHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
End Sub

