Attribute VB_Name = "modGL_BV"
Option Explicit

Public dynamicShape As Shape

Sub shp_GL_BV_Actualiser_Click() '2025-06-03 @ 20:23

    Dim ws As Worksheet
    Set ws = wshGL_BV
    
    Application.ScreenUpdating = True
    Application.EnableEvents = False
    wshGL_BV.Range("C2").Value = "Au " & Format$(ws.Range("J1").Value, wsdADMIN.Range("B1").Value)
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
    If lastUsedRow > 3 Then
        ws.Range("D4:G" & lastUsedRow).Clear
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    Dim resumeGL As Variant
    resumeGL = Get_Summary_By_GL_Account(#7/31/2024#, ActiveSheet.Range("J1").Value)
    
    Call Afficher_BV_Summary(resumeGL)
    
    'Libérer la mémoire
    Set ws = Nothing

End Sub

Function Get_Summary_By_GL_Account(dateMin As Date, dateMax As Date) As Variant '2025-06-03 @ 20:16

    Dim cn As Object, rs As Object
    Dim sql As String, tmpFile As String
    Dim dPlanComptable As Object, dSoldeParGL As Object
    Dim arrPC As Variant, rsKey As String
    Dim i As Long, tDebit As Currency, tCredit As Currency, solde As Currency
    Dim tblData() As Variant, tblFinal() As Variant, key As Variant, soldes As Variant
    Const COL_CODE As Long = 1
    Const COL_DESC As Long = 2
    Const COL_DEBIT As Long = 3
    Const COL_CREDIT As Long = 4

    On Error GoTo ErrHandler

    'Copier les données vers un fichier temporaire (silencieusement)
    tmpFile = CreerCopieTemporaireSolide("GL_Trans")
'    tmpFile = CréerCopieTemporaireSansFlash("GL_Trans")
    If tmpFile = vbNullString Then Exit Function

    sql = "SELECT [NoCompte], SUM([Débit]) AS TotalDébit, SUM([Crédit]) AS TotalCrédit " & _
          "FROM [GL_Trans$] " & _
          "WHERE [Date] >= #" & Format(dateMin, "yyyy-mm-dd") & "# " & _
          "AND [Date] <= #" & Format(dateMax, "yyyy-mm-dd") & "# " & _
          "GROUP BY [NoCompte] ORDER BY [NoCompte]"

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & tmpFile & ";" & _
            "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    Set rs = cn.Execute(sql)
    If rs.EOF Then GoTo CleanUp

    Set dPlanComptable = CreateObject("Scripting.Dictionary")
    arrPC = Fn_Get_Plan_Comptable(2)
    For i = 1 To UBound(arrPC, 1)
        If Not dPlanComptable.Exists(arrPC(i, 1)) Then
            dPlanComptable.Add arrPC(i, 1), arrPC(i, 2)
        End If
    Next i

    Set dSoldeParGL = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        rsKey = rs.Fields("NoCompte").Value
        If Not dPlanComptable.Exists(rsKey) Then
            dPlanComptable.Add rsKey, "Compte inconnu"
        End If
        tDebit = Nz(rs.Fields("TotalDébit"))
        tCredit = Nz(rs.Fields("TotalCrédit"))
        If tDebit <> 0 Or tCredit <> 0 Then
            If Not dSoldeParGL.Exists(rsKey) Then
                dSoldeParGL.Add rsKey, Array(tDebit, tCredit)
            End If
        End If
        rs.MoveNext
    Loop

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
            tblData(i, COL_CODE) = key
            tblData(i, COL_DESC) = dPlanComptable(key)
            If solde >= 0 Then
                tblData(i, COL_DEBIT) = solde
            Else
                tblData(i, COL_CREDIT) = -solde
            End If
            i = i + 1
        End If
    Next key

    If i = 1 Then GoTo CleanUp ' Aucune ligne à afficher
    ReDim tblFinal(1 To i - 1, 1 To 4)
    For i = 1 To UBound(tblFinal, 1)
        tblFinal(i, COL_CODE) = tblData(i, COL_CODE)
        tblFinal(i, COL_DESC) = tblData(i, COL_DESC)
        tblFinal(i, COL_DEBIT) = tblData(i, COL_DEBIT)
        tblFinal(i, COL_CREDIT) = tblData(i, COL_CREDIT)
    Next i

    Get_Summary_By_GL_Account = tblFinal

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.state = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.state = 1 Then cn.Close
    If Len(Dir(tmpFile, vbNormal)) > 0 Then Kill tmpFile
    Exit Function

ErrHandler:
    Resume CleanUp
    
End Function

'Function Get_Summary_By_GL_Account(dateMin As Date, dateMax As Date) As Variant '2025-06-01 @ 13:46
'
'    Dim cn As Object, rs As Object
'    Dim wbTemp As Workbook, wsDest As Worksheet
'    Dim sql As String, tmpFile As String
'    Dim arr(), i As Long, totalDebit As Currency, totalCredit As Currency
'    Const HDR_ROW As Long = 4
'
'    On Error GoTo ErrHandler
'
'    'Copie temporaire de la feuille GL_Trans
'    tmpFile = CréerCopieTemporaireSansFlash("GL_Trans")
'    If tmpFile = "" Then Exit Function
'
'    ' Requête SQL pour résumer les débits et crédits par compte
'    sql = "SELECT [NoCompte], " & _
'          "SUM([Débit]) AS TotalDébit, SUM([Crédit]) AS TotalCrédit " & _
'          "FROM [GL_Trans$] " & _
'          "WHERE [Date] >= #" & Format(dateMin, "yyyy-mm-dd") & "# " & _
'          "AND [Date] <= #" & Format(dateMax, "yyyy-mm-dd") & "# " & _
'          "GROUP BY [NoCompte] " & _
'          "ORDER BY [NoCompte]"
'
'    ' Connexion ADO
'    Set cn = CreateObject("ADODB.Connection")
'    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & tmpFile & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
'
'    Set rs = cn.Execute(sql)
'    If rs.EOF Then
'        rs.Close: cn.Close
'        If Len(Dir(tmpFile, vbNormal)) > 0 Then Kill tmpFile
'        Exit Function
'    End If
'
'    'Construction d'un dictionnaire pour le Plan Comptable (tous les comptes)
'    Dim dPlanComptable As Object: Set dPlanComptable = CreateObject("Scripting.Dictionary")
'    Dim arrPC As Variant
'    arrPC = Fn_Get_Plan_Comptable(2) 'Retourne 2 colonnes (Code & Description)
'    For i = 1 To UBound(arrPC, 1)
'        If Not dPlanComptable.Exists(arrPC(i, 1)) Then
'            dPlanComptable.Add arrPC(i, 1), arrPC(i, 2)
'        End If
'    Next i
'
'    'Fusion du recordSet et du dictionnaire du Plan Comptable
'    Dim solde As Currency, tDebit As Currency, tCredit As Currency
'    Dim dSoldeParGL As Object: Set dSoldeParGL = CreateObject("Scripting.Dictionary")
'
'    Dim rsKey As String
'    Do While Not rs.EOF
'        rsKey = rs.Fields("NoCompte").Value
'        'S'assurer que tous les comptes du résultat SQL sont présents dans le plan comptable
'        If Not dPlanComptable.Exists(rsKey) Then
'            dPlanComptable.Add rsKey, "Compte inconnu"
'        End If
'        tDebit = Nz(rs.Fields("TotalDébit"))
'        tCredit = Nz(rs.Fields("TotalCrédit"))
'        If tDebit <> 0 Or tCredit <> 0 Then
'            If Not dSoldeParGL.Exists(rsKey) Then
'                dSoldeParGL.Add rsKey, Array(tDebit, tCredit)
'            End If
'        End If
'        rs.MoveNext
'    Loop
'
'    'Création d'un tableau pour emmagasiner les informations
'    Dim soldes As Variant
'    Dim tblData() As Variant
'    ReDim tblData(1 To dPlanComptable.count, 1 To 4)
'    i = 1
'    Dim key As Variant
'    Const COL_CODE = 1, COL_DESC = 2, COL_DEBIT = 3, COL_CREDIT = 4
'    For Each key In dPlanComptable.keys
'        soldes = Array(0, 0)
'        If dSoldeParGL.Exists(key) Then
'            soldes = dSoldeParGL(key)
'            solde = soldes(0) - soldes(1)
'        Else
'            solde = 0
'        End If
'
'        If soldes(0) <> 0 Or soldes(1) <> 0 Then
'            tblData(i, COL_CODE) = key
'            tblData(i, COL_DESC) = dPlanComptable(key)
'            If solde >= 0 Then
'                tblData(i, COL_DEBIT) = solde
'            Else
'                tblData(i, COL_CREDIT) = -solde
'            End If
'            i = i + 1
'        End If
'    Next key
'
'    'Enlève les lignes qui n'ont pas MINIMALEMENT un débit ou un crédit
'    Dim tblFinal() As Variant
'    Dim j As Long
'    If i > 1 Then
'        ReDim tblFinal(1 To i - 1, 1 To 4)
'        For j = 1 To i - 1
'            tblFinal(j, COL_CODE) = tblData(j, COL_CODE)
'            tblFinal(j, COL_DESC) = tblData(j, COL_DESC)
'            tblFinal(j, COL_DEBIT) = tblData(j, COL_DEBIT)
'            tblFinal(j, COL_CREDIT) = tblData(j, COL_CREDIT)
'        Next j
'        'Utilisez tblFinal à la place de tblData
'        tblData = tblFinal
'    Else
'        Erase tblData
'        Exit Function
'    End If
'    Erase tblFinal
'
'    'Écrire résultats + calculer totaux
'    Dim ligne As Long
'    ligne = 4
'    Dim globalDebit As Currency, globalCredit As Currency
'    Application.EnableEvents = False
'    For i = 1 To UBound(tblData, 1)
'        wshGL_BV.Cells(ligne, 4).Resize(1, 4).Value = Array(tblData(i, COL_CODE), tblData(i, COL_DESC), tblData(i, COL_DEBIT), tblData(i, COL_CREDIT))
'        globalDebit = globalDebit + tblData(i, COL_DEBIT)
'        globalCredit = globalCredit + tblData(i, COL_CREDIT)
'        ligne = ligne + 1
'    Next i
'
'   'Afficher les totaux
'    ligne = ligne + 1
'    With wshGL_BV.Cells(ligne, 4)
'        .Value = "TOTALS"
'        .Font.Bold = True
'    End With
'    wshGL_BV.Cells(ligne, 6).Value = globalDebit
'    wshGL_BV.Cells(ligne, 7).Value = globalCredit
'
'    With wshGL_BV.Range("F" & ligne & ":" & "G" & ligne)
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
'        .Font.Bold = True
'        .NumberFormat = "#,##0.00 $"
'    End With
'
'    wshGL_BV.Range("D4:D" & ligne).HorizontalAlignment = xlCenter
'
'    'Vérification intégrité (DT ?= CT)
'    If Round(globalDebit, 2) <> Round(globalCredit, 2) Then
'        MsgBox "Il y a une différence entre le total des débits et le total des crédits : " & Format(globalDebit - globalCredit, "0.00"), vbExclamation
'    End If
'
'    Exit Function
'
'ErrHandler:
'    On Error Resume Next
'    If Not rs Is Nothing Then If rs.state = 1 Then rs.Close
'    If Not cn Is Nothing Then If cn.state = 1 Then cn.Close
'    If Len(Dir(tmpFile, vbNormal)) > 0 Then Kill tmpFile
'
'End Function

'Function Get_Summary_By_GL_Account(dateMin As Date, dateMax As Date) As ADODB.Recordset '2025-05-27 @ 17:51 - v6.C.7 - ChatPGT
'
'    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modGL_BV:Get_Summary_By_GL_Account", "", 0)
'
'    Dim cn As ADODB.Connection
'    Dim rs As ADODB.Recordset
'    Dim strSQL As String
'
'    'Fichier actif
'    Dim sWBPath As String
'    sWBPath = ThisWorkbook.FullName
'
'    'Requête SQL
'    strSQL = "SELECT [NoCompte], " & _
'           "SUM([Débit]) AS TotalDébit, " & _
'           "SUM([Crédit]) AS TotalCrédit " & _
'           "FROM [GL_Trans$] " & _
'           "WHERE [Date] >= #" & Format(dateMin, "yyyy-mm-dd") & "# " & _
'           "AND [Date] <= #" & Format(dateMax, "yyyy-mm-dd") & "# " & _
'           "GROUP BY [NoCompte] " & _
'           "ORDER BY [NoCompte];"
'    Debug.Print "Calcul de la BV" & vbNewLine & "Get_Summary_By_GL_Account - strSQL = " & strSQL
'
'    'Connexion ADO
'    Set cn = New ADODB.Connection
'    With cn
'        .Provider = "Microsoft.ACE.OLEDB.12.0"
'        .ConnectionString = "Data Source=" & sWBPath & ";" & _
'                            "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
'        .Open
'    End With
'
'    'Recordset
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, cn, adOpenStatic, adLockReadOnly
'
'    'Retour
'    Set Get_Summary_By_GL_Account = rs
'
'    'Libérer
'    Set cn = Nothing
'
'    Call EnregistrerLogApplication("modGL_BV:Get_Summary_By_GL_Account", "", startTime)
'
'End Function

Sub Afficher_BV_Summary(tblData As Variant, Optional ligneDépart As Long = 4) '2025-06-03 @ 20:18

    Dim i As Long, ligne As Long
    Dim globalDebit As Currency, globalCredit As Currency
    Const COL_CODE = 1, COL_DESC = 2, COL_DEBIT = 3, COL_CREDIT = 4

    If IsEmpty(tblData) Then Exit Sub

    ligne = ligneDépart
    Application.EnableEvents = False

    ' Écriture des lignes
    For i = 1 To UBound(tblData, 1)
        wshGL_BV.Cells(ligne, 4).Resize(1, 4).Value = Array( _
            tblData(i, COL_CODE), _
            tblData(i, COL_DESC), _
            tblData(i, COL_DEBIT), _
            tblData(i, COL_CREDIT))
        globalDebit = globalDebit + tblData(i, COL_DEBIT)
        globalCredit = globalCredit + tblData(i, COL_CREDIT)
        ligne = ligne + 1
    Next i

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

    If Round(globalDebit, 2) <> Round(globalCredit, 2) Then
        MsgBox "Il y a une différence entre le total des débits et des crédits : " & _
               Format(globalDebit - globalCredit, "0.00"), vbExclamation
    End If

    Application.EnableEvents = True
    
End Sub

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
    Dim Debit As Currency, Credit As Currency, solde As Currency, soldeInitial As Currency

    'Feuilles
    Set wsTrans = wsdGL_Trans
    Set wsResult = wshGL_BV

    'Compte & description (passés en paramètre)
    If compte = vbNullString Then
        MsgBox "Aucun compte sélectionné.", vbExclamation
        Exit Sub
    End If

    'Nettoyer la zone existante (M5 vers le bas) & ajuster l'entête
    Call GL_BV_EffacerZoneTransactionsDetaillees(wsResult)
    
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
    
    'Connexion ADO
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & ThisWorkbook.FullName & ";" & _
                            "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";"
        .Open
    End With

    'Calcul du solde initial avant DateMin
    dateMin = wsResult.Range("B8").Value
'    wsResult.Range("L2").Value = "Du " & Format$(dateMin, wsdADMIN.Range("B1").Value) & " au " & Format$(dateMax, wsdADMIN.Range("B1").Value)
    Set rsInit = CreateObject("ADODB.Recordset")
    
    strSQL = "SELECT SUM(Débit) AS TotalDebit, SUM(Crédit) AS TotalCredit FROM [GL_Trans$] " & _
             "WHERE NoCompte = '" & compte & "' AND Date < #" & Format(dateMin, "mm/dd/yyyy") & "#"
    Debug.Print "GL_BV_Display_Trans_For_Selected_Account - strSQL1 = " & strSQL
    
    rsInit.Open strSQL, cn, 1, 1
    If Not rsInit.EOF Then
        soldeInitial = Nz(rsInit.Fields("TotalDebit").Value) - Nz(rsInit.Fields("TotalCredit").Value)
    End If
    rsInit.Close: Set rsInit = Nothing
    
    'Requête SQL complète (toutes les dates) pour le compte
    strSQL = "SELECT Date, NoEntrée, Description, Source, Débit, Crédit, AutreRemarque FROM [GL_Trans$] " & _
             "WHERE NoCompte = '" & Replace(compte, "'", "''") & "'" & _
             "AND Date >= #" & Format(dateMin, "yyyy-mm-dd") & "# " & _
             "AND Date <= #" & Format(dateMax, "yyyy-mm-dd") & "# " & _
             "ORDER BY Date, NoEntrée"
    Debug.Print "GL_BV_Display_Trans_For_Selected_Account - strSQL2 = " & strSQL
    
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
        wsResult.Range("P4").Value = "Solde d'ouverture au " & Format(dateMin, wsdADMIN.Range("B1"))
        wsResult.Range("S4").Value = soldeInitial
        With wsResult.Range("P4:S4")
            .Font.Name = "Aptos Narrow"
            .Font.size = 9
            .Font.Bold = True
        End With
        solde = soldeInitial
        ligne = 1 'Commencer les écritures de transactions à la 1ère ligne du tableau
        Application.EnableEvents = True

        Do While Not rs.EOF
            Debit = Nz(rs.Fields("Débit").Value)
            Credit = Nz(rs.Fields("Crédit").Value)
            solde = solde + Debit - Credit

            tableau(ligne, 1) = rs.Fields("Date").Value
            tableau(ligne, 2) = rs.Fields("NoEntrée").Value
            tableau(ligne, 3) = rs.Fields("Description").Value
            tableau(ligne, 4) = rs.Fields("Source").Value
            tableau(ligne, 5) = IIf(Debit > 0, Debit, vbNullString)
            tableau(ligne, 6) = IIf(Credit > 0, Credit, vbNullString)
            tableau(ligne, 7) = solde
            tableau(ligne, 8) = rs.Fields("AutreRemarque")

            ligne = ligne + 1
            rs.MoveNext
        Loop

        'Écriture de tableau dans la plage, en commençant à M5 - @TODO - 2025-07-11 @ 03:14
        Application.EnableEvents = False
        wsResult.Range("M5").Resize(nbLignes, 8).Value = tableau
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
        wsResult.Range("L4").Value = vbNullString
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
    lastUsedRow = ws.Cells(ws.Rows.count, "M").End(xlUp).Row
    
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
    visibleRows = ActiveWindow.VisibleRange.Rows.count
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

Sub shp_GL_BV_Impression_BV_Click()

    Call GL_BV_Setup_And_Print

End Sub

Sub GL_BV_Setup_And_Print()
    
    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modGL_BV:GL_BV_Setup_And_Print", vbNullString, 0)
    
    Dim lastRow As Long
    lastRow = wshGL_BV.Cells(wshGL_BV.Rows.count, "D").End(xlUp).Row + 2
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
    
    Call EnregistrerLogApplication("modGL_BV:GL_BV_Setup_And_Print", vbNullString, startTime)

End Sub

Sub shp_GL_BV_Setup_And_Print_Trans_Click()

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modGL_BV:shp_GL_BV_Setup_And_Print_Trans_Click", vbNullString, 0)
    
    Call GL_BV_Setup_And_Print_Trans

    Call EnregistrerLogApplication("modGL_BV:shp_GL_BV_Setup_And_Print_Trans_Click", vbNullString, startTime)

End Sub

Sub GL_BV_Setup_And_Print_Trans()
    
    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modGL_BV:GL_BV_Setup_And_Print_Trans", vbNullString, 0)
    
    Dim lastRow As Long
    lastRow = wshGL_BV.Cells(wshGL_BV.Rows.count, "M").End(xlUp).Row
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
    
    Call EnregistrerLogApplication("modGL_BV:GL_BV_Setup_And_Print_Trans", vbNullString, startTime)

End Sub

Sub GL_BV_SetUp_And_Print_Document(myPrintRange As Range, pagesTall As Long)
    
    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modGL_BV:GL_BV_SetUp_And_Print_Document", vbNullString, 0)
    
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
        Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1 is completed", -1)
        
        'Page Header & Footer
'        .LeftHeader = ""
        .CenterHeader = "&""Aptos Narrow,Gras""&18 " & wsdADMIN.Range("NomEntreprise").Value
        Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1.A is completed", -1)
        
'        .RightHeader = ""
        .LeftFooter = "&9&D - &T"
'        .CenterFooter = ""
        .RightFooter = "&9Page &P de &N"
        Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 1.B is completed", -1)
        
        'Page Margins
        Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 is starting", -1)
        .LeftMargin = Application.InchesToPoints(0.16)
        .RightMargin = Application.InchesToPoints(0.16)
         Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Left & Right) margins", -1)
         
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
         Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Top & Bottom) margins", -1)
         
        .CenterHorizontally = True
        .CenterVertically = False
         Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Center Horizontal & Vertical)", -1)
         
        'Header and Footer margins
        .HeaderMargin = Application.InchesToPoints(0.16)
        .FooterMargin = Application.InchesToPoints(0.16)
        Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Block 2 (Header & Footer) margins", -1)
        
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    Call EnregistrerLogApplication("   modGL_BV:GL_BV_SetUp_And_Print_Document - Speed Measure", -1)
    
    wshGL_BV.PrintPreview '2024-08-15 @ 14:53
 
    Call EnregistrerLogApplication("modGL_BV:GL_BV_SetUp_And_Print_Document", vbNullString, startTime)
 
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

Sub GL_BV_Display_JE_Trans_With_Shape()

    Call GL_BV_Create_Dynamic_Shape
    Call GL_BV_Adjust_The_Shape
    Call GL_BV_Show_Dynamic_Shape
    
End Sub

Sub GL_BV_Create_Dynamic_Shape()

    'Check if the shape has already been created
    If dynamicShape Is Nothing Then
        'Create the text box shape
        wshGL_BV.Unprotect
        Set dynamicShape = wshGL_BV.Shapes.AddShape(msoShapeRoundedRectangle, 2000, 100, 600, 100)
    End If

End Sub

Sub GL_BV_Adjust_The_Shape()

    Dim startTime As Double: startTime = Timer: Call EnregistrerLogApplication("modGL_BV:GL_BV_Adjust_The_Shape", vbNullString, 0)
    
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
            texteOneLine = Fn_Pad_A_String(.Range("AG" & i).Value, " ", 5, "R") & _
                            " - " & Fn_Pad_A_String(.Range("AH" & i).Value, " ", 35, "R") & _
                            "  " & Fn_Pad_A_String(Format$(.Range("AI" & i).Value, "#,##0.00 $"), " ", 14, "L") & _
                            "  " & Fn_Pad_A_String(Format$(.Range("AJ" & i).Value, "#,##0.00 $"), " ", 14, "L")
            If Trim$(.Range("AF" & i).Value) = Trim$(wshGL_BV.Range("B6").Value) Then
                texteOneLine = " * " & texteOneLine
            Else
                texteOneLine = "   " & texteOneLine
            End If
            texteOneLine = Fn_Pad_A_String(texteOneLine, " ", 79, "R")
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
      
    Call EnregistrerLogApplication("modGL_BV:GL_BV_Adjust_The_Shape", vbNullString, startTime)
      
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


