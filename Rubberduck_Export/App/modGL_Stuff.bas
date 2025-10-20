Attribute VB_Name = "modGL_Stuff"
Option Explicit

'Structure pour une écriture comptable (données communes)
Public Type tGL_Entry '2025-06-08 @ 06:59
    DateTrans As Date
    source As String
    noCompte As String
    autreRemarque As String
End Type

'Structure pour une écriture comptable (données spécifiques à chaque ligne)
Public Type tGL_EntryLine '2025-06-08 @ 07:02
    noCompte As String
    description As String
    montant As Double
End Type

Public Sub ObtenirSoldeCompteEntreDebutEtFin(glNo As String, dateDeb As Date, dateFin As Date, ByRef rResult As Range)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:ObtenirSoldeCompteEntreDebutEtFin", glNo & " - De " & dateDeb & " à " & dateFin, 0)

    'Les données à AF proviennent de GL_Trans
    Dim ws As Worksheet: Set ws = wsdGL_Trans
    
    'wsdGL_Trans_AF#1

    'Effacer les données de la dernière utilisation
    ws.Range("M6:M10").ClearContents
    ws.Range("M6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_GL_Trans[#All]")
    ws.Range("M7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("L2:N3")
    ws.Range("L3").Value = glNo
    ws.Range("M3").Value = ">=" & CLng(dateDeb)
    ws.Range("N3").Value = "<=" & CLng(dateFin)
    ws.Range("M8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("P1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("P1:Y1")
    ws.Range("M9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "P").End(xlUp).Row
    ws.Range("M10").Value = lastUsedRow - 1 & " lignes"
    
    If lastUsedRow > 2 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsdGL_Trans.Range("T2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par numéro de compte
            .SortFields.Add key:=wsdGL_Trans.Range("Q2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par date
            .SortFields.Add key:=wsdGL_Trans.Range("P2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Tri par numéro d'écriture
            .SetRange wsdGL_Trans.Range("P2:Y" & lastUsedRow) 'Set Range
            .Apply 'Apply Sort
        End With
    End If

    'Retourne le Range des résultats
    Set rResult = wsdGL_Trans.Range("P1:Y" & lastUsedRow)
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:ObtenirSoldeCompteEntreDebutEtFin", vbNullString, startTime)

End Sub

Sub ObtenirEcritureAvecAF(noEJ As Long) '2024-11-17 @ 12:08

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_EJ:ObtenirEcritureAvecAF", vbNullString, 0)

    Dim ws As Worksheet: Set ws = wsdGL_Trans
    
    'wsdGL_Trans_AF#2

    'Effacer les données de la dernière utilisation
    ws.Range("AA6:AA10").ClearContents
    ws.Range("AA6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_GL_Trans[#All]")
    ws.Range("AA7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("AA2:AA3")
    ws.Range("AA3").Value = noEJ
    ws.Range("AA8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("AC1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("AC1:AL1")
    ws.Range("AA9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AC").End(xlUp).Row
    ws.Range("AA10").Value = lastUsedRow - 1 & " lignes"

    'On tri les résultats par noGL / par date?
    If lastUsedRow > 2 Then
        With ws.Sort 'Sort - NoEntrée, Débit(D) et Crédit (D)
        .SortFields.Clear
            'First sort On NoEntrée
            .SortFields.Add key:=ws.Range("AC2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Débit(D)
            .SortFields.Add key:=ws.Range("AI2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlDescending, _
                DataOption:=xlSortNormal
            'Third, sort On Crédit(D)
            .SortFields.Add key:=ws.Range("AJ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlDescending, _
                DataOption:=xlSortNormal
            .SetRange wsdGL_Trans.Range("AC2:AL" & lastUsedRow)
            .Apply 'Apply Sort
         End With
    End If
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_EJ:ObtenirEcritureAvecAF", vbNullString, startTime)

End Sub

Sub AjouterFormeRetourEnHaut()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:AjouterFormeRetourEnHaut", vbNullString, 0)
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Dim btn As Shape
    Dim leftPosition As Double
    Dim topPosition As Double

    'Trouver la dernière ligne de la plage L4:T*
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "M").End(xlUp).Row

    If lastRow >= 5 Then
        'Calculer les positions (Left & Top) du bouton
        leftPosition = ws.Range("T" & lastRow).Left
        topPosition = ws.Range("S" & lastRow).Top + (2 * ws.Range("S" & lastRow).Height)
    
        ' Ajouter une Shape
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, Left:=leftPosition, Top:=topPosition, _
                                                        Width:=90, Height:=30)
        With btn
            .Name = "shpRetour"
            .TextFrame2.TextRange.text = "Retour"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.TextRange.Font.size = 14
            .TextFrame2.TextRange.Font.Bold = True
            .TextFrame2.HorizontalAnchor = msoAnchorCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .Fill.ForeColor.RGB = RGB(166, 166, 166)
            .OnAction = "EffacerZoneTransDetailleesEtForme"
        End With
    End If
    
    'Libérer la mémoire
    Set btn = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:AjouterFormeRetourEnHaut", vbNullString, startTime)

End Sub

Sub EffacerZoneTransDetailleesEtForme()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:EffacerZoneTransDetailleesEtForme", vbNullString, 0)
    
    'Effacer la plage
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Application.EnableEvents = False
    ws.Range("L1:T" & ws.Cells(ws.Rows.count, "M").End(xlUp).Row).Offset(3, 0).Clear
    Application.EnableEvents = True

    'Supprimer les formes shpRetour
    Call SupprimerToutesFormesRetour(ws)

    Call EffacerFormeDynamique
    
    'Ramener le focus à C4
    Application.EnableEvents = False
    ws.Range("D4").Select
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:EffacerZoneTransDetailleesEtForme", vbNullString, startTime)

End Sub

Sub EffacerZoneBV(w As Worksheet)

    Application.EnableEvents = False
    Dim lastUsedRow As Long
    lastUsedRow = w.Cells(w.Rows.count, "D").End(xlUp).Row
    If lastUsedRow >= 4 Then
        w.Range("D4:G" & lastUsedRow).Clear
    End If
    Application.EnableEvents = True

End Sub

Sub SupprimerToutesFormesRetour(w As Worksheet)

    Dim shp As Shape

    For Each shp In w.Shapes
        If shp.Name = "shpRetour" Then
            shp.Delete
        End If
    Next shp
    
End Sub
    
'@Description "Retourne un dictionnaire avec sommaire par noCompte & Solde à une date donnée"
Function Fn_SoldesParCompteAvecADO(noCompteGLMin As String, noCompteGLMax As String, dateLimite As Date, _
                                       inclureEcrCloture As Boolean) As Dictionary '2025-08-02 @ 10:04

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_BV:Fn_SoldesParCompteAvecADO", vbNullString, 0)
    
    Dim strSQL As String
    Dim dictSoldes As Object: Set dictSoldes = CreateObject("Scripting.Dictionary")
    Dim cle As String
    Dim montant As Currency
    
    'Si un seul compte est spécifié, le MAX = MIN
    If noCompteGLMax = vbNullString Then
        noCompteGLMax = noCompteGLMin
    End If
    
    'Fichier fermé est GCF_BD_MASTER.xlsx et la feuille est 'GL_Trans'
    Dim cheminFichier As String
    cheminFichier = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & _
                    Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    Dim nomFeuille As String
    nomFeuille = "GL_Trans"
    
    'Connexion ADO à un classeur fermé
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & cheminFichier & ";" & _
              "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'Requête : somme des montants pour chaque compte (>= 4000), jusqu’à la date de clôture incluse
    strSQL = "SELECT NoCompte, SUM(IIF(Débit IS NULL, 0, Débit)) - SUM(IIF(Crédit IS NULL, 0, Crédit)) AS Solde " & _
             "FROM [" & nomFeuille & "$] " & _
             "WHERE NoCompte >= '" & noCompteGLMin & "' AND NoCompte <= '" & noCompteGLMax & _
             "' AND Date <= #" & Format(dateLimite, "yyyy-mm-dd") & "#"
    If Not inclureEcrCloture Then
        strSQL = strSQL & " AND (Source IS NULL OR NOT (Date = #" & Format(dateLimite, "yyyy-mm-dd") & "# AND Source = 'Clôture annuelle'))"
    End If
    
    strSQL = strSQL & " GROUP BY NoCompte"

    recSet.Open strSQL, conn, 1, 1

    Do While Not recSet.EOF
        cle = CStr(recSet.Fields("NoCompte").Value)
'        Debug.Print "Construction du dictionary (dictSoldes): " & cle & " = " & Format$(recSet.Fields("Solde").Value, "#,##0.00")
        montant = Nz(recSet.Fields("Solde").Value)
        If Not dictSoldes.Exists(cle) Then
            dictSoldes.Add cle, montant
        Else
            dictSoldes(cle) = dictSoldes(cle) + montant
        End If
        recSet.MoveNext
    Loop

    recSet.Close
    conn.Close
    
    Set Fn_SoldesParCompteAvecADO = dictSoldes
    
    GoTo Exit_Function

ErrHandler:
    MsgBox "Erreur dans Fn_SoldesParCompteAvecADO : " & Err.description, vbCritical
    On Error Resume Next
    If Not recSet Is Nothing Then If recSet.state = 1 Then recSet.Close
    If Not conn Is Nothing Then If conn.state = 1 Then conn.Close
    Set Fn_SoldesParCompteAvecADO = Nothing
    
Exit_Function:
    Call modDev_Utils.EnregistrerLogApplication("modGL_BV:Fn_SoldesParCompteAvecADO", vbNullString, startTime)

End Function

Public Function Nz(val As Variant) As Currency '2025-07-17 @ 09:57

    If IsNull(val) Or IsEmpty(val) Then
        Nz = 0
    Else
        Nz = val
    End If
    
End Function

Function Fn_DateFinExercice(dateSaisie As Date) As Date '2025-07-20 @ 08:49

    Dim anneeExercice As Integer
    
    Dim moisFinExercice As Integer
    moisFinExercice = wsdADMIN.Range("MoisFinAnnéeFinancière").Value

    'Si le mois de la date saisie est supérieur au mois de fin, alors la fin d'exercice est l'année suivante
    If month(dateSaisie) > moisFinExercice Then
        anneeExercice = year(dateSaisie) + 1
    Else
        anneeExercice = year(dateSaisie)
    End If

    ' Dernier jour du mois de fin d’exercice
    Fn_DateFinExercice = DateSerial(anneeExercice, moisFinExercice + 1, 0)
    
End Function

Public Sub AjouterEcritureGLADOPlusLocale(entry As clsGL_Entry, Optional afficheMessage As Boolean = True) '2025-06-08 @ 09:37

    '=== BLOC 1 : Écriture dans GCF_BD_MASTER.xslx en utilisant ADO ===
    Dim cheminMASTER As String
    Dim nextNoEntree As Long
    Dim ts As String
    Dim i As Long
    Dim l As clsGL_EntryLine
    Dim strSQL As String

    On Error GoTo CleanUpADO

    'Chemin du classeur MASTER.xlsx
    cheminMASTER = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    
    'Ouvre connexion ADO
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminMASTER & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Détermine le prochain numéro d'écriture
    Dim recSet As Object: Set recSet = conn.Execute("SELECT MAX([NoEntrée]) AS MaxNo FROM [GL_Trans$]")
    If Not recSet.EOF And Not IsNull(recSet!MaxNo) Then
        nextNoEntree = recSet!MaxNo + 1
    Else
        nextNoEntree = 1
    End If
    entry.NoEcriture = nextNoEntree
    recSet.Close
    Set recSet = Nothing

    'Timestamp unique pour l'écriture
    ts = Format(Now, "yyyy-mm-dd hh:mm:ss")

    'Ajoute chaque ligne d'écriture dans le classeur MASTER.xlsx
    For i = 1 To entry.lignes.count
        Set l = entry.lignes(i)
        strSQL = "INSERT INTO [GL_Trans$] " & _
              "([NoEntrée],[Date],[Description],[Source],[NoCompte],[Compte],[Débit],[Crédit],[AutreRemarque],[TimeStamp]) " & _
              "VALUES (" & _
              entry.NoEcriture & "," & _
              "'" & Format(entry.DateEcriture, "yyyy-mm-dd") & "'," & _
              "'" & Replace(entry.description, "'", "''") & "'," & _
              "'" & Replace(entry.source, "'", "''") & "'," & _
              "'" & l.noCompte & "'," & _
              "'" & Replace(l.description, "'", "''") & "'," & _
              IIf(l.montant >= 0, Replace(l.montant, ",", "."), "NULL") & "," & _
              IIf(l.montant < 0, Replace(-l.montant, ",", "."), "NULL") & "," & _
              "'" & Replace(l.autreRemarque, "'", "''") & "'," & _
              "'" & ts & "'" & _
              ")"
        conn.Execute strSQL
    Next i

    conn.Close: Set conn = Nothing

    '=== BLOC 2 - Écriture dans feuille locale (GL_Trans)
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim wsLocal As Worksheet
    Dim lastRow As Long

    'Mémoriser l’état initial d’Excel
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wsLocal = ThisWorkbook.Sheets("GL_Trans")
    lastRow = wsLocal.Cells(wsLocal.Rows.count, 1).End(xlUp).Row

    For i = 1 To entry.lignes.count
        Set l = entry.lignes(i)
        With wsLocal
            .Cells(lastRow + i, 1).Value = entry.NoEcriture
            .Cells(lastRow + i, 2).Value = entry.DateEcriture
            .Cells(lastRow + i, 3).Value = entry.description
            .Cells(lastRow + i, 4).Value = entry.source
            .Cells(lastRow + i, 5).Value = l.noCompte
            .Cells(lastRow + i, 6).Value = l.description
            If l.montant >= 0 Then
                .Cells(lastRow + i, 7).Value = l.montant
                .Cells(lastRow + i, 8).Value = vbNullString
            Else
                .Cells(lastRow + i, 7).Value = vbNullString
                .Cells(lastRow + i, 8).Value = -l.montant
            End If
            .Cells(lastRow + i, 9).Value = l.autreRemarque
            .Cells(lastRow + i, 10).Value = ts
        End With
    Next i

    If afficheMessage Then
        MsgBox "L'écriture comptable a été complétée avec succès", vbInformation, "Écriture au Grand Livre"
    End If

CleanUpADO:
    On Error Resume Next
    If Not recSet Is Nothing Then If recSet.state = 1 Then recSet.Close
    Set recSet = Nothing
    If Not conn Is Nothing Then If conn.state = 1 Then conn.Close
    Set conn = Nothing
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l’écriture au G/L : " & Err.description, vbCritical, "AjouterEcritureGLADOPlusLocale"
    End If
    On Error GoTo 0
    
End Sub

Function Fn_Tableau24MoisSommeTransGL(dateLimite As Date, inclureEcrCloture As Boolean) As Variant '2025-08-05 @ 05:58

    Dim collComptes As Collection
    Dim tableau24Mois() As Variant
    Dim fichier As String
    Dim strSQL As String
    Dim dateDebutOperations As Date
    Dim compteTrouve As Boolean
    Dim periode As String
    
    periode = Fn_Construire24PeriodesGL(dateLimite)
    
    'Chemin du classeur MASTER.xlsx
    fichier = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    
    dateDebutOperations = Format$(#7/31/2024#, "yyyy-mm-dd")
    
    'Connexion ADO
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fichier & ";" & _
              "Extended Properties='Excel 12.0 Xml;HDR=YES'"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")
    
    'Requête SQL
    strSQL = "SELECT [NoCompte], year([Date]) as Annee, month([Date]) as MoisNum, " & _
             "SUM(IIF([Débit] IS NULL, 0, [Débit]) - IIF([Crédit] IS NULL, 0, [Crédit])) AS transTotal " & _
             "FROM [GL_Trans$] " & _
             "WHERE [Date] >= #" & Format$(dateDebutOperations, "yyyy-mm-dd") & "# " & _
             "AND [Date] <= #" & Format$(dateLimite, "yyyy-mm-dd") & "#"
    If Not inclureEcrCloture Then
        strSQL = strSQL & " AND ([Source] IS NULL OR NOT ([Date] = #" & Format(dateLimite, "yyyy-mm-dd") & "# AND [Source] = 'Clôture annuelle'))"
    End If
    
    strSQL = strSQL & "GROUP BY [NoCompte], Year([Date]), Month([Date]) " & _
                      "ORDER BY [NoCompte], Year([Date]), Month([Date]) "
    
    Debug.Print strSQL
    
    recSet.CursorLocation = adUseClient
    recSet.CursorType = adOpenKeyset
    
    recSet.Open strSQL, conn
    
    Debug.Print "Total lignes renvoyées dans le recordSet:", recSet.RecordCount
    
    'Liste unique des collComptes
    Set collComptes = New Collection
    recSet.MoveFirst
    
    On Error Resume Next
    Do While Not recSet.EOF
        collComptes.Add recSet("NoCompte").Value, CStr(recSet("NoCompte").Value)
        recSet.MoveNext
    Loop
    On Error GoTo 0
    
    'tableau24Mois [Compte, Ouverture & 24 mois]
    ReDim tableau24Mois(0 To collComptes.count - 1, 0 To 25)
    
    'Remplir colonne 0 avec les collComptes
    Dim i As Long
    For i = 0 To collComptes.count - 1
        tableau24Mois(i, 0) = collComptes(i + 1) 'Collection indexée à partir de 1
    Next

    'Remplissage des mois
    Dim annee As Long
    Dim mois As Long
    Dim j As Long
    
    recSet.MoveFirst
    
    Do While Not recSet.EOF
        compteTrouve = False
        For i = 0 To collComptes.count - 1
            If recSet("NoCompte").Value = tableau24Mois(i, 0) Then
                annee = recSet("Annee").Value
                mois = recSet("MoisNum").Value
                j = InStr(periode, Format$(annee, "0000") & "-" & Format$(mois, "00"))
                If Not j < 1 Then
                    j = ((j + 7) / 8) + 1
                Else
                    j = 1
                End If
                Debug.Print "C", i, tableau24Mois(i, 0), annee, mois, j, CCur(recSet("transTotal"))
                
                tableau24Mois(i, j) = tableau24Mois(i, j) + CCur(recSet("transTotal"))
                compteTrouve = True
                Exit For
            End If
        Next
        If Not compteTrouve Then Debug.Print "Compte non trouvé :"; recSet("NoCompte")
        recSet.MoveNext
    Loop

    'Libérer la mémoire
    recSet.Close: conn.Close
    Set recSet = Nothing: Set conn = Nothing
    
    'Résultat
    Fn_Tableau24MoisSommeTransGL = tableau24Mois
    
End Function

Function Fn_Construire24PeriodesGL(dateLimite As Date) As String

    Dim periodes As String
    Dim tmpAnnee As Long
    Dim tmpMois As Long
    tmpAnnee = year(dateLimite)
    tmpMois = month(dateLimite)
    
    Dim i As Long
    For i = 1 To 24
        periodes = Format$(tmpAnnee, "0000") & "-" & Format$(tmpMois, "00") & " " & periodes
        tmpMois = tmpMois - 1
        If tmpMois < 1 Then
            tmpMois = 12
            tmpAnnee = tmpAnnee - 1
        End If
    Next i
    
    Fn_Construire24PeriodesGL = periodes

End Function

Sub zz_TestTableau24MoisGLDansExcel() '2025-08-05 @ 05:58

    Dim tableau() As Variant
    Dim i As Long, j As Long
    
    Dim dateCutOff As Date
    dateCutOff = #7/31/2025#
    
    Dim periodes As String
    periodes = Fn_Construire24PeriodesGL(dateCutOff)
    
    'Détermine le mois de l'année financière en fonction de la date limite
    Dim dernierMoisAnneeFinanciere As Long
    dernierMoisAnneeFinanciere = wsdADMIN.Range("MoisFinAnnéeFinancière")
    Dim moisAnneeFinanciere As Long
    moisAnneeFinanciere = month(dateCutOff)
    If moisAnneeFinanciere > dernierMoisAnneeFinanciere Then
        moisAnneeFinanciere = moisAnneeFinanciere - dernierMoisAnneeFinanciere
    Else
        moisAnneeFinanciere = moisAnneeFinanciere + 12 - dernierMoisAnneeFinanciere
    End If
    
    Debug.Print "Pour la date '" & Format$(dateCutOff, "yyyy-mm-dd") & "' le mois de l'année financière est " & moisAnneeFinanciere
    
    'Feuille de travail
    Dim feuilleNom As String
    feuilleNom = "X_GLTableau24Mois"
    Call modDev_Utils.EffacerEtRecreerWorksheet(feuilleNom)
    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Sheets(feuilleNom)
    
    wsOutput.Range("A1:Y100").Font.Name = "Aptos Narrow"
    wsOutput.Range("A1:Y100").Font.size = 10
    
    'Appel de la fonction
    Dim inclureEcritureCloture As Boolean
    inclureEcritureCloture = False
    tableau = Fn_Tableau24MoisSommeTransGL(dateCutOff, inclureEcritureCloture)
    
    With wsOutput
        .Cells(1, 1) = 0
        .Cells(2, 1) = "Compte"
        .Cells(1, 2) = 1
        .Cells(2, 2) = "Ouverture"
        For i = 1 To Len(periodes) Step 8
            .Cells(1, ((i + 7) / 8) + 2) = ((i + 7) / 8) + 1
            .Cells(2, ((i + 7) / 8) + 2) = "'" & Mid(periodes, i, 7)
        Next i
        .Cells(1, 27) = 26
        .Cells(2, 27) = "Solde"
    End With

    'Exemple d’affichage dans la fenêtre de débogage
    Dim solde As Currency
    Dim k As Long
    Dim r As Long
    r = 3
    For i = LBound(tableau, 1) To UBound(tableau, 1)
    wsOutput.Cells(r, 1) = tableau(i, 0) 'noCompteGL
        For j = 1 To 25
            wsOutput.Cells(r, j + 1) = tableau(i, j)
        Next j
        solde = 0
        If tableau(i, 0) < "4000" Then
            For k = 1 To 25
                solde = solde + tableau(i, k)
            Next k
        Else
            For k = (25 - moisAnneeFinanciere + 1) To 25
                solde = solde + tableau(i, k)
            Next k
        End If
        wsOutput.Cells(r, 27) = CCur(solde)
        r = r + 1
    Next i
    
    Dim col As Long
    r = r + 1
    wsOutput.Cells(r, 1) = "TOTAUX"
    For col = 2 To 27
        wsOutput.Cells(r, col).formula = "=SUM(" & wsOutput.Cells(3, col).Address & ":" & wsOutput.Cells(r - 2, col).Address & ")"
    Next col
    
    wsOutput.Columns("A").HorizontalAlignment = xlCenter
    wsOutput.Columns("B:AA").HorizontalAlignment = xlRight
    wsOutput.Rows("1:2").HorizontalAlignment = xlCenter
    wsOutput.Columns.AutoFit
    
End Sub

Sub zz_TestTableau24MoisMemoire() '2025-08-12 @ 19:39

    Dim tableau() As Variant
    Dim i As Long, j As Long
    
    Dim dateCutOff As Date
    dateCutOff = #7/31/2025#
    
    Dim periodes As String
    periodes = Fn_Construire24PeriodesGL(dateCutOff)
    
    'Détermine le mois de l'année financière en fonction de la date limite
    Dim dernierMoisAnneeFinanciere As Long
    dernierMoisAnneeFinanciere = wsdADMIN.Range("MoisFinAnnéeFinancière")
    Dim moisAnneeFinanciere As Long
    moisAnneeFinanciere = month(dateCutOff)
    If moisAnneeFinanciere > dernierMoisAnneeFinanciere Then
        moisAnneeFinanciere = moisAnneeFinanciere - dernierMoisAnneeFinanciere
    Else
        moisAnneeFinanciere = moisAnneeFinanciere + 12 - dernierMoisAnneeFinanciere
    End If
    
    'Appel de la fonction
    Dim inclureEcritureCloture As Boolean
    inclureEcritureCloture = False
    tableau = Fn_Tableau24MoisSommeTransGL(dateCutOff, inclureEcritureCloture)
    
    'Exemple d’affichage dans la fenêtre de débogage
    Dim soldeAC As Currency
    Dim soldeAP As Currency
    Dim k As Long
    Dim r As Long
    For i = LBound(tableau, 1) To UBound(tableau, 1)
        soldeAC = 0
        soldeAP = 0
        If tableau(i, 0) < "4000" Then
            For k = 1 To 13
                soldeAP = soldeAP + tableau(i, k)
            Next k
            For k = 1 To 25
                soldeAC = soldeAC + tableau(i, k)
            Next k
        Else
            For k = (13 - moisAnneeFinanciere + 1) To 13
                soldeAP = soldeAP + tableau(i, k)
            Next k
            For k = (25 - moisAnneeFinanciere + 1) To 25
                soldeAC = soldeAC + tableau(i, k)
            Next k
        End If
        Debug.Print tableau(i, 0), Format$(soldeAP, "#,##0.00"), Format$(soldeAC, "#,##0.00")
    Next i
    
End Sub

