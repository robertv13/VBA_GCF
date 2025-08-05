Attribute VB_Name = "modGL_Stuff"
Option Explicit

'Structure pour une écriture comptable (données communes)
Public Type tGL_Entry '2025-06-08 @ 06:59
    DateTrans As Date
    Source As String
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

Sub GL_Posting_To_DB(df As Date, desc As String, Source As String, arr As Variant, ByRef GLEntryNo As Long) 'Generic routine 2024-06-06 @ 07:00

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_Posting_To_DB", vbNullString, 0)

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "GL_Trans$"
    
    'Initialize connection, connection string, open the connection and declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(NoEntrée) AS MaxEJNo FROM [" & destinationTab & "]"

    'Open recordset to find out the next JE number
    recSet.Open strSQL, conn
    
    'Get the last used row
    Dim MaxEJNo As Long
    Dim lastJE As Long
    If IsNull(recSet.Fields("MaxEJNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 0
    Else
        lastJE = recSet.Fields("MaxEJNo").Value
    End If
    
    'Calculate the new JE number
    GLEntryNo = lastJE + 1

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    recSet.Close
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim i As Long, j As Long
    'Loop through the array and post each row
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = vbNullString Then GoTo Nothing_to_Post
            recSet.AddNew
                'RecordSet are ZERO base, and Enums are not, so the '-1' is mandatory !!!
                recSet.Fields(fGlTNoEntrée - 1).Value = GLEntryNo
                recSet.Fields(fGlTDate - 1).Value = CDate(df)
                recSet.Fields(fGlTDescription - 1).Value = desc
                recSet.Fields(fGlTSource - 1).Value = Source
                recSet.Fields(fGlTNoCompte - 1).Value = CStr(arr(i, 1))
                recSet.Fields(fGlTCompte - 1).Value = modFunctions.ObtenirDescriptionCompte(CStr(arr(i, 1)))
                If arr(i, 3) > 0 Then
                    recSet.Fields(fGlTDébit - 1).Value = arr(i, 3)
                Else
                    recSet.Fields(fGlTCrédit - 1).Value = -arr(i, 3)
                End If
                recSet.Fields(fGlTAutreRemarque - 1).Value = arr(i, 4)
                recSet.Fields(fGlTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            recSet.Update
            
Nothing_to_Post:
    Next i

    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_Posting_To_DB", vbNullString, startTime)

End Sub

Sub GL_Posting_Locally(df As Date, desc As String, Source As String, arr As Variant, ByRef GLEntryNo As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("*** modGL_Stuff:GL_Posting_Locally", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'What is the last used row in GL_Trans ?
    Dim rowToBeUsed As Long
    rowToBeUsed = wsdGL_Trans.Cells(wsdGL_Trans.Rows.count, 1).End(xlUp).Row + 1
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim i As Long, j As Long
    'Loop through the array and post each row
    With wsdGL_Trans
        For i = LBound(arr, 1) To UBound(arr, 1)
            If arr(i, 1) <> vbNullString Then
                .Range("A" & rowToBeUsed).Value = GLEntryNo
                .Range("B" & rowToBeUsed).Value = CDate(df)
                .Range("C" & rowToBeUsed).Value = desc
                .Range("D" & rowToBeUsed).Value = Source
                .Range("E" & rowToBeUsed).Value = arr(i, 1)
                .Range("F" & rowToBeUsed).Value = modFunctions.ObtenirDescriptionCompte(CStr(arr(i, 1)))
                If arr(i, 3) > 0 Then
                     .Range("G" & rowToBeUsed).Value = CDbl(arr(i, 3))
                Else
                     .Range("H" & rowToBeUsed).Value = -CDbl(arr(i, 3))
                End If
                .Range("I" & rowToBeUsed).Value = arr(i, 4)
                .Range("J" & rowToBeUsed).Value = Format$(timeStamp, "dd/mm/yyyy hh:mm:ss")
                rowToBeUsed = rowToBeUsed + 1
                Call modDev_Utils.EnregistrerLogApplication("   modGL_Stuff:GL_Posting_Locally", -1)
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_Posting_Locally", vbNullString, startTime)

End Sub

Sub GL_BV_Ajouter_Shape_Retour()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_BV_Ajouter_Shape_Retour", vbNullString, 0)
    
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
            .OnAction = "GL_BV_Effacer_Zone_Et_Shape"
        End With
    End If
    
    'Libérer la mémoire
    Set btn = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_BV_Ajouter_Shape_Retour", vbNullString, startTime)

End Sub

Sub GL_BV_Effacer_Zone_Et_Shape()
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_BV_Effacer_Zone_Et_Shape", vbNullString, 0)
    
    'Effacer la plage
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Application.EnableEvents = False
    ws.Range("L1:T" & ws.Cells(ws.Rows.count, "M").End(xlUp).Row).Offset(3, 0).Clear
    Application.EnableEvents = True

    'Supprimer les formes shpRetour
    Call GL_BV_SupprimerToutesLesFormes_shpRetour(ws)

    Call EffacerFormeDynamique
    
    'Ramener le focus à C4
    Application.EnableEvents = False
    ws.Range("D4").Select
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_BV_Effacer_Zone_Et_Shape", vbNullString, startTime)

End Sub

Sub GL_BV_EffacerZoneBV(w As Worksheet)

    Application.EnableEvents = False
    Dim lastUsedRow As Long
    lastUsedRow = w.Cells(w.Rows.count, "D").End(xlUp).Row
    If lastUsedRow >= 4 Then
        w.Range("D4:G" & lastUsedRow).Clear
    End If
    Application.EnableEvents = True

End Sub

Sub GL_BV_SupprimerToutesLesFormes_shpRetour(w As Worksheet)

    Dim shp As Shape

    For Each shp In w.Shapes
        If shp.Name = "shpRetour" Then
            shp.Delete
        End If
    Next shp
    
End Sub
    
'@Description "Retourne un dictionnaire avec sommaire par noCompte & Solde à une date donnée"
Function ObtenirSoldesParCompteAvecADO(noCompteGLMin As String, noCompteGLMax As String, dateLimite As Date, _
                                       inclureEcrCloture As Boolean) As Dictionary '2025-08-02 @ 10:04

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_BV:ObtenirSoldesParCompteAvecADO", vbNullString, 0)
    
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

    Debug.Print "ObtenirSoldesParCompteAvecADO: " & strSQL
    
    recSet.Open strSQL, conn, 1, 1

    Do While Not recSet.EOF
        cle = CStr(recSet.Fields("NoCompte").Value)
        Debug.Print "Construction du dictionary (dictSoldes): " & cle & " = " & Format$(recSet.Fields("Solde").Value, "#,##0.00")
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
    
    Set ObtenirSoldesParCompteAvecADO = dictSoldes
    
    GoTo Exit_Function

ErrHandler:
    MsgBox "Erreur dans ObtenirSoldesParCompteAvecADO : " & Err.description, vbCritical
    On Error Resume Next
    If Not recSet Is Nothing Then If recSet.state = 1 Then recSet.Close
    If Not conn Is Nothing Then If conn.state = 1 Then conn.Close
    Set ObtenirSoldesParCompteAvecADO = Nothing
    
Exit_Function:
    Call modDev_Utils.EnregistrerLogApplication("modGL_BV:ObtenirSoldesParCompteAvecADO", vbNullString, startTime)

End Function

Public Function Nz(val As Variant) As Currency '2025-07-17 @ 09:57

    If IsNull(val) Or IsEmpty(val) Then
        Nz = 0
    Else
        Nz = val
    End If
    
End Function

Function ObtenirFinExercice(dateSaisie As Date) As Date '2025-07-20 @ 08:49

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
    ObtenirFinExercice = DateSerial(anneeExercice, moisFinExercice + 1, 0)
    
End Function

Public Sub AjouterEcritureGLADOPlusLocale(entry As clsGL_Entry, Optional afficherMessage As Boolean = True) '2025-06-08 @ 09:37

    '=== BLOC 1 : Écriture dans GCF_BD_MASTER.xslx en utilisant ADO ===
    Dim cheminMaster As String
    Dim nextNoEntree As Long
    Dim ts As String
    Dim i As Long
    Dim l As clsGL_EntryLine
    Dim strSQL As String

    On Error GoTo CleanUpADO

    'Chemin du classeur MASTER.xlsx
    cheminMaster = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    
    'Ouvre connexion ADO
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminMaster & ";" & _
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
              "'" & Replace(entry.Source, "'", "''") & "'," & _
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
            .Cells(lastRow + i, 4).Value = entry.Source
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

    If afficherMessage Then
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

Function ConstruireTableau24MoisGL(dateLimite As Date, inclureEcrCloture As Boolean) As Variant '2025-08-03 @ 14:16

    Dim dicoComptes As Object
    Dim dicoMois As Object
    Dim comptes As Collection
    Dim tableau() As Variant
    Dim fichier As String
    Dim strSQL As String
    Dim dateDebutOperations As Date
    Dim dateDernierJourAnneeFinanciereCourante As Date
    Dim dateDernierJourAnneeFinancierePrecedente As Date
    Dim datePremierJourAnneeFinanciereCourante As Date
    Dim datePremierJourAnneeFinancierePrecedente As Date
    Dim dernierMoisAnneeFinanciere As Long
    Dim compteTrouve As Boolean
    
    'Chemin du classeur MASTER.xlsx
    fichier = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & wsdADMIN.Range("MASTER_FILE").Value
    
    'Établir la date du premier jour de l'année financière
    datePremierJourAnneeFinanciereCourante = PremierJourAnneeFinanciere(dateLimite)
    dateDernierJourAnneeFinanciereCourante = DernierJourAnneeFinanciere(dateLimite)
    Debug.Print "Année courante : " & datePremierJourAnneeFinanciereCourante & " à " & dateDernierJourAnneeFinanciereCourante
    
    datePremierJourAnneeFinancierePrecedente = DateAdd("yyyy", -1, datePremierJourAnneeFinanciereCourante)
    dateDernierJourAnneeFinancierePrecedente = DernierJourAnneeFinanciere(datePremierJourAnneeFinancierePrecedente)
    Debug.Print "Année précédente : " & datePremierJourAnneeFinancierePrecedente & " à " & dateDernierJourAnneeFinancierePrecedente
    
    dateDebutOperations = Format$(#7/31/2024#, "yyyy-mm-dd")
    
    'Connexion ADO
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fichier & ";" & _
              "Extended Properties='Excel 12.0 Xml;HDR=YES'"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")
    
    'Requête SQL
    strSQL = "SELECT [NoCompte], year([Date]) as Annee, month([Date]) as MoisNum, " & _
             "SUM(IIF([Débit] IS NULL, 0, [Débit]) - IIF([Crédit] IS NULL, 0, [Crédit])) AS Total " & _
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
    
    'Liste unique des comptes
    Set comptes = New Collection
    recSet.MoveFirst
    
    On Error Resume Next
    Do While Not recSet.EOF
        comptes.Add recSet("NoCompte").Value, CStr(recSet("NoCompte").Value)
        recSet.MoveNext
    Loop
    On Error GoTo 0
    
    'Établir les périodes à conserver
    Dim periodesAnneeCourante As String
    Dim periodesAnneePrecedente As String
    Dim annee As Long
    Dim mois As Long
    Dim periode As String
    Dim i As Long
    
    'Établir les périodes de l'année financière précédente
    annee = year(datePremierJourAnneeFinancierePrecedente)
    mois = month(datePremierJourAnneeFinancierePrecedente)
    For i = 1 To 12
        periode = periode & Format$(annee, "0000") & "-" & Format$(mois, "00") & " "
        mois = mois + 1
        If mois > 12 Then
            annee = annee + 1
            mois = mois - 12
        End If
    Next i

    'Établir les périodes de l'année financière précédente
    annee = year(datePremierJourAnneeFinanciereCourante)
    mois = month(datePremierJourAnneeFinanciereCourante)
    For i = 1 To 12
        periode = periode & Format$(annee, "0000") & "-" & Format$(mois, "00") & " "
        mois = mois + 1
        If mois > 12 Then
            annee = annee + 1
            mois = mois - 12
        End If
    Next i
    
    Debug.Print periode
    
    Feuil2.Cells(2, 2) = "Ouverture"
    For i = 1 To Len(periode) Step 8
        Feuil2.Cells(2, ((i + 7) / 8) + 2) = "'" & Mid(periode, i, 7)
    Next i
    
    'Tableau [nb comptes x 25]
    ReDim tableau(0 To comptes.count - 1, 0 To 26)
    
    'Remplir colonne 0 avec les comptes
    Dim j As Long
    For i = 0 To comptes.count - 1
        tableau(i, 0) = comptes(i + 1) 'Collection indexée à partir de 1
    Next

    'Remplissage des mois
    recSet.MoveFirst
    Do While Not recSet.EOF
        compteTrouve = False
        For i = 0 To comptes.count - 1
            If recSet("NoCompte").Value = tableau(i, 0) Then
'                Debug.Print recSet("NoCompte").Value, recSet("Annee").Value, recSet("MoisNum").Value, recSet("Total").Value
                annee = recSet("Annee").Value
                mois = recSet("MoisNum").Value
                j = InStr(periode, Format$(annee, "0000") & "-" & Format$(mois, "00"))
                If Not j < 1 Then
                    j = ((j + 7) / 8) + 1
                Else
                    j = 1
                End If
                tableau(i, j) = IIf(IsNull(recSet("Total")), 0, recSet("Total"))
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
    ConstruireTableau24MoisGL = tableau
    
End Function

Sub TestTableauGL() '2025-08-03 @ 14:16

    Dim tableau() As Variant
    Dim i As Long, j As Long
    
    Dim dateCutoff As Date
    dateCutoff = #7/31/2025#
    
    Feuil2.Range("A2:Y100").ClearContents

    'Appel de la fonction
    tableau = ConstruireTableau24MoisGL(dateCutoff, False)
    
    'Exemple d’affichage dans la fenêtre de débogage - @TODO
    Dim r As Long
    r = 3
    For i = LBound(tableau, 1) To UBound(tableau, 1)
    Feuil2.Cells(r, 1) = tableau(i, 0)
        For j = 1 To 25
            Feuil2.Cells(r, j + 1) = tableau(i, j)
        Next j
    
        r = r + 1
        
'        Debug.Print "Compte: " & tableau(i, 0) & _
'            " Solde ouv. (A/P) = " & Fn_Pad_A_String(Format$(tableau(i, 1), "#,##0.00"), " ", 13, "L") & _
'            " 2024/07/31 = " & Fn_Pad_A_String(Format$(tableau(i, 13), "#,##0.00"), " ", 13, "L") & _
'            " 2025/07/31 = " & Fn_Pad_A_String(Format$(tableau(i, 25), "#,##0.00"), " ", 13, "L") & _
'            " SOLDE = " & Fn_Pad_A_String(Format$(solde, "#,##0.00"), " ", 13, "L")
    Next
    
End Sub

