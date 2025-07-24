Attribute VB_Name = "modGL_Stuff"
Option Explicit

'Structure pour une écriture comptable (données communes)
Public Type tGL_Entry '2025-06-08 @ 06:59
    DateTrans As Date
    Source As String
    noCompte As String
    AutreRemarque As String
End Type

'Structure pour une écriture comptable (données spécifiques à chaque ligne)
Public Type tGL_EntryLine '2025-06-08 @ 07:02
    noCompte As String
    description As String
    montant As Double
End Type

Public Sub GL_Get_Account_Trans_AF(glNo As String, dateDeb As Date, dateFin As Date, ByRef rResult As Range)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_Get_Account_Trans_AF", glNo & " - De " & dateDeb & " à " & dateFin, 0)

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
    
    Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_Get_Account_Trans_AF", vbNullString, startTime)

End Sub

Sub GL_Posting_To_DB(df As Date, desc As String, Source As String, arr As Variant, ByRef GLEntryNo As Long) 'Generic routine 2024-06-06 @ 07:00

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_Stuff:GL_Posting_To_DB", vbNullString, 0)

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans$"
    
    'Initialize connection, connection string, open the connection and declare rs Object
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJNo As Long
    strSQL = "SELECT MAX(NoEntrée) AS MaxEJNo FROM [" & destinationTab & "]"

    'Open recordset to find out the next JE number
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 0
    Else
        lastJE = rs.Fields("MaxEJNo").Value
    End If
    
    'Calculate the new JE number
    GLEntryNo = lastJE + 1

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim i As Long, j As Long
    'Loop through the array and post each row
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = vbNullString Then GoTo Nothing_to_Post
            rs.AddNew
                'RecordSet are ZERO base, and Enums are not, so the '-1' is mandatory !!!
                rs.Fields(fGlTNoEntrée - 1).Value = GLEntryNo
                rs.Fields(fGlTDate - 1).Value = CDate(df)
                rs.Fields(fGlTDescription - 1).Value = desc
                rs.Fields(fGlTSource - 1).Value = Source
                rs.Fields(fGlTNoCompte - 1).Value = CStr(arr(i, 1))
                rs.Fields(fGlTCompte - 1).Value = modFunctions.ObtenirDescriptionCompte(CStr(arr(i, 1)))
                If arr(i, 3) > 0 Then
                    rs.Fields(fGlTDébit - 1).Value = arr(i, 3)
                Else
                    rs.Fields(fGlTCrédit - 1).Value = -arr(i, 3)
                End If
                rs.Fields(fGlTAutreRemarque - 1).Value = arr(i, 4)
                rs.Fields(fGlTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            rs.Update
            
Nothing_to_Post:
    Next i

    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
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

    Call GL_BV_Hide_Dynamic_Shape
    
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

Sub GL_BV_EffacerZoneTransactionsDetaillees(w As Worksheet)

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
    Call GL_BV_SupprimerToutesLesFormes_shpRetour(w)

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

'@Description "Procédure pour obtenir les soldes en date de la fin d'année financière et"
'             "Effectuer l'écriture de clôture pour l'exercice"
Sub ComptabiliserEcritureCloture() '2025-07-20 @ 08:35

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
    cheminFichier = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
    Dim nomFeuilleSource As String
    nomFeuilleSource = "GL_Trans"
    Dim compteBNR As String
    compteBNR = ObtenirNoGlIndicateur("Bénéfices Non Répartis")

    'Récupération des soldes par ADO (classeur, feuille, premierGL, dernierGL, dateLimite, rejet écriture clôture)
    Set soldes = ObtenirSoldesParCompteAvecADO(cheminFichier, nomFeuilleSource, "4000", "9999", dateCloture, False)
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
    Dim cmd As Object: Set cmd = CreateObject("ADODB.Command")
    
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties='Excel 12.0 Xml;HDR=YES';"
              
    Dim cpte As Variant
    Dim montant As Currency
    Dim totalResultat As Currency
    Dim ecr As clsGL_Entry
    
    'Instanciation de l'écrituire globale
    Set ecr = New clsGL_Entry
    ecr.DateEcriture = dateCloture
    ecr.description = "Écriture de clôture annuelle"
    ecr.Source = "Clôture Annuelle"
    ecr.AutreRemarque = "Générée par l'application"
    
    'Parcours du dictionaire
    Dim descCompte As String
    For Each cpte In soldes.keys
        montant = soldes(cpte)
        If montant <> 0 Then
            'Montant inverse pour solder le compte
            descCompte = modFunctions.ObtenirDescriptionCompte(CStr(cpte))
            ecr.AjouterLigne CStr(cpte), descCompte, -montant 'Inverse pour solder
            totalResultat = totalResultat + montant
        End If
    Next cpte

    'Ligne de contrepartie pour BNR
    If totalResultat <> 0 Then
        descCompte = ObtenirDescriptionCompte(compteBNR)
        ecr.AjouterLigne CStr(compteBNR), descCompte, totalResultat
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
    
    cheminMaster = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
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

'@Description "Retourne un dictionnaire avec sommaire par noCompte & Solde"
Public Function ObtenirSoldesParCompteAvecADO(cheminFichier As String, nomFeuille As String, _
                                              noCompteGLMin As String, noCompteGLMax As String, dateCloture As Date, _
                                              inclureEcrCloture As Boolean) As Dictionary '2025-07-21 @ 12:49

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modGL_BV:ObtenirSoldesParCompteAvecADO", vbNullString, 0)
    
    Dim reqSQL As String
    Dim soldes As Object: Set soldes = CreateObject("Scripting.Dictionary")
    Dim cle As String
    Dim montant As Currency
    
    'Si un seul compte est spécifié, le MAX = MIN
    If noCompteGLMax = vbNullString Then
        noCompteGLMax = noCompteGLMin
    End If

    'Connexion ADO à un classeur fermé
    Dim conn As Object 'ADODB.Connection
    Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object 'ADODB.Recordset
    Set rs = CreateObject("ADODB.Recordset")

'    On Error GoTo ErrHandler

    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & cheminFichier & ";" & _
              "Extended Properties='Excel 12.0 Xml;HDR=YES';"

    'Requête : somme des montants pour chaque compte (>= 4000), jusqu’à la date de clôture incluse
    reqSQL = "SELECT NoCompte, SUM(IIF(Débit IS NULL, 0, Débit)) - SUM(IIF(Crédit IS NULL, 0, Crédit)) AS Solde " & _
                    "FROM [" & nomFeuille & "$] " & _
                    "WHERE NoCompte >= '" & noCompteGLMin & "' AND NoCompte <= '" & noCompteGLMax & _
                    "' AND Date <= #" & Format(dateCloture, "yyyy-mm-dd") & "#"
                    
                If Not inclureEcrCloture Then
                    reqSQL = reqSQL & " AND NOT (Date = #" & Format(dateCloture, "yyyy-mm-dd") & "# AND Source = 'Clôture annuelle')"
                End If
                
                reqSQL = reqSQL & " GROUP BY NoCompte"

    Debug.Print reqSQL
    
    rs.Open reqSQL, conn, 1, 1

    Do While Not rs.EOF
        cle = CStr(rs.Fields("NoCompte").Value)
        Debug.Print "Construction du dictionary : " & cle & " = " & Format$(rs.Fields("Solde").Value, "#,##0.00")
        montant = Nz(rs.Fields("Solde").Value)
        If Not soldes.Exists(cle) Then
            soldes.Add cle, montant
        Else
            soldes(cle) = soldes(cle) + montant
        End If
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Set ObtenirSoldesParCompteAvecADO = soldes
    GoTo Exit_Function

ErrHandler:
    MsgBox "Erreur dans ObtenirSoldesParCompteAvecADO : " & Err.description, vbCritical
    On Error Resume Next
    If Not rs Is Nothing Then If rs.state = 1 Then rs.Close
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
    Dim cn As Object
    Dim rs As Object
    Dim cheminMaster As String
    Dim nextNoEntree As Long
    Dim ts As String
    Dim i As Long
    Dim l As clsGL_EntryLine
    Dim strSQL As String

    On Error GoTo CleanUpADO

    'Chemin du classeur MASTER.xlsx
    cheminMaster = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
    
    'Ouvre connexion ADO
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminMaster & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Détermine le prochain numéro d'écriture
    Set rs = cn.Execute("SELECT MAX([NoEntrée]) AS MaxNo FROM [GL_Trans$]")
    If Not rs.EOF And Not IsNull(rs!MaxNo) Then
        nextNoEntree = rs!MaxNo + 1
    Else
        nextNoEntree = 1
    End If
    entry.NoEcriture = nextNoEntree
    rs.Close
    Set rs = Nothing

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
              "'" & Replace(entry.AutreRemarque, "'", "''") & "'," & _
              "'" & ts & "'" & _
              ")"
        cn.Execute strSQL
    Next i

    cn.Close: Set cn = Nothing

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
            .Cells(lastRow + i, 9).Value = entry.AutreRemarque
            .Cells(lastRow + i, 10).Value = ts
        End With
    Next i

    If afficherMessage Then
        MsgBox "L'écriture comptable a été complétée avec succès", vbInformation, "Écriture au Grand Livre"
    End If

CleanUpADO:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.state = 1 Then rs.Close
    Set rs = Nothing
    If Not cn Is Nothing Then If cn.state = 1 Then cn.Close
    Set cn = Nothing
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l’écriture au G/L : " & Err.description, vbCritical, "AjouterEcritureGLADOPlusLocale"
    End If
    On Error GoTo 0
    
End Sub


