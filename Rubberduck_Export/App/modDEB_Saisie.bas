Attribute VB_Name = "modDEB_Saisie"
'@IgnoreModule ValueRequired
'@Folder("Saisie_Déboursé")

Option Explicit

'Variables globales
Public gSauvegardesCaracteristiquesForme As Object
Public gNumeroDebourseARenverser As Long

Sub shpMettreAJourDEB_Click()

    Call MettreAJourDebours

End Sub

Sub MettreAJourDebours()

    If wshDEB_Saisie.Range("B7").Value = True Then
        Call MettreAJourDEBRenversement
        Exit Sub
    End If
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:MettreAJourDebours", vbNullString, 0)
    
    'Remove highlight from last cell
    If wshDEB_Saisie.Range("B4").Value <> vbNullString Then
        wshDEB_Saisie.Range(wshDEB_Saisie.Range("B4").Value).Interior.Color = xlNone
    End If
    
    'Date is not valid OR the transaction does not balance
    If Fn_DateEstElleValide(wshDEB_Saisie.Range("O4").Value) = False Or _
        Fn_SaisieDEBBalance = False Then
        Exit Sub
    End If
    
    'Is every line of the transaction well entered ?
    Dim rowDebSaisie As Long
    rowDebSaisie = wshDEB_Saisie.Range("E23").End(xlUp).Row  'Last Used Row in wshDEB_Saisie
    If Fn_SaisieDEBEstElleValide(rowDebSaisie) = False Then Exit Sub
    
    'Get the FournID
    wshDEB_Saisie.Range("B5").Value = Fn_ClientIDAPartirDuNomDeFournisseur(wshDEB_Saisie.Range("J4").Value)

    'Transfert des données vers DEB_Trans
    Call AjouterDebBDMaster(rowDebSaisie)
    Call AjouterDebBDLocale(rowDebSaisie)
    
    'GL posting
    Call ComptabiliserDebours
    
    If wshDEB_Saisie.ckbRecurrente = True Then
        Call SauvegarderDEBRecurrent(rowDebSaisie)
    End If
    
    'Retrieve the CurrentDebours number
    Dim CurrentDeboursNo As String
    CurrentDeboursNo = wshDEB_Saisie.Range("B1").Value
    
    MsgBox "Le déboursé, numéro '" & CurrentDeboursNo & "' a été reporté avec succès"
    
    'Get ready for a new one
    Call EffacerCellulesSaisieDEB
    
    Application.EnableEvents = True
    
    wshDEB_Saisie.Activate
    wshDEB_Saisie.Range("F4").Select
        
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:MettreAJourDebours", vbNullString, startTime)
        
End Sub

Sub MettreAJourDEBRenversement()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:MettreAJourDEBRenversement", vbNullString, 0)
    
    Dim ws As Worksheet
    Set ws = wshDEB_Saisie
    
    'Est-ce que la transaction balance ?
    If ws.Range("O6").Value <> ws.Range("I26").Value Then
        MsgBox "Le déboursé à renverser ne balance pas !!!", vbCritical
        Exit Sub
    End If
    
    Dim rowLastUsed As Long
    rowLastUsed = ws.Range("E24").End(xlUp).Row  'Last Used Row in wshDEB_Saisie
    If rowLastUsed < 9 Then
        Exit Sub
    End If
    
    'Get the FournID
    ws.Range("B5").Value = Fn_ClientIDAPartirDuNomDeFournisseur(wshDEB_Saisie.Range("J4").Value)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Renverser les signes des montants
    ws.Cells(6, "O").Value = -ws.Cells(6, "O").Value
    Dim i As Integer
    For i = 9 To rowLastUsed
        ws.Cells(i, 9).Value = -ws.Cells(i, 9).Value
        ws.Cells(i, 12).Value = -ws.Cells(i, 12).Value
        ws.Cells(i, 13).Value = -ws.Cells(i, 13).Value
        ws.Cells(i, 14).Value = -ws.Cells(i, 14).Value
    Next i
    
    'Transfert des données vers wsdDEB_Trans
    Call AjouterDebBDMaster(rowLastUsed)
    Call AjouterDebBDLocale(rowLastUsed)
    
    'Mettre à jour le débouré renversé
    Call MettreAJourDEBRenversementBDMaster
    Call MettreAJourDEBRenversementBDLocale
    
    'GL posting
    Call ComptabiliserDebours
    
    MsgBox "Le déboursé a été RENVERSÉ avec succès", vbInformation, "Confirmation de traitement"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    DoEvents
    
    'Reorganise wsdDEB_Trans
    Application.ScreenUpdating = False
    Dim shp As Shape
    Set shp = ws.Shapes("shpMettreAJour")
    Call RestaurerParametresForme(shp)
    
    Application.EnableEvents = False
    
    'Renverser les montants
        ws.Cells(6, "O").Value = -ws.Cells(6, "O").Value
    For i = 9 To rowLastUsed
        ws.Cells(i, 9).Value = -ws.Cells(i, 9).Value
        ws.Cells(i, 12).Value = -ws.Cells(i, 12).Value
        ws.Cells(i, 13).Value = -ws.Cells(i, 13).Value
        ws.Cells(i, 14).Value = -ws.Cells(i, 14).Value
    Next i
    
    ws.Range("F4, J4, O4, F6, M6, O6").Font.Color = vbBlack
    ws.Range("E9:O23").Font.Color = vbBlack

    'Retour à la source
    ws.Range("F4").Select
    
    DoEvents
    
    'Mode normal (pas renversement)
    gNumeroDebourseARenverser = -1
    ws.Range("B7").Value = False

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:MettreAJourDEBRenversement", vbNullString, startTime)
    
End Sub

Sub AjouterDebBDMaster(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:AjouterDebBDMaster", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "DEB_Trans$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(NoEntrée) AS MaxDebTransNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    recSet.Open strSQL, conn
    
    'Get the last used row
    Dim lastDebTrans As Long
    If IsNull(recSet.Fields("MaxDebTransNo").Value) Then
        'Handle empty table (assign a default value, e.g., 0)
        lastDebTrans = 0
    Else
        lastDebTrans = recSet.Fields("MaxDebTransNo").Value
    End If
    
    'Calculate the new Debourse Number
    Dim currDebTransNo As Long
    currDebTransNo = lastDebTrans + 1
    Application.EnableEvents = False
    wshDEB_Saisie.Range("B1").Value = currDebTransNo
    Application.EnableEvents = True
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    recSet.Close
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Read all line from Journal Entry
    Dim l As Long
    For l = 9 To r
        recSet.AddNew
            With wshDEB_Saisie
                'Add fields to the recordset before updating it
                recSet.Fields(fDebTNoEntrée - 1).Value = currDebTransNo
                recSet.Fields(fDebTDate - 1).Value = .Range("O4").Value
                recSet.Fields(fDebTType - 1).Value = .Range("F4").Value
                recSet.Fields(fDebTBeneficiaire - 1).Value = .Range("J4").Value
                recSet.Fields(fDebTFournID - 1).Value = .Range("B5").Value
                recSet.Fields(fDebTDescription - 1).Value = .Range("F6").Value & IIf(.Range("B7"), " (RENVERSEMENT de " & gNumeroDebourseARenverser & ")", vbNullString)
                recSet.Fields(fDebTReference - 1).Value = .Range("M6").Value
                
                recSet.Fields(fDebTNoCompte - 1).Value = .Range("Q" & l).Value
                recSet.Fields(fDebTCompte - 1).Value = .Range("E" & l).Value
                recSet.Fields(fDebTCodeTaxe - 1).Value = .Range("H" & l).Value
                recSet.Fields(fDebTTotal - 1).Value = CDbl(.Range("I" & l).Value)
                recSet.Fields(fDebTTPS - 1).Value = CDbl(.Range("J" & l).Value)
                recSet.Fields(fDebTTVQ - 1).Value = CDbl(.Range("K" & l).Value)
                recSet.Fields(fDebTCréditTPS - 1).Value = CDbl(.Range("L" & l).Value)
                recSet.Fields(fDebTCréditTVQ - 1).Value = CDbl(.Range("M" & l).Value)
                'Montant de dépense (Total - creditTPS - creditTVQ)
                recSet.Fields(fDebTDépense - 1).Value = CDbl(.Range("I" & l).Value _
                                                  - .Range("L" & l).Value _
                                                  - .Range("M" & l).Value)
                recSet.Fields(fDebTAutreRemarque - 1).Value = vbNullString
                recSet.Fields(fDebTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            End With
        recSet.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:AjouterDebBDMaster", vbNullString, startTime)

End Sub

Sub AjouterDebBDLocale(r As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("*** modDEB_Saisie:AjouterDebBDLocale", CStr(r), 0)
    
    Dim ws As Worksheet
    Set ws = wsdDEB_Trans
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim currentDebTransNo As Long
    currentDebTransNo = wshDEB_Saisie.Range("B1").Value
    
    'What is the last used row in DEB_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wsdDEB_Trans.Cells(wsdDEB_Trans.Rows.count, "A").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim i As Long
    For i = 9 To r
        With wshDEB_Saisie
            ws.Cells(rowToBeUsed, fDebTNoEntrée).Value = currentDebTransNo
            ws.Cells(rowToBeUsed, fDebTDate).Value = .Range("O4").Value
            ws.Cells(rowToBeUsed, fDebTType).Value = .Range("F4").Value
            ws.Cells(rowToBeUsed, fDebTBeneficiaire).Value = .Range("J4").Value
            ws.Cells(rowToBeUsed, fDebTFournID).Value = .Range("B5").Value
            ws.Cells(rowToBeUsed, fDebTDescription).Value = .Range("F6").Value & IIf(.Range("B7"), " (RENVERSEMENT de " & gNumeroDebourseARenverser & ")", vbNullString)
            ws.Cells(rowToBeUsed, fDebTReference).Value = .Range("M6").Value
            
            ws.Cells(rowToBeUsed, fDebTNoCompte).Value = .Range("Q" & i).Value
            ws.Cells(rowToBeUsed, fDebTCompte).Value = .Range("E" & i).Value
            ws.Cells(rowToBeUsed, fDebTCodeTaxe).Value = .Range("H" & i).Value
            ws.Cells(rowToBeUsed, fDebTTotal).Value = .Range("I" & i).Value
            ws.Cells(rowToBeUsed, fDebTTPS).Value = .Range("J" & i).Value
            ws.Cells(rowToBeUsed, fDebTTVQ).Value = .Range("K" & i).Value
            ws.Cells(rowToBeUsed, fDebTCréditTPS).Value = .Range("L" & i).Value
            ws.Cells(rowToBeUsed, fDebTCréditTVQ).Value = .Range("M" & i).Value
            '$ dépense = Total - creditTPS - creditTVQ
            ws.Cells(rowToBeUsed, fDebTDépense).Value = .Range("I" & i).Value _
                                                          - .Range("L" & i).Value _
                                                          - .Range("M" & i).Value
            ws.Cells(rowToBeUsed, fDebTAutreRemarque).Value = vbNullString
            ws.Cells(rowToBeUsed, fDebTTimeStamp).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        End With
        rowToBeUsed = rowToBeUsed + 1
        Call modDev_Utils.EnregistrerLogApplication("    modDEB_Saisie:AjouterDebBDLocale", "", -1)
    Next i
    
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:AjouterDebBDLocale", vbNullString, startTime)

    Application.ScreenUpdating = True

End Sub

Sub MettreAJourDEBRenversementBDMaster()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:MettreAJourDEBRenversementBDMaster", vbNullString, 0)
    
    'Définition des paramètres
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "DEB_Trans$"

    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'Requête SQL pour rechercher la ligne correspondante
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE [NoEntrée] = " & gNumeroDebourseARenverser

    'Ouvrir le Recordset
    recSet.Open strSQL, conn, 1, 3 'adOpenKeyset (1) + adLockOptimistic (3) pour modifier les données

    'Vérifier si des enregistrements existent
    If recSet.EOF Then
        MsgBox "Aucun enregistrement trouvé.", vbCritical, "Impossible de mettre à jour les déboursés RENVERSÉS"
    Else
        'Boucler à travers les enregistrements
        Do While Not recSet.EOF
        ' Vérifier si Reference contient déjà "RENVERSÉ" pour éviter les doublons
        If InStr(1, recSet.Fields(fDebTDescription - 1).Value, " (RENVERSÉ", vbTextCompare) = 0 Then
            recSet.Fields(fDebTDescription - 1).Value = recSet.Fields(fDebTDescription - 1).Value & " (RENVERSÉ par " & wshDEB_Saisie.Range("B1").Value & ")"
            recSet.Update
        End If
        'Passer à l'enregistrement suivant
        recSet.MoveNext
        Loop
    End If
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:MettreAJourDEBRenversementBDMaster", vbNullString, startTime)
    
End Sub

Sub MettreAJourDEBRenversementBDLocale()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:MettreAJourDEBRenversementBDLocale", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = wsdDEB_Trans
    
    'Dernière ligne de la table
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    'Boucler sur toutes les lignes pour trouver les correspondances
    Dim cell As Range
    For Each cell In ws.Range("A2:A" & lastUsedRow)
        If cell.Value = gNumeroDebourseARenverser Then
            'Vérifier si "RENVERSÉ" est déjà présent pour éviter les doublons
            If InStr(1, cell.offset(0, fDebTDescription - 1).Value, " (RENVERSÉ par ", vbTextCompare) = 0 Then
                'Ajouter "RENVERSÉ" à la colonne "Reference" (colonne B)
                cell.offset(0, fDebTDescription - 1).Value = cell.offset(0, fDebTDescription - 1).Value & " (RENVERSÉ par " & wshDEB_Saisie.Range("B1").Value & ")"
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set ws = Nothing

    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:MettreAJourDEBRenversementBDLocale", vbNullString, startTime)
    
End Sub

Sub AfficherDeboursRecurrent()

    ufListeDEBAuto.show

End Sub

Sub PreparerListeDEBRecurrentPourAfficher()

    'Afficher le UserForm
    ufListeDebourse.show vbModal
    
    If gNumeroDebourseARenverser = -1 Then
        wshDEB_Saisie.Range("F4").Value = vbNullString
        wshDEB_Saisie.Range("F4").Select
    Else
        wshDEB_Saisie.Range("B7").Value = True
    End If
    
End Sub

Sub ConstruireEcritureDEBRenversement() '2025-02-23 @ 16:56

    Dim ws As Worksheet: Set ws = wsdDEB_Trans
    
    '1. Quelle écriture doit-on renverser (à partir d'un ListBox)
    Call PreparerListeDEBRecurrentPourAfficher
    
    If gNumeroDebourseARenverser = -1 Then
        MsgBox "Vous n'avez sélectionné aucun déboursé à renverser", vbInformation, "Sélection d'un déboursé à renverser"
        Application.EnableEvents = True
        wshDEB_Saisie.Range("F4").Value = vbNullString
        wshDEB_Saisie.Range("F4").Select
        Application.EnableEvents = False
        GoTo Nettoyage
    End If
    
    '2. Aller chercher les debourses pour le numero choisi (0 à n lignes)
    Dim debTransSubset As Variant
    debTransSubset = Fn_RangeeAPartirNumeroColonne1(ws, gNumeroDebourseARenverser)
    
    Application.EnableEvents = False

    Dim totalDeb As Currency
    With wshDEB_Saisie
        'Entête
        .Range("F4").Value = debTransSubset(1, fDebTType)
        .Range("J4").Value = debTransSubset(1, fDebTBeneficiaire)
        .Range("O4").Value = Format$(debTransSubset(1, fDebTDate), wsdADMIN.Range("B1").Value)
        .Range("F6").Value = debTransSubset(1, fDebTDescription)
        .Range("M6").Value = debTransSubset(1, fDebTReference)
        'Détail
        Dim i As Long, r As Long
        r = 9
        Dim compteGL As String
        For i = 1 To UBound(debTransSubset, 1)
            compteGL = debTransSubset(i, fDebTCompte)
            .Range("E" & r).Value = compteGL
            .Range("H" & r).Value = debTransSubset(i, fDebTCodeTaxe)
            With .Range("I" & r)
                .Value = CCur(debTransSubset(i, fDebTTotal))
                .NumberFormat = "#,##0.00;-#,##0.00;0.00"
            End With
            With .Range("L" & r)
                .Value = CCur(debTransSubset(i, fDebTCréditTPS))
                .NumberFormat = "#,##0.00;-#,##0.00;0.00"
            End With
            With .Range("M" & r)
                .Value = CCur(debTransSubset(i, fDebTCréditTVQ))
                .NumberFormat = "#,##0.00;-#,##0.00;0.00"
            End With
            With .Range("N" & r)
                .Value = CCur(debTransSubset(i, fDebTDépense))
                .NumberFormat = "#,##0.00;-#,##0.00;0.00"
            End With
            .Range("Q" & r).Value = Fn_GetGL_Code_From_GL_Description(compteGL)
            totalDeb = totalDeb + CCur(debTransSubset(i, fDebTTotal))
            r = r + 1
        Next i
        With .Range("O6")
            .Value = CCur(totalDeb)
            .NumberFormat = "#,##0.00;-#,##0.00;0.00"
        End With
    End With
    Application.EnableEvents = True

    'On affiche le déboursé à renverser en rouge
    wshDEB_Saisie.Range("F4, J4, O4, F6, M6, O6, E9:N" & r - 1).Font.Color = vbRed

    'Change le libellé du Bouton & caractéristiques
    Dim shp As Shape
    Set shp = wshDEB_Saisie.Shapes("shpMettreAJour")
    Call ModifierForme(shp)

Nettoyage:

    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub ComptabiliserDebours() '2025-08-05 @ 11:22

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:ComptabiliserDebours", vbNullString, 0)

    Dim ws As Worksheet: Set ws = wshDEB_Saisie
    
    'Y a-t-il des lignes à traiter ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(wshDEB_Saisie.Rows.count, "E").End(xlUp).Row

    'Récupère les variables à partir de la feuille DEB_Saisie
    Dim dateDebours As Date
    dateDebours = ws.Range("O4").Value
    Dim montant As Currency
    montant = ws.Range("O6").Value
    Dim deboursType As String
    deboursType = ws.Range("F4").Value
    Dim descGL_Trans As String
    descGL_Trans = deboursType & " - " & ws.Range("F6").Value
    If Trim$(ws.Range("M6").Value) <> vbNullString Then
        descGL_Trans = descGL_Trans & " [" & ws.Range("M6").Value & "]"
    End If
    Dim source As String
    If wshDEB_Saisie.Range("B7").Value = False Then
        source = "DÉBOURSÉ:" & Format$(ws.Range("B1").Value, "00000")
    Else
        source = "RENV/DÉBOURSÉ:" & Format$(gNumeroDebourseARenverser, "00000")
    End If
    
    'Déclaration et instanciation d'un objet GL_Entry
    Dim ecr As clsGL_Entry
    Set ecr = New clsGL_Entry

    'Remplissage des propriétés communes
    ecr.DateEcriture = dateDebours
    ecr.description = descGL_Trans
    ecr.source = source
    
    Dim codeGL As String
    Dim descGL As String
    
    'La portion Crédit varie en fonction du type de déboursé
    Select Case deboursType
        Case "Chèque", "Virement", "Paiement pré-autorisé", "Autre"
            codeGL = modFunctions.Fn_NoCompteAPartirIndicateurCompte("Encaisse")
            descGL = modFunctions.Fn_DescriptionAPartirNoCompte(codeGL)
        Case "Carte de crédit"
            codeGL = modFunctions.Fn_NoCompteAPartirIndicateurCompte("Carte de crédit")
            descGL = modFunctions.Fn_DescriptionAPartirNoCompte(codeGL)
        Case "Avances avec Guillaume Charron"
            codeGL = modFunctions.Fn_NoCompteAPartirIndicateurCompte("Avances Guillaume Charron")
            descGL = modFunctions.Fn_DescriptionAPartirNoCompte(codeGL)
        Case "Avances avec 9249-3626 Québec inc."
            codeGL = modFunctions.Fn_NoCompteAPartirIndicateurCompte("Avances 9249-3626 Québec inc.")
            descGL = modFunctions.Fn_DescriptionAPartirNoCompte(codeGL)
        Case "Avances avec 9333-4829 Québec inc."
            codeGL = modFunctions.Fn_NoCompteAPartirIndicateurCompte("Avances 9333-4829 Québec inc.")
            descGL = modFunctions.Fn_DescriptionAPartirNoCompte(codeGL)
        Case Else
            codeGL = modFunctions.Fn_NoCompteAPartirIndicateurCompte("Encaisse")
            descGL = modFunctions.Fn_DescriptionAPartirNoCompte(codeGL)
    End Select
    
    'Portion CRÉDIT de l'écriture
    ecr.AjouterLigne codeGL, descGL, -montant, vbNullString
    
    Dim l As Long
    For l = 9 To lastUsedRow
        codeGL = ws.Range("Q" & l).Value
        descGL = ws.Range("E" & l).Value
        ecr.AjouterLigne codeGL, descGL, CCur(ws.Range("N" & l).Value), vbNullString
        
        If wshDEB_Saisie.Range("L" & l).Value <> 0 Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("TPS Payée")
            descGL = modFunctions.Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, CCur(ws.Range("L" & l).Value), vbNullString
        End If

        If wshDEB_Saisie.Range("M" & l).Value <> 0 Then
            codeGL = Fn_NoCompteAPartirIndicateurCompte("TVQ Payée")
            descGL = modFunctions.Fn_DescriptionAPartirNoCompte(codeGL)
            ecr.AjouterLigne codeGL, descGL, CCur(ws.Range("M" & l).Value), vbNullString
        End If
    Next l
    
    'Écriture
    Call modGL_Stuff.AjouterEcritureGLADOPlusLocale(ecr, False)
    
    'Libérer la mémoire
    Set ecr = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:ComptabiliserDebours", vbNullString, startTime)

End Sub

Sub ChargerDEBRecurrentDansSaisie(DEBAutoDesc As String, noDEBAuto As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:ChargerDEBRecurrentDansSaisie", vbNullString, 0)
    
    'On copie l'écriture automatique vers wshDEB_Saisie
    Dim rowDEBAuto As Long, rowDEB As Long
    rowDEBAuto = wsdDEB_Recurrent.Cells(wsdDEB_Recurrent.Rows.count, "C").End(xlUp).Row  'Last Row used in wshDEB_Recurrent
    
    Call EffacerCellulesSaisieDEB
    
    rowDEB = 9
    
    Application.EnableEvents = False
    Dim r As Long, totAmount As Currency, typeDEB As String
    For r = 2 To rowDEBAuto
        If wsdDEB_Recurrent.Range("A" & r).Value = noDEBAuto And wsdDEB_Recurrent.Range("F" & r).Value <> vbNullString Then
            wshDEB_Saisie.Range("E" & rowDEB).Value = wsdDEB_Recurrent.Range("G" & r).Value
            wshDEB_Saisie.Range("H" & rowDEB).Value = wsdDEB_Recurrent.Range("H" & r).Value
            wshDEB_Saisie.Range("I" & rowDEB).Value = wsdDEB_Recurrent.Range("I" & r).Value
            wshDEB_Saisie.Range("J" & rowDEB).Value = wsdDEB_Recurrent.Range("J" & r).Value
            wshDEB_Saisie.Range("K" & rowDEB).Value = wsdDEB_Recurrent.Range("K" & r).Value
            wshDEB_Saisie.Range("L" & rowDEB).Value = wsdDEB_Recurrent.Range("L" & r).Value
            wshDEB_Saisie.Range("M" & rowDEB).Value = wsdDEB_Recurrent.Range("M" & r).Value
            wshDEB_Saisie.Range("N" & rowDEB).Value = wsdDEB_Recurrent.Range("I" & r).Value _
                                                      - wsdDEB_Recurrent.Range("L" & r).Value _
                                                      - wsdDEB_Recurrent.Range("M" & r).Value
            wshDEB_Saisie.Range("Q" & rowDEB).Value = wsdDEB_Recurrent.Range("F" & r).Value
            totAmount = totAmount + wsdDEB_Recurrent.Range("I" & r).Value
            If typeDEB = vbNullString Then
                typeDEB = wsdDEB_Recurrent.Range("C" & r).Value
            End If
            rowDEB = rowDEB + 1
        End If
    Next r
    wshDEB_Saisie.Range("F4").Value = typeDEB
    wshDEB_Saisie.Range("F6").Value = "[Auto]-" & DEBAutoDesc
    wshDEB_Saisie.Range("O6").Value = Format$(totAmount, "#,##0.00")
    
    Application.EnableEvents = True
    
    wshDEB_Saisie.Range("O4").Activate
    wshDEB_Saisie.Range("O4").Select

    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:ChargerDEBRecurrentDansSaisie", vbNullString, startTime)
    
End Sub

Sub SauvegarderDEBRecurrent(ll As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:SauvegarderDEBRecurrent", vbNullString, 0)
    
    Dim rowDEBLast As Long
    rowDEBLast = wshDEB_Saisie.Cells(wshDEB_Saisie.Rows.count, "E").End(xlUp).Row  'Last Used Row in wshDEB_Saisie
    
    Call AjouterDEBRecurrentBDMaster(rowDEBLast)
    Call AjouterDEBRecurrentBDLocale(rowDEBLast)
    
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:SauvegarderDEBRecurrent", vbNullString, startTime)
    
End Sub

Sub AjouterDEBRecurrentBDMaster(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:AjouterDEBRecurrentBDMaster", vbNullString, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "DEB_Recurrent$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxDebRecNo As Long
    strSQL = "SELECT MAX(NoDebRec) AS MaxDebRecNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    recSet.Open strSQL, conn
    
    'Get the last used row
    Dim lastDR As Long, nextDRNo As Long
    If IsNull(recSet.Fields("MaxDebRecNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastDR = 0
    Else
        lastDR = recSet.Fields("MaxDebRecNo").Value
    End If
    
    'Calculate the new ID
    nextDRNo = lastDR + 1
    wshDEB_Saisie.Range("B2").Value = nextDRNo

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    recSet.Close
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        recSet.AddNew
            With wshDEB_Saisie
                'Add fields to the recordset before updating it
                recSet.Fields(fDebRNoDebRec - 1).Value = nextDRNo
                recSet.Fields(fDebRDate - 1).Value = .Range("O4").Value
                recSet.Fields(fDebRType - 1).Value = .Range("F4").Value
                recSet.Fields(fDebRBeneficiaire - 1).Value = .Range("J4").Value
                recSet.Fields(fDebRReference - 1).Value = .Range("M6").Value
                recSet.Fields(fDebRNoCompte - 1).Value = .Range("Q" & l).Value
                recSet.Fields(fDebRCompte - 1).Value = .Range("E" & l).Value
                recSet.Fields(fDebRCodeTaxe - 1).Value = .Range("H" & l).Value
                recSet.Fields(fDebRTotal - 1).Value = .Range("I" & l).Value
                recSet.Fields(fDebRTPS - 1).Value = .Range("J" & l).Value
                recSet.Fields(fDebRTVQ - 1).Value = .Range("K" & l).Value
                recSet.Fields(fDebRCréditTPS - 1).Value = .Range("L" & l).Value
                recSet.Fields(fDebRCréditTVQ - 1).Value = .Range("M" & l).Value
                recSet.Fields(fDebRTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            End With
        recSet.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:AjouterDEBRecurrentBDMaster", vbNullString, startTime)

End Sub

Sub AjouterDEBRecurrentBDLocale(r As Long) 'Write records to local file
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:AjouterDEBRecurrentBDLocale", vbNullString, 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim DEBRecNo As Long
    DEBRecNo = wshDEB_Saisie.Range("B2").Value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wsdDEB_Recurrent.Cells(wsdDEB_Recurrent.Rows.count, "C").End(xlUp).Row
    rowToBeUsed = lastUsedRow + 1
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim i As Long
    For i = 9 To r
        With wshDEB_Saisie
            wsdDEB_Recurrent.Range("A" & rowToBeUsed).Value = DEBRecNo
            wsdDEB_Recurrent.Range("B" & rowToBeUsed).Value = .Range("O4").Value
            wsdDEB_Recurrent.Range("C" & rowToBeUsed).Value = .Range("F4").Value
            wsdDEB_Recurrent.Range("D" & rowToBeUsed).Value = .Range("J4").Value
            wsdDEB_Recurrent.Range("E" & rowToBeUsed).Value = .Range("M6").Value
            
            wsdDEB_Recurrent.Range("F" & rowToBeUsed).Value = .Range("Q" & i).Value
            wsdDEB_Recurrent.Range("G" & rowToBeUsed).Value = .Range("E" & i).Value
            wsdDEB_Recurrent.Range("H" & rowToBeUsed).Value = .Range("H" & i).Value
            wsdDEB_Recurrent.Range("I" & rowToBeUsed).Value = .Range("I" & i).Value
            wsdDEB_Recurrent.Range("J" & rowToBeUsed).Value = .Range("J" & i).Value
            wsdDEB_Recurrent.Range("K" & rowToBeUsed).Value = .Range("K" & i).Value
            wsdDEB_Recurrent.Range("L" & rowToBeUsed).Value = .Range("L" & i).Value
            wsdDEB_Recurrent.Range("M" & rowToBeUsed).Value = .Range("M" & i).Value
            wsdDEB_Recurrent.Range("N" & rowToBeUsed).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        End With
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call ConstruireSommaireDEBRecurrent
    
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:AjouterDEBRecurrentBDLocale", vbNullString, startTime)
    
End Sub

Sub ModifierForme(forme As Shape)

    'Appliquer des modifications à la forme
    Application.ScreenUpdating = True
    forme.Fill.ForeColor.RGB = RGB(255, 0, 0)  ' Rouge
    forme.Line.ForeColor.RGB = RGB(255, 255, 255) ' Noir
    forme.TextFrame2.TextRange.text = "Renversement"
    forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    DoEvents
    
End Sub

Sub ConstruireSommaireDEBRecurrent()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:ConstruireSommaireDEBRecurrent", vbNullString, 0)
    
    'Build the summary at column K & L
    Dim lastUsedRow1 As Long
    lastUsedRow1 = wsdDEB_Recurrent.Cells(wsdDEB_Recurrent.Rows.count, "A").End(xlUp).Row
    
    Dim lastUsedRow2 As Long
    lastUsedRow2 = wsdDEB_Recurrent.Cells(wsdDEB_Recurrent.Rows.count, "P").End(xlUp).Row
    If lastUsedRow2 > 1 Then
        wsdDEB_Recurrent.Range("P2:S" & lastUsedRow2).ClearContents
    End If
    
    With wsdDEB_Recurrent
        Dim i As Long, k As Long, oldEntry As String
        k = 2
        For i = 2 To lastUsedRow1
            If .Range("A" & i).Value <> oldEntry Then
                .Range("P" & k).Value = .Range("A" & i).Value
                .Range("Q" & k).Value = .Range("D" & i).Value
                .Range("R" & k).Value = .Range("I" & i).Value
                .Range("S" & k).Value = .Range("B" & i).Value
                oldEntry = .Range("A" & i).Value
                k = k + 1
            End If
        Next i
    End With

    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:ConstruireSommaireDEBRecurrent", vbNullString, startTime)

End Sub

Sub EffacerCellulesSaisieDEB()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:EffacerCellulesSaisieDEB", vbNullString, 0)

    Dim ws As Worksheet
    Set ws = wshDEB_Saisie
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    With ws
        .Range("F4:H4, J4:M4, O4, F6:J6, M6, O6, E9:O23, Q9:Q29").ClearContents
        .Range("O4").Value = Format$(Date, wsdADMIN.Range("B1").Value)
        wshDEB_Saisie.ckbRecurrente = False
    End With
    
    'Toutes les cellules sont en noir (élimine le mode renversement)
    With ws.Range("F4:H4, J4:M4, O4, F6:J6, M6, O6, E9:O23").Font
        .Color = vbBlack
    End With
    
    'Toutes les cellules sont sans surbrillance (élimine le vert pâle)
    With ws.Range("F4:H4, J4:M4, O4, F6:J6, M6, O6, E9:O23").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Creer une grille dans le tableau de saisie
    Dim plages As Variant
    plages = Array("E9:O23", "L26:O26")
    Call AppliquerGrille(ws, plages)
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'Protection de la feuille, seules les cellules non-verrouillées peuvent être sélectionnées
    With wshDEB_Saisie
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modDEB_Saisie:EffacerCellulesSaisieDEB", vbNullString, startTime)

End Sub

Sub shpRetournerAuMenu_Click()

    Call RetournerAuMenu
    
End Sub

Sub RetournerAuMenu()
    
    'Rétablir la forme du bouton (Mettre à jour / Renverser)
    Dim shp As Shape
    Set shp = wshDEB_Saisie.Shapes("shpMettreAJour")
    Call RestaurerParametresForme(shp)

    Call modAppli.QuitterFeuillePourMenu(wshMenuGL, True) '2025-08-19 @ 06:52
    
End Sub

Sub CalculerTaxesEtIntrants(d As Date, _
                                  taxCode As String, _
                                  total As Currency, _
                                  gst As Currency, pst As Currency, _
                                  gstCredit As Currency, pstCredit As Currency, _
                                  netAmount As Currency)

    Dim gstRate As Double, pstRate As Double
    gstRate = Fn_Get_Tax_Rate(d, "TPS")
    pstRate = Fn_Get_Tax_Rate(d, "TVQ")
    
    If total <> 0 Then 'Calculate the amount before taxes
        'GST & PST calculation
        If taxCode = "TPS/TVQ" Or taxCode = "REP" Then
            gst = Round(total / (1 + gstRate + pstRate) * gstRate, 2)
            pst = Round(total / (1 + gstRate + pstRate) * pstRate, 2)
        Else
            gst = 0
            pst = 0
        End If
        
        'Tax credits - REP cust the credit by 50%
        If taxCode = "REP" Then
            gstCredit = Round(gst / 2, 2)
            pstCredit = Round(pst / 2, 2)
        Else
            gstCredit = gst
            pstCredit = pst
        End If
        
        If taxCode = "M" Then
            gst = 0
            gstCredit = 0
            pst = 0
            pstCredit = 0
        End If
        
        'Net amount (Expense) = Total - gstCredit - pstCredit
        netAmount = total - gstCredit - pstCredit
        Exit Sub
    End If
    
    If netAmount <> 0 Then 'Calculate the taxes from the net amount
        'gst calculation
        If taxCode = "TPS/TVQ" Or taxCode = "REP" Then
            gst = Round(netAmount * gstRate, 2)
            pst = Round(netAmount * pstRate, 2)
        Else
            gst = 0
            pst = 0
        End If
        
        If taxCode = "REP" Then
            gstCredit = Round(gst / 2, 2)
            pstCredit = Round(pst / 2, 2)
        Else
            gstCredit = gst
            pstCredit = pst
        End If
        
        If taxCode = "M" Then
            gst = 0
            gstCredit = 0
            pst = 0
            pstCredit = 0
        End If
        
        total = netAmount + gstCredit + pstCredit
        
    End If
    
End Sub

Sub SauvegarderParametresForme(forme As Shape)

    ' Vérifier si le Dictionary est déjà instancié, sinon le créer
    If gSauvegardesCaracteristiquesForme Is Nothing Then
        Set gSauvegardesCaracteristiquesForme = CreateObject("Scripting.Dictionary")
    End If

    'Sauvegarder les caractéristiques originales de la forme
    gSauvegardesCaracteristiquesForme("Left") = forme.Left
    gSauvegardesCaracteristiquesForme("Width") = forme.Width
    gSauvegardesCaracteristiquesForme("Height") = forme.Height
    gSauvegardesCaracteristiquesForme("FillColor") = forme.Fill.ForeColor.RGB
    gSauvegardesCaracteristiquesForme("LineColor") = forme.Line.ForeColor.RGB
    gSauvegardesCaracteristiquesForme("Text") = forme.TextFrame2.TextRange.text
    gSauvegardesCaracteristiquesForme("TextColor") = forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    
End Sub

Sub RestaurerParametresForme(forme As Shape)

    'Vérifiez si les caractéristiques originales sont sauvegardées
    If gSauvegardesCaracteristiquesForme Is Nothing Then
        Exit Sub
    End If

    'Restaurer les caractéristiques de la forme
    forme.Left = gSauvegardesCaracteristiquesForme("Left")
    forme.Width = gSauvegardesCaracteristiquesForme("Width")
    forme.Height = gSauvegardesCaracteristiquesForme("Height")
    forme.Fill.ForeColor.RGB = gSauvegardesCaracteristiquesForme("FillColor")
    forme.Line.ForeColor.RGB = gSauvegardesCaracteristiquesForme("LineColor")
    forme.TextFrame2.TextRange.text = gSauvegardesCaracteristiquesForme("Text")
    forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = gSauvegardesCaracteristiquesForme("TextColor")

End Sub


